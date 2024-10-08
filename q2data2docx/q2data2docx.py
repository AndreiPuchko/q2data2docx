#    Copyright © 2024 Andrei Puchko
#
#    Licensed under the Apache License, Version 2.0 (the "License");
#    you may not use this file except in compliance with the License.
#    You may obtain a copy of the License at
#
#        http://www.apache.org/licenses/LICENSE-2.0
#
#    Unless required by applicable law or agreed to in writing, software
#    distributed under the License is distributed on an "AS IS" BASIS,
#    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#    See the License for the specific language governing permissions and
#    limitations under the License.

import glob
import re
import json
import html
import subprocess
from io import BytesIO
import os
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED
from decimal import Decimal
from datetime import datetime

BUILTIN_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: r'_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
    42: r'_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
    43: r'_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',
    44: r'_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}


def N(t):
    try:
        return Decimal(f"{t}")
    except Exception:
        return 0


class rowEvaluator:
    def __init__(self):
        pass

    def setData(self, row, proxy):
        self.row = row.copy()
        self.row["N"] = N
        self.proxy = {proxy[x]: x for x in proxy}
        self.proxy["N"] = "N"

    def __getitem__(self, key):
        return self.row.get(self.proxy.get(key, key), "")


def excelDataFormat(cellText, formatStr):
    formatStr = formatStr.replace("yyyy", "%Y").replace("/", r".")
    formatStr = formatStr.replace("mmm", "%b").replace("mm", "%m")
    formatStr = formatStr.replace("ddd", "%d %A").replace("dd", "%d")
    if "%d" not in formatStr:
        formatStr = formatStr.replace("d", "%d")
    if "%m" not in formatStr:
        formatStr = formatStr.replace("m", "%m")

    formatStr = formatStr.replace('"', "")

    cellText = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(N(cellText)) - 2).strftime(
        formatStr
    )
    return cellText


def merge(file1, file2, outputFile):
    d2d = q2data2docx()
    for x in [file1, file2]:
        if not d2d.loadFile(x):
            print(f"Input file not found: {x}")
            return False
    if d2d.merge():
        outputFile = d2d.checkOutputFileName(outputFile)
        if d2d.saveFile(outputFile):
            print(f"Processing finished. See result in {outputFile}")
            return True
        else:
            print(f"Can't create output file: {outputFile}")
    else:
        print(f"Processing error: {d2d.error}")
    return False


class q2data2docx:
    def __init__(self, dataDic={}, docxTemplateBinary="", xlsxBinary="", jsonBinary=""):

        self.error = ""
        self.xlsxBinary = None
        self.jsonBinary = None
        self.docxTemplateBinary = None
        self.dataDic = dataDic

        self.usedFormatStrings = {}

        self.xlsxRowsCount = 0

        self.docxSizeLimit = 0
        self.xlsxSizeLimit = 0
        self.xlsxRowLimit = 0
        self.xlsxSheetLimit = 0

        self.setJsonBinary(jsonBinary)
        self.setDocxTemplateBinary(docxTemplateBinary)
        self.setXlsxBinary(xlsxBinary)

        self.docxResultBinary = ""

    def loadFile(self, fileName):
        return {
            ".json": self.loadJsonFile,
            ".docx": self.loadDocxFile,
            ".xlsx": self.loadXlsxFile,
        }.get(
            os.path.splitext(fileName)[1].lower(), lambda x: None
        )(fileName)

    # prepare excel data
    def loadXlsxFile(self, fileIn):
        if glob.glob(fileIn):
            binXls = None
            try:
                binXls = open(fileIn, "rb").read()
            except Exception:
                return False
            if binXls:
                self.setXlsxBinary(binXls)
            return True
        else:
            return False

    def setXlsxBinary(self, xlsxBinary=None):
        self.jsonBinary = ""
        if xlsxBinary:
            self.xlsxBinary = xlsxBinary
            self.xlsxBinary2dataDic()

    def xlsxBinary2dataDic(self):
        if not self.xlsxBinary:
            return
        self.xlsxRowsCount = 0

        memzip = BytesIO()
        memzip.write(self.xlsxBinary)
        xlsxZip = ZipFile(memzip)

        sheetNames = self.extractSheetNames(xlsxZip)
        cellXfs, numFmts = self.extractFormats(xlsxZip)
        sharedStrings = self.extractSharedString(xlsxZip)
        sheetIdDic = self.extractSheets(xlsxZip)

        for x in xlsxZip.namelist():
            if not x.startswith("xl/worksheets/s"):
                continue
            sheetName = sheetIdDic[sheetNames[os.path.basename(x)]]
            if not sheetName:
                return
            sheetText = xlsxZip.open(x).read().decode("utf-8")
            # remove empty cells
            sheetText = "".join(
                [
                    y
                    for y in re.split(r"(<[^<]+>)", sheetText)
                    if y != "" and not (y.startswith("<c") and y.endswith("/>"))
                ]
            )
            for child in ET.fromstring(sheetText):
                if not child.tag.endswith("sheetData"):
                    continue
                for row in child:
                    rowNumber = int(N(row.attrib["r"]))
                    sheetRow = {}
                    for cell in row:
                        colLetter = re.sub(r"\d", "", cell.attrib["r"])
                        for st in cell:
                            if not st.tag.endswith("}v"):
                                continue
                            if cell.attrib.get("t", "") in ["s"]:  # in sharedStrings
                                sheetRow[colLetter] = sharedStrings[int(st.text)]
                            else:
                                if st.text:  # if any text
                                    sheetRow[colLetter] = st.text
                                    formatStr = numFmts.get(cellXfs.get(cell.attrib.get("s"), ""), "")
                                    sheetRow[colLetter] = self.setNmFmt(sheetRow[colLetter], formatStr)
                        if sheetRow[colLetter] == "":
                            del sheetRow[colLetter]
                    if sheetRow != {}:
                        self.check4char(sheetRow)
                        self.dataDic[sheetName][rowNumber] = sheetRow
                        self.xlsxRowsCount += 1

    def extractSheetNames(self, xlsxZip):
        sheetNames = {}
        for child in ET.fromstring(xlsxZip.open("xl/_rels/workbook.xml.rels").read()):
            if child.attrib["Type"].endswith("/worksheet"):
                sheetNames[os.path.basename(child.attrib["Target"])] = child.attrib["Id"]
        return sheetNames

    def extractFormats(self, xlsxZip):
        cellXfs = {}
        numFmts = {}
        for child in ET.fromstring(xlsxZip.open("xl/styles.xml").read()):
            if child.tag.endswith("cellXfs"):
                fCo = 0
                for xf in child:
                    cellXfs[f"{fCo}"] = xf.attrib.get("numFmtId", "")
                    fCo += 1
            if child.tag.endswith("numFmts"):
                for numFmt in child:
                    numFmts[numFmt.attrib.get("numFmtId", "")] = numFmt.attrib.get("formatCode", "")
        numFmts.update(BUILTIN_FORMATS)
        return cellXfs, numFmts

    def extractSharedString(self, xlsxZip):
        sharedStrings = []
        for child in ET.fromstring(xlsxZip.open("xl/sharedStrings.xml").read()):
            for si in child:
                sharedStrings.append(si.text)
        return sharedStrings

    def extractSheets(self, xlsxZip):
        self.dataDic = {}
        sheetIdDic = {}
        for child in ET.fromstring(xlsxZip.open("xl/workbook.xml").read()):
            if child.tag.endswith("sheets"):
                for sheet in child:
                    for rid in sheet.attrib:
                        if rid.endswith("}id"):
                            sheetIdDic[sheet.attrib[rid]] = sheet.attrib["name"]
                    self.dataDic[sheet.attrib["name"]] = {}
        return sheetIdDic

    def check4char(self, row):
        for x in row:
            if "&" in row[x]:
                row[x] = html.escape(row[x])

    def setNmFmt(self, cellText, formatStr):
        if formatStr not in self.usedFormatStrings:
            self.usedFormatStrings[formatStr] = cellText
        if re.match(r"0\.0*E\+00", formatStr):  # scientific
            ln = len(re.sub(r"0\.|E\+00", "", formatStr))
            cellText = ("{:.%sE}" % ln).format(N(cellText))
        elif "m" in formatStr and "y" in formatStr and "y" in formatStr:
            cellText = excelDataFormat(cellText, formatStr)
        elif re.match(r"m/d/yyyy", formatStr):
            cellText = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(N(cellText)) - 2).strftime(
                "%m/%d/%Y"
            )
        elif re.match(r"0\.0*", formatStr):
            cellText = ("{:.%sf}" % len(formatStr.split(".")[1])).format(N(cellText))
        return cellText

    # prepare json data
    def loadJsonFile(self, fileIn):
        if glob.glob(fileIn):
            binJson = None
            try:
                binJson = open(fileIn).read()
            except Exception:
                return False
            if binJson:
                self.setJsonBinary(binJson)
            return True
        else:
            return False

    def setJsonBinary(self, jsonBinary=None):
        self.xlsxBinary = ""
        if jsonBinary:
            self.jsonBinary = jsonBinary
            self.jsonBinary2dataDic()

    def jsonBinary2dataDic(self):
        self.dataDic = json.loads(self.jsonBinary)
        if isinstance(self.dataDic, dict):
            for x in self.dataDic:
                if isinstance(self.dataDic[x], list):
                    self.dataDic[x] = {y: self.dataDic[x][y] for y in range(len(self.dataDic[x]))}
                elif isinstance(self.dataDic[x], dict):
                    self.dataDic[x] = {int(y): self.dataDic[x][y] for y in self.dataDic[x]}
        else:
            self.dataDic = {}
        self.prepareData(self.dataDic)

    def prepareData(self, dataDic):
        pass

    def setDataDic(self, dataDic):
        self.dataDic = dataDic

    # prepare template
    def loadDocxFile(self, fileIn=""):
        if glob.glob(fileIn):
            try:
                self.setDocxTemplateBinary(open(fileIn, "rb").read())
            except Exception:
                return False
            return True
        else:
            return False

    def setDocxTemplateBinary(self, docxTemplateBinary=None):
        if docxTemplateBinary:
            self.docxTemplateBinary = docxTemplateBinary

    def processFile(self, fileIn="", fileOut=""):
        if glob.glob(fileIn):
            if not fileOut:
                fileOut = fileIn.replace(".docx", "_out.docx")
            self.setDocxTemplateBinary(open(fileIn, "rb").read())
            if self.merge():
                open(fileOut, "wb").write(self.docxResultBinary)

    def cleanPar(self, parList):  # clean paragraph dummy tags
        rez = []
        dxTmp = ""
        for x in parList:
            co = x.count("#")
            if x.startswith("<"):
                if not dxTmp:
                    rez.append(x)
            elif dxTmp:
                dxTmp += x
                if dxTmp.count("#") % 2 == 0:
                    rez.append(dxTmp)
                    dxTmp = ""
            elif co % 2 == 0:
                rez.append(x)
            elif co:
                dxTmp += x
            else:
                rez.append(x)
        return "".join(rez)

    def prepareDocxTemplate(self):
        if self.docxTemplateBinary is None:
            self.error = "Template not found or not loaded"
            return False
        memzip = BytesIO()
        memzip.write(self.docxTemplateBinary)
        docxZip = ZipFile(memzip)
        if "word/document.xml" not in docxZip.namelist():
            return None
        dxDoc = docxZip.open("word/document.xml").read().decode("utf-8").replace("\n", "").replace("\t", "")
        dxDocList = []
        dxBinary = []
        parList = []
        # remove intag formats
        for x in [x for x in re.split(r"(<[^<]+>)", dxDoc) if x != ""]:
            if x == "<w:p>" or x.startswith("<w:p ") or parList:  # collecting paragaph
                parList.append(x)
                if x == "</w:p>":  # end of paragraph
                    parXml = "".join(parList)
                    dCount = parXml.count("#")
                    if dCount == 0 or dCount % 2 == 1:  # no or bad template paragraph
                        dxBinary.append(parXml)
                        dxDocList.append(f"<@{(len(dxBinary) - 1)}@>")
                    else:  # template-prepare paragraph
                        dxDocList.append(self.cleanPar(parList))
                    parList = []
            else:
                if "#" in x:
                    dxBinary.append(x)
                    x = f"<@{(len(dxBinary) - 1)}@>"
                dxDocList.append(x)
        return "".join(dxDocList), dxBinary, docxZip

    def merge(self):
        dxDoc, dxBinary, docxZip = self.prepareDocxTemplate()
        if not dxDoc:
            return
        tableTags2clean = []
        # обработка таблиц
        for tableName in self.dataDic:
            if not isinstance(self.dataDic[tableName], dict):
                continue
            if f"#{tableName}#" not in dxDoc:
                continue
            # process snippets
            docxRowXml = self.getSnippetRow(dxDoc, tableName)
            docxRows = []
            for y in range(len(docxRowXml)):
                tableTags2clean.append(f"#{tableName}{docxRowXml[y]['tableProps']}#")
                tableTags2clean.append(f"#{tableName}#")
                docxRows.append([])
                # process datatable rows
                (
                    startRow,
                    columnNamesRow,
                    endRow,
                    filterRow,
                    rowEval,
                    columnNamesProxy,
                ) = self.getTableParams(docxRowXml[y], tableName)
                for rowCount in range(1, max(self.dataDic[tableName]) + 1):
                    row = self.dataDic[tableName].get(rowCount)
                    if not row:
                        continue
                    if startRow and rowCount < startRow or rowCount == columnNamesRow:
                        continue
                    if endRow and rowCount > endRow:
                        break
                    if filterRow:
                        rowEval.setData(row, columnNamesProxy)
                        if not eval(filterRow, {}, rowEval):
                            continue
                    tmpDocxXml = docxRowXml[y]["snippet"]
                    # process datatable column:  x:column name
                    for columnName in row:
                        row[columnName] = html.escape(row[columnName])
                        tmpDocxXml = tmpDocxXml.replace(
                            f"#{columnNamesProxy.get(columnName, columnName)}#",
                            row[columnName],
                        )
                    docxRows[y].append(tmpDocxXml)
            for z in range(len(docxRowXml)):
                dxDoc = dxDoc.replace(docxRowXml[z]["snippet"], "".join(docxRows[z]))

        # processing non table data
        for dataKey, dataValue in self.dataDic.items():
            if dataKey == "":
                continue
            if not isinstance(self.dataDic[dataKey], dict):
                dxDoc = dxDoc.replace(f"#{dataKey}#" , dataValue)
        # process first sheet as non table data
        first_sheet = self.dataDic[list(self.dataDic.keys())[0]]
        if isinstance(first_sheet, dict):
            for cell_key, cell_value in {
                f"{key}{row_key}": value for row_key in first_sheet for key, value in first_sheet[row_key].items()
            }.items():
                if not isinstance(cell_value, dict):
                    dxDoc = dxDoc.replace(f"#{cell_key}#", cell_value)
        # remove datatables tags first
        # replace table names to #@#
        for x in tableTags2clean:
            dxDoc = dxDoc.replace(x, "#@#")
        # remove #@# gaps
        startPoint = 0
        while "#@#" in dxDoc[startPoint:]:
            startPoint = parEnd = parStart = dxDoc.index("#@#", startPoint)
            while parStart > 0:
                parStart = dxDoc.rfind("<w:p", 0, parStart)
                if dxDoc[parStart : parStart + 5] in ["<w:p ", "<w:p>"]:
                    break
                if parStart == 0:
                    break
            parEnd = dxDoc.find("</w:p>", parStart)
            if parStart > 0 and parEnd > 0:
                startPoint = parStart
                parXml = dxDoc[parStart : parEnd + 6]
                xmlns = "".join(
                    [
                        f"""xmlns:{x}="{x}{x}" """
                        for x in {
                            x.translate("".maketrans("", "", "</: "))
                            for x in re.findall(r"\W(\w*):", parXml)
                            if x != "xml"
                        }
                    ]
                )
                gg = ET.fromstring(f"""<r {xmlns}>{parXml}</r>""")
                parText = (
                    ("".join([x.text for x in gg.findall("./{ww}p/{ww}r/{ww}t")])).replace("#@#", "").strip()
                )
                if parText == "":
                    dxDoc = dxDoc.replace(parXml, "")
                    startPoint = parStart + 1
                else:
                    startPoint = parEnd
            startPoint += 1
        # replace absolute links to table.colLetterRowNumber
        for y in dxDoc.split("</w:t>"):
            y = y.split("<w:t")[-1].split(">")[-1]
            for x in re.findall(r"#\s*(\w*)\.([a-zA-Z]*)(\d*)\s*#", y):
                colRowData = self.dataDic.get(x[0], {}).get(int(N(x[2])), {}).get(x[1], "")
                if colRowData:
                    dxDoc = dxDoc.replace("#{}.{}{}#".format(x[0], x[1], x[2]), colRowData)
        # remove other unused tags
        dxDoc = re.sub(r"#(\s*[^#]+\s*)#", "", dxDoc)
        if "</w:tbl><w:sectPr>" in dxDoc:  # fix last row at the end of docs
            dxDoc = dxDoc.replace("</w:tbl><w:sectPr>", """</w:tbl><w:p></w:p><w:sectPr>""")
        if "</w:tbl></w:tc>" in dxDoc:  # fix last row at the end of docs
            dxDoc = dxDoc.replace("</w:tbl></w:tc>", """</w:tbl><w:p></w:p></w:tc>""")
        dxDoc = dxDoc.replace("\n", "")
        # put back binary data
        for x in range(len(dxBinary)):
            dxDoc = dxDoc.replace(f"<@{x}@>", dxBinary[x])
        # create result file as binary
        outmemzip = BytesIO()
        newZip = ZipFile(outmemzip, "w", ZIP_DEFLATED)
        for x in docxZip.namelist():
            docXml = docxZip.open(x).read()
            if x == "word/document.xml":
                docXml = dxDoc.encode("utf8")
            newZip.writestr(x, docXml)
        newZip.close()
        self.docxResultBinary = outmemzip.getvalue()
        return True

    def getSnippetRow(self, xml, tableName):
        tag = f"#{tableName}"
        rez = []
        while tag in xml:
            snippet = ""
            tag1_pos = xml.index(tag)
            tableProps = xml[tag1_pos + 1 : xml[tag1_pos + 1 :].index("#") + tag1_pos + 1].replace(
                tableName, ""
            )
            if tag in xml[tag1_pos + 1 :]:
                tag2_pos = xml.index(tag, len(tag) + tag1_pos) + len(tag) + 1
                snippet = xml[tag1_pos:tag2_pos]
                if "<w:tbl>" in snippet and "</w:tbl>" not in snippet:
                    snippet = ""
            if snippet:
                rez.append({"snippet": snippet, "tableProps": tableProps})
                tag1_pos = tag2_pos
            xml = xml[(tag1_pos + 1) :]
        return rez

    def getTableParams(self, tableSnippet, tableName):
        tmpTPList = (tableSnippet["tableProps"][1:]).split(":")
        tmpTPList = tmpTPList + ([""] * (4 - len(tmpTPList)) if len(tmpTPList) < 4 else [])
        (columnNamesRow, startRow, endRow) = [int(x) if x.isdigit() else 0 for x in tmpTPList[:3]]
        if self.jsonBinary:
            columnNamesRow = 0
        filterRow = tmpTPList[3]
        if columnNamesRow != 0:
            columnNamesProxy = self.dataDic[tableName][columnNamesRow]
        else:
            columnNamesProxy = {}
        rowEval = None
        if filterRow:
            filterRow = html.unescape(filterRow)
            try:
                filterRow = re.sub(r"([^=!\<\>])(=)([^=])", r"\1\2\2\3", filterRow)
                filterRow = compile(filterRow, "", "eval")
                rowEval = rowEvaluator()
            except Exception:
                filterRow = None
        return startRow, columnNamesRow, endRow, filterRow, rowEval, columnNamesProxy

    # save resust to file
    def checkOutputFileName(self, fileName=""):
        if not fileName.lower().endswith(".docx"):
            fileName += ".docx"

        if not os.path.isdir(os.path.dirname(fileName)):
            os.mkdir(os.path.dirname(fileName))
        co = 0
        name, ext = os.path.splitext(fileName)
        while True:
            if os.path.isfile(fileName):
                try:
                    os.remove(fileName)
                except Exception as e:
                    co += 1
                    fileName = f"{name}{co:03d}{ext}"
                    continue
            break
            # lockfile = f"{os.path.dirname(fileName)}/.~lock.{os.path.basename(fileName)}#"
            # if os.path.isfile(lockfile):
            #     co += 1
            #     fileName = f"{name}{co:03d}{ext}"
            # else:
            #     break
        return fileName

    def saveFile(self, fileOut="", open_output_file=True):
        fileOut = self.checkOutputFileName(fileOut)
        try:
            open(fileOut, "wb").write(self.docxResultBinary)
        except Exception:
            return False
        if open_output_file:
            q2data2docx.open_document(fileOut)
        return True

    @staticmethod
    def open_document(file_name):
        if os.path.isfile(file_name):
            if sys.platform == "win32":
                subprocess.Popen(
                    ["start", os.path.abspath(file_name)],
                    close_fds=True,
                    shell=True,
                    creationflags=subprocess.DETACHED_PROCESS,
                )
            else:
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([opener, file_name], close_fds=True, shell=False)

    def setTestData(self):
        self.dataDic = {
            "header": "Headet Text Example",
            "data": [
                {"id": "1", "name": "john", "address": "new york"},
                {"id": "2", "name": "piter", "address": "paris"},
                {"id": "3", "name": "alex", "address": "berlin"},
            ],
        }


if __name__ == "__main__":
    testSourceFolder = f"{os.path.dirname(__file__)}/../test-data/test01/"
    testResultFolder = f"{os.path.dirname(__file__)}/../test-result"
    result_file = os.path.abspath(f"{testResultFolder}/test-result.docx")
    if merge(
        f"{testSourceFolder}test.docx",
        f"{testSourceFolder}test.xlsx",
        result_file,
    ):
        os.system(f"explorer {result_file}")
        print("Ok")
