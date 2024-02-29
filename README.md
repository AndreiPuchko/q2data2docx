http://d2d.penta.by/#Tutorial
# Usage
## import
```python
import q2data2docx.q2data2docx
```

## merge local files
```python
d2d = q2data2docx.q2data2docx.q2data2docx()
d2d.loadJsonFile("your-file-name.json")  # JSON data file
# or
d2d.loadJsonFile("your-file-name.xlsx")  # XLSX data file

d2d.loadDocxFile("your-file-name.docx")  # DOCX template file
if d2d.merge():
    d2d.saveFile("result")

if d2d.merge():
    d2d.saveFile("result")
```
*or*
```python
q2data2docx.q2data2docx.merge("your-file-name.docx", "your-file-name.json", "result")
# or
q2data2docx.q2data2docx.merge("your-file-name.docx", "your-file-name.xlsx", "result")
```


## merge data from memory
```python
d2d = q2data2docx.q2data2docx.q2data2docx()

d2d.setJsonBinary(open("your-file-name.json", "rb").read())
# or
d2d.setXlsxBinary(open("your-file-name.xlsx", "rb").read())

d2d.setDocxTemplateBinary(open("your-file-name.docx", "rb").read())
if d2d.merge():
    d2d.saveFile("result")
```
*or*
```python
d2d = q2data2docx.q2data2docx.q2data2docx(
    docxTemplateBinary=open("your-file-name.docx", "rb").read(),
    jsonBinary=open("your-file-name.json", "rb").read(),
)
# or 
d2d = q2data2docx.q2data2docx.q2data2docx(
    docxTemplateBinary=open("your-file-name.docx", "rb").read(),
    xlsxBinary=open("your-file-name.xlsx", "rb").read(),
)
if d2d.merge():
    d2d.saveFile("result")
```
