import sys
import os

import q2data2docx.q2data2docx as q2data2docx


test_result_file_name = "test-result/test-result.docx"
test_data_folder = "test-data/test01/"

# if not os.path.exists("test-data"):
#     test_result_file_name = f"../{test_result_file_name}"
#     test_data_folder = f"../{test_data_folder}"

test_input_docx_filename = f"{test_data_folder}test.docx"
test_input_xlsx_filename = f"{test_data_folder}test.xlsx"


def test_merge():

    result = q2data2docx.merge(test_input_docx_filename, test_input_xlsx_filename, test_result_file_name)

    assert result is True


def test_class():
    d2d = q2data2docx.q2data2docx()
    assert d2d.loadDocxFile(test_input_docx_filename)
    assert d2d.loadXlsxFile(test_input_xlsx_filename)
    assert d2d.merge()
    assert d2d.saveFile(test_result_file_name)


if __name__ == "__main__":
    if os.path.isfile(test_result_file_name):
        os.remove(test_result_file_name)

    # test_merge()
    test_class()

    if os.path.isfile(test_result_file_name):
        os.system(os.path.abspath(test_result_file_name))
