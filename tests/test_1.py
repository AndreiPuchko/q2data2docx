import os

import q2data2docx.q2data2docx as q2data2docx

test_data_folder = "test-data/"


def get_test_set():
    for folder in os.listdir(test_data_folder):
        test_input_docx_filename = f"{test_data_folder}/{folder}/test.docx"
        test_input_xlsx_filename = f"{test_data_folder}/{folder}/test.xlsx"
        test_result_file_name = f"test-result/test-result_{folder}.docx"
        yield test_input_docx_filename, test_input_xlsx_filename, test_result_file_name


def test_merge():
    for test_input_docx_filename, test_input_xlsx_filename, test_result_file_name in get_test_set():
        result = q2data2docx.merge(test_input_docx_filename, test_input_xlsx_filename, test_result_file_name)
        assert result is True


def test_class():
    d2d = q2data2docx.q2data2docx()
    for test_input_docx_filename, test_input_xlsx_filename, test_result_file_name in get_test_set():
        assert d2d.loadDocxFile(test_input_docx_filename)
        assert d2d.loadXlsxFile(test_input_xlsx_filename)
        assert d2d.merge()
        assert d2d.saveFile(test_result_file_name)


if __name__ == "__main__":
    test_merge()
    test_class()
