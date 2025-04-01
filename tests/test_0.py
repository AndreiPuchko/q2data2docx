import os
import time
import q2data2docx.q2data2docx as q2data2docx

test_data_folder = "test-data/test01/"


def test_manual_merge():
    test_input_docx_filename = f"{test_data_folder}/test.docx"
    test_input_xlsx_filename = f"{test_data_folder}/test.xlsx"
    test_result_file_name = "test-result/test-result.docx"

    d2d = q2data2docx.q2data2docx()

    d2d.loadDocxFile(test_input_docx_filename)
    d2d.loadXlsxFile(test_input_xlsx_filename)
    d2d.merge()

    result_name = d2d.saveFile(test_result_file_name, open_output_file=True)
    assert result_name
    assert os.path.exists(result_name)

if __name__ == "__main__":
    t = time.time()
    test_manual_merge()
    print(time.time() - t)
