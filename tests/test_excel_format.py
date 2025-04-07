from q2data2docx.excel_format import format_number
import importlib

def test_1():
    res = [
        (format_number("1234.59", "#") == "1235"),  # 1
        (format_number("1234.59", "####.#") == "1234.6"),  # 1
        (format_number("8.9", "#.000") == "8.900"),  # 2
        (format_number("0.631", "0.#") == "0.6"),  # 3
        (format_number("12", "#.0#") == "12.0"),  # 4
        (format_number("1234.568", "#.0#") == "1234.57"),  # 5
        (format_number("44.398", "#.0#") == "44.4"),  # 6
        (format_number("44.398", "????.???") == "  44.398"),  # 7
        (format_number("102.65", "???.???") == "102.65 "),  # 8
        (format_number("2.65", "??.???") == " 2.65 "),  # 9
        (format_number("2.8", "???.???") == "  2.8  "),  # 10
        (format_number("1234.59", "####.#") == "1234.6"),  # 11
        (format_number("8.9", "#.000") == "8.900"),  # 12
        (format_number(".631", "0.#") == "0.6"),  # 13
        (format_number("1234.568", "#.0#") == "1234.57"),  # 14
        (format_number("5.25", "# ???/???") == "5   1/4  "),  # 15
        (format_number("5.3", "# ???/???") == "5   3/10 "),  # 16
        (format_number("5.3", "# ?/?") == "5 3/10"),  # 17
        (format_number("12000", "#,###") == "12,000"),  # 18
        (format_number("12000", "#,") == "12"),  # 19
        (format_number("12200000", "0.0,,") == "12.2"),  # 20
        (format_number("-1234", "#,##0;[Red](#,##0)")) == "(1,234)",  # 21
        (format_number("1234", "#,##0;[Red](#,##0)")) == "1,234",  # 22
        (format_number("0", "#,##0;[Red](#,##0);Zero")) == "Zero",  # 23
        (format_number("0.146", "0.####%")) == "14.6%",  # 24
        (format_number("4566", "$0.####")) == "$4566",  # 25
        (format_number("43221", "yy-m")) == "18-5",  # 25
        (format_number("43221", "yyyy-m")) == "2018-5",  # 25
        (format_number("43221", "yyyy-mmm")) == "2018-May",  # 25
        (format_number("43221", "yyyy-mmm-d")) == "2018-May-1",  # 25
        (format_number("43221", "yyyy-mmm-dd")) == "2018-May-01",  # 25
        (format_number("43221", "yyyy-mmm-ddd")) == "2018-May-Tue",  # 25
        (format_number("43221", "yyyy mmm  dddd")) == "2018 May  Tuesday",  # 25
        (format_number("43221.456", "YYYY-MM-DD HH:MM:SS.000")) == "2018-05-01 10:56:38.000",  # 25
    ]

    assert sum([1 if x else 0 for x in res]) == len(res)

if __name__ == "__main__":
    for x in [x for x in glob.glob("*.py") if x != "test.py"]:
        mod = importlib.import_module(x.removesuffix(".py"))
        format_number = getattr(mod, "format_number")
        print(x.ljust(40), test_1(), "> ", sum(test_1()))

        print("----" * 10)
        [print(f"{index+1:2}", ">>", x) for index, x in enumerate(test_1())]
        print("> ", sum(test_1()), "from", len(test_1()))
