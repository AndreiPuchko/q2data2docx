# import math
from fractions import Fraction
from datetime import datetime, timedelta
import re


RE_COLOR_TAGS = re.compile(r"\[[^\]]+\]")
RE_DATAFORMAT = re.compile(r"am/pm|[ymdqhs]+|[^ymdqhs]", re.IGNORECASE)
RE_ZERO = re.compile(r"[^#0?,./]+")


def select_format_section(number_str: float, format_str: str) -> str:
    sections = format_str.split(";")
    if len(sections) == 1:
        return sections[0]
    elif len(sections) == 2:
        return sections[0] if number_str >= 0 else sections[1]
    elif len(sections) >= 3:
        if number_str > 0:
            return sections[0]
        elif number_str < 0:
            return sections[1]
        else:
            return sections[2]


def strip_color_tags(fmt: str) -> str:
    return RE_COLOR_TAGS.sub("", fmt)


def rstrip_limit(s, char, limit):
    count = 0
    while s and s[-1] == char and count < limit:
        s = s[:-1]
        count += 1
    return s


def format_data(number_str: str, format_str: str) -> str:
    data_format_chars = ["y", "m", "d", "q", "s", "h", "am", "pm"]

    is_data_format = False
    for x in data_format_chars:
        if x in format_str.lower():
            is_data_format = True
            break
    if not is_data_format:
        return ""
    format_map = {
        "yyyy": "%Y",
        "yy": "%y",
        "mmmm": "%B",
        "mmm": "%b",
        "mm": "%m",
        "m": "%#m",
        "dddd": "%A",
        "ddd": "%a",
        "dd": "%d",
        "d": "%#d",  # %#d removes leading zero on day
        "hh": "%H",
        "h": "%#H",
        "ss": "%S",
        "s": "%#S",
        "am/pm": "%p",
    }

    minuts_map = {"mm": "%M", "m": "%#M"}

    date = excel_datetime(number_str)
    spl = RE_DATAFORMAT.findall(format_str)

    for index, x in enumerate(spl):
        if x.lower() not in format_map:
            continue
        found_hours = False
        if x.lower() in ["m", "mm"] and index > 0:
            for i2 in range(index - 1, 0, -1):
                if not spl[i2].startswith("%"):
                    continue
                if spl[i2].lower().endswith("h"):
                    found_hours = True
                    spl[index] = minuts_map[x.lower()]
                break
        if found_hours:
            spl[index] = minuts_map[x.lower()]
        else:
            spl[index] = format_map[x.lower()]

    formatted_datetime = date.strftime("".join(spl))

    return formatted_datetime


def excel_datetime(excel_datestring, date_system=1900):
    if not isinstance(excel_datestring, float):
        excel_datestring = float(excel_datestring)
    etime = excel_datestring % 1
    if excel_datestring < 0:
        etime = 1 - etime
    edate = excel_datestring // 1

    if edate > 59:
        edate -= 1

    seconds = int(round(etime * 86400))

    date = datetime.fromordinal(
        datetime(date_system, 1, 1).toordinal() + int(edate) - (1 if date_system == 1900 else 0)
    ) + timedelta(seconds=seconds)

    return date


def format_number(number_str: str, format_str: str) -> str:
    # if format_str == "":
    #     format_str = "#"
    try:
        value = float(number_str)
    except ValueError:
        return number_str

    # Handle fraction formats like "# ???/???"
    if "?" in format_str and "/" in format_str:
        # Determine how many digits are used for numerator/denominator
        frac_digits = format_str.split("/")[-1].count("?")
        max_denominator = 10**frac_digits
        frac = Fraction(value).limit_denominator(max_denominator)
        whole = int(frac)
        numerator = frac.numerator - whole * frac.denominator
        if numerator == 0:
            result = f"{whole}"
        else:
            num_field = "?" * frac_digits
            den_field = "?" * frac_digits
            result = (
                f"{whole} "
                f"{str(numerator).rjust(len(num_field))}/"
                f"{str(frac.denominator).ljust(len(den_field))}"
            )
        return result

    format_section = select_format_section(value, format_str)
    format_section = strip_color_tags(format_section)
    format_str = format_section

    if data_formatted := format_data(number_str, format_str):
        return data_formatted

    if value == 0:
        if RE_ZERO.fullmatch(format_section.strip()):
            # If format is just literal text (e.g., "Zero")
            return format_section.strip()

    value_is_negative = value < 0
    value = abs(value)
    negative_parentheses = format_section.startswith("(") and format_section.endswith(")")

    if negative_parentheses:
        format_str = format_str[1:-1]

    procent_format = False
    if format_str.endswith("%"):
        procent_format = True
        value *= 100.0
        format_str = format_str[:-1]

    currency_format = False
    if format_str.startswith("$"):
        currency_format = True
        format_str = format_str[1:]

    # Handle comma scaling
    comma_scale = 0
    while format_str and format_str[-1 - comma_scale] == ",":
        comma_scale += 1

    value /= 1000**comma_scale
    
    if "." in format_str:
        int_fmt, dec_fmt = format_str.split(".")
    elif format_str in ("General", ""):
        int_fmt = "#"
        if value != round(value):
            dec_fmt = "#" * len(str(value).split(".")[-1])
        else:
            dec_fmt = ""

    else:
        int_fmt, dec_fmt = format_str, ""

    # Count decimal digits to round
    num_decimals = len(dec_fmt.replace("?", "").replace("#", "").replace("0", ""))
    if num_decimals == 0:
        num_decimals = len(dec_fmt)
    rounded = round(value, num_decimals)

    int_part, _, dec_part = f"{rounded:.{num_decimals}f}".partition(".")

    # Add thousands separator if format has comma
    if "," in format_str and comma_scale == 0:
        int_part = f"{int(int_part):,}"

    # Apply padding or stripping rules to integer part
    if "?" in int_fmt:
        int_part = int_part.rjust(len(int_fmt))
    # elif "0" in int_fmt:
    #     int_part = int_part.zfill(len(int_fmt))

    # Handle decimal part
    if dec_fmt:
        dec_required = dec_fmt.count("0")
        dec_optional = dec_fmt.count("#") + dec_fmt.count("?")
        total_places = dec_required + dec_optional
        dec_part = (dec_part + "0" * total_places)[:total_places]

        if "?" in dec_fmt:
            # Replace missing optional digits with spaces
            full = dec_part[: len(dec_fmt)]
            dec_part = "".join(
                ch if f in "0" or (f in "#?" and ch != "0") else " " for ch, f in zip(full, dec_fmt)
            )
        elif "#" in dec_fmt:
            dec_part = rstrip_limit(dec_part, "0", dec_fmt.count("#"))

        result = f"{int_part}.{dec_part}" if dec_part else int_part
    else:
        result = int_part

    # Handle negative formatting with parentheses
    if procent_format:
        result += "%"
    if currency_format:
        result = f"${result}"
    if value_is_negative:
        if negative_parentheses:
            result = f"({result})"
        else:
            result = f"-{result}"
    return result
