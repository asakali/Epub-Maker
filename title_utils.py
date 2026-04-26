"""标题/章节识别工具函数"""

import re

MANUAL_HEADING_PATTERN = re.compile(r"^【--([1-4])】(.+)$")
MANUAL_NON_HEADING_PREFIX = "【==】"
MAX_HEADING_LINE_LENGTH = 30

SUSPICIOUS_TITLE_KEYWORDS = [
    "求订阅",
    "求月票",
    "求推荐票",
    "补昨天",
    "补更",
    "加更",
    "今日",
    "二更",
    "三更",
    "四更",
    "五更",
    "上架感言",
    "感谢",
    "PS",
    "ps",
    "读者",
    "请假",
]


def is_heading_title(line, keyword):
    line = line.strip()
    if not line:
        return False
    if len(line) > MAX_HEADING_LINE_LENGTH:
        return False
    pattern = (
        rf"^第[一二三四五六七八九十百千万零〇两\d]+{re.escape(keyword)}(?:\s*.*)?$"
    )
    return re.match(pattern, line) is not None


def is_manual_heading_title(line):
    line = line.strip()
    match = MANUAL_HEADING_PATTERN.match(line)
    return bool(match and match.group(2).strip())


def is_manual_non_heading(line):
    return line.strip().startswith(MANUAL_NON_HEADING_PREFIX)


def strip_manual_non_heading_marker(line):
    leading_match = re.match(r"^(\s*)", line)
    leading = leading_match.group(1) if leading_match else ""
    content = line[len(leading):]
    if content.startswith(MANUAL_NON_HEADING_PREFIX):
        return leading + content[len(MANUAL_NON_HEADING_PREFIX):].lstrip()
    return line


def split_title(title_line, keyword):
    """
    返回:
    prefix = 第九章
    subtitle = 二面
    full_title = 第九章 二面 或 第九章
    number = 9（用于连续性检查，尽量转换，失败则 None）
    """
    pattern = rf"^(第[一二三四五六七八九十百千万零〇两\d]+{re.escape(keyword)})(?:\s+(.*))?$"

    m = re.match(pattern, title_line.strip())
    if not m:
        return None, None, title_line.strip(), None

    prefix = m.group(1).strip()
    subtitle = (m.group(2) or "").strip()
    full_title = prefix if not subtitle else f"{prefix} {subtitle}"
    number = extract_cn_number(prefix)
    return prefix, subtitle, full_title, number


def split_manual_heading_title(title_line):
    match = MANUAL_HEADING_PATTERN.match(title_line.strip())
    if not match:
        return 1, "", ""
    level = int(match.group(1) or 1)
    title = match.group(2).strip()
    return level, title, title


def extract_cn_number(text):
    m = re.search(r"第([一二三四五六七八九十百千万零〇两\d]+)", text)
    if not m:
        return None
    s = m.group(1)
    if s.isdigit():
        return int(s)
    return chinese_to_int(s)


def chinese_to_int(s):
    digit_map = {
        "零": 0,
        "〇": 0,
        "一": 1,
        "二": 2,
        "两": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
    }
    unit_map = {"十": 10, "百": 100, "千": 1000, "万": 10000, "亿": 100000000}

    if not s:
        return None

    total = 0
    section = 0
    number = 0

    for ch in s:
        if ch in digit_map:
            number = digit_map[ch]
        elif ch in unit_map:
            unit = unit_map[ch]
            if unit >= 10000:
                section = (section + (number if number != 0 else 0)) * unit
                total += section
                section = 0
                number = 0
            else:
                if number == 0:
                    number = 1
                section += number * unit
                number = 0
        else:
            return None

    return total + section + number


def has_suspicious_title_tail(line):
    return any(k in line for k in SUSPICIOUS_TITLE_KEYWORDS)


def next_line_starts_with_punct(lines, index, heading_configs):
    if index + 1 >= len(lines):
        return False

    next_line = lines[index + 1].strip()
    if not next_line:
        return False

    for config in heading_configs:
        if is_heading_title(next_line, config["keyword"]):
            return False

    suspicious_prefixes = (
        "！",
        "？",
        "…",
        "——",
        "—",
        "）",
        "】",
        "”",
        "」",
        "』",
        "：",
        "、",
        ".",
        "。",
    )
    return next_line.startswith(suspicious_prefixes)
