"""
Parse scraper log output and extract structured progress data.

Usage:
    python parse_progress.py <log_text.txt
    python parse_progress.py --json <log_text.txt

Output: JSON with keys:
    timestamp, total_scraped, categories: {code: {count, target, pct}},
    active_tags: [{tag, clc_code, confidence}], errors: [str], rate: float|None
"""
import sys
import re
import json
from datetime import datetime


CATEGORIES = {
    "A": ("马列", 1500), "B": ("哲学宗教", 5000), "C": ("社科总论", 5000),
    "D": ("政治法律", 4000), "E": ("军事", 2500), "F": ("经济", 5000),
    "G": ("文化教育", 4000), "H": ("语言文字", 3500), "I": ("文学", 8000),
    "J": ("艺术", 5000), "K": ("历史地理", 6000), "N": ("自然科学", 3000),
    "O": ("数理化", 3000), "P": ("天文地球", 2000), "Q": ("生物", 3000),
    "R": ("医药卫生", 4000), "S": ("农业", 2500), "T": ("工业技术", 7000),
    "U": ("交通", 1500), "V": ("航空航天", 1500), "X": ("环境", 2000),
    "Z": ("综合", 4000),
}


def parse_log(text: str) -> dict:
    result = {
        "timestamp": None,
        "total_scraped": None,
        "categories": {},
        "active_tags": [],
        "errors": [],
        "rate": None,
    }

    # Timestamp
    ts_match = re.search(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})", text)
    if ts_match:
        result["timestamp"] = ts_match.group(1)

    # Total scraped
    total_match = re.search(r"进度:\s*(\d+)/100000", text)
    if total_match:
        result["total_scraped"] = int(total_match.group(1))

    # Category distribution lines
    for line in text.splitlines():
        cat_match = re.match(
            r"\s*([A-Z]):\s+[#.]*\s*(\d+)/(\d+)\s+\((\d+)%\)", line
        )
        if cat_match:
            code = cat_match.group(1)
            count = int(cat_match.group(2))
            target = int(cat_match.group(3))
            pct = int(cat_match.group(4))
            result["categories"][code] = {"count": count, "target": target, "pct": pct}

    # Active tag
    tag_match = re.search(
        r"\[douban\]\s+(.+?)\s+→\s+(\w+)\s+\(([\d.]+)\)", text
    )
    if tag_match:
        result["active_tags"].append({
            "tag": tag_match.group(1),
            "clc_code": tag_match.group(2),
            "confidence": float(tag_match.group(3)),
        })

    # Dangdang active source
    dd_match = re.search(
        r"\[dangdang\]\s+(.+?)\s+→\s+(\w+)\s+\(([\d.]+)\)", text
    )
    if dd_match:
        result["active_tags"].append({
            "tag": dd_match.group(1),
            "clc_code": dd_match.group(2),
            "confidence": float(dd_match.group(3)),
        })

    # Errors
    for err_pattern in [
        r"PermissionError:.*",
        r"\[Errno \d+\].*",
        r"HTTP.*(?:418|429|403|500|404).*",
        r"网络错误.*",
        r"被封.*",
    ]:
        for match in re.finditer(err_pattern, text):
            result["errors"].append(match.group(0).strip())

    # Rate
    rate_match = re.search(r"(\d+)\s*条/小时", text)
    if rate_match:
        result["rate"] = int(rate_match.group(1))

    return result


def compute_delta(prev: dict, curr: dict) -> dict:
    """Compute per-category deltas between two snapshots."""
    delta = {}
    for code in CATEGORIES:
        prev_count = prev.get(code, {}).get("count", 0) if prev else 0
        curr_count = curr.get(code, {}).get("count", 0)
        diff = curr_count - prev_count
        delta[code] = diff
    return delta


def main():
    text = sys.stdin.read()
    result = parse_log(text)

    if "--json" in sys.argv:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Timestamp:  {result['timestamp']}")
        print(f"Total:      {result['total_scraped']}/100000")
        print(f"Rate:       {result['rate']} 条/小时")
        print(f"Categories: {len(result['categories'])} classes")
        for code, info in sorted(result["categories"].items()):
            print(f"  {code}: {info['count']}/{info['target']} ({info['pct']}%)")
        if result["active_tags"]:
            print(f"Tags:       {result['active_tags']}")
        if result["errors"]:
            print(f"Errors:     {len(result['errors'])}")
            for e in result["errors"]:
                print(f"  - {e}")


if __name__ == "__main__":
    main()
