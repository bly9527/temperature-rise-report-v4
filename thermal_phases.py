"""Map screenshot elapsed times to measured temperatures in development reports."""
from bisect import bisect_left
from collections import defaultdict
from hashlib import sha256
from pathlib import Path
from statistics import median
import math
import re


def image_digest(path):
    return sha256(Path(path).read_bytes()).hexdigest() if path else ''


def clock_text(seconds):
    milliseconds = round(seconds * 1000)
    minutes, remainder = divmod(milliseconds, 60000)
    hours, minute = divmod(minutes, 60)
    sec, fraction = divmod(remainder, 1000)
    result = f'{hours}:{minute:02d}:{sec:02d}'
    return result + (f'.{fraction:03d}'.rstrip('0') if fraction else '')


def parse_elapsed(text):
    if not re.fullmatch(r'\d+:[0-5]\d:[0-5]\d(?:\.\d{1,3})?', text.strip()):
        raise ValueError('时间格式示例：0:06:45 或 0:06:45.3。')
    hours, minutes, seconds = text.strip().split(':')
    return int(hours)*3600 + int(minutes)*60 + float(seconds)


def phase_text(value):
    return '\n'.join(line.strip() for line in value.splitlines() if line.strip()
                     and not re.search(r'测温采集|带载(?:阶段)?时长待核', line))


def meter_steps(picture):
    """Use the source trace when a Power-Z screenshot has a matching database."""
    from powerz_db import analyze_meter
    db = Path(picture).with_suffix('.db')
    if not db.is_file():
        return None
    return analyze_meter(db)


def locate_events(data, channels, events, load_start=0):
    """Find a sustained multichannel rise, then map to actual, gap-checked scans."""
    sensors = [i + 2 for i, ch in enumerate(channels)
               if not re.search('环境|环温|ambient', ch['custom_name'], re.I)]
    if not sensors:
        return None, {}, ['只有环境通道']
    if any(not isinstance(row[0], (int, float)) for row in data):
        return None, {}, ['扫描时间无法换算']
    seconds = [(r[0] - data[0][0]) * 86400 for r in data]
    gaps = [b - a for a, b in zip(seconds, seconds[1:])]
    if not gaps or min(gaps) < 0 or max(gaps) == 0:
        return None, {}, ['扫描时间不足或未按顺序排列']
    period = median(gap for gap in gaps if gap > 0)
    # Look forward over a short confirmation window, then backtrack the first
    # persistent upward departure. Cooling before the test is common.
    window = max(3, math.ceil(30 / period))
    support = min(2, len(sensors))
    onset = None
    for i in range(len(data) - window):
        end = i + window
        if seconds[end] - seconds[i] > 2 * window * period:
            continue
        rising = []
        for col in sensors:
            start = data[i][col]
            tail = [data[j][col] for j in range(end - 2, end + 1)]
            if start is not None and all(v is not None for v in tail) and min(tail) >= start + 0.35:
                rising.append(col)
        if len(rising) >= support:
            for j in range(i + 1, end + 1):
                if sum(data[j][k] is not None and data[j][k] >= data[i][k] + 0.10
                       for k in rising) >= support:
                    onset = max(i, j - 1)
                    break
        if onset is not None:
            break
    if onset is None:
        return None, {}, ['未检出持续升温']
    marks, skipped = defaultdict(list), []
    for event in events:
        delta = event['seconds'] - load_start
        target = seconds[onset] + delta
        idx = bisect_left(seconds, target)
        nearby = [j for j in (idx - 1, idx) if 0 <= j < len(data)]
        nearest = min(nearby, key=lambda j: abs(seconds[j] - target))
        if delta < 0 or target > seconds[-1] + 0.001 or abs(seconds[nearest] - target) > period / 2 + 0.02:
            skipped.append(clock_text(event['seconds']) + ' 超出有效测温区间')
        else:
            marks[nearest].append(clock_text(event['seconds']) + ' ' + event['label'])
    return onset, dict(marks), skipped
