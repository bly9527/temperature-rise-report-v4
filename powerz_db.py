"""Read-only analysis of native Power-Z VBUS/IBUS recordings."""
from bisect import bisect_left, bisect_right
from contextlib import closing
from math import isfinite
from pathlib import Path
import sqlite3
from statistics import median
from thermal_phases import clock_text, image_digest, parse_elapsed

POWERZ_DB_FILTER = 'Power-Z (*.db *.sqlite *.sqlite3)'


def read_trace(path):
    path = Path(path)
    if not path.is_file():
        raise ValueError(f'找不到 Power-Z 数据库：{path}')
    try:
        with closing(sqlite3.connect(path.resolve().as_uri() + '?mode=ro', uri=True)) as db:
            names = {col[1].lower() for col in db.execute('PRAGMA table_info(table_1)')}
            elapsed = None
            if 'elapsedtime' in names:
                first = db.execute('SELECT Unix, ElapsedTime FROM table_1 WHERE Unix IS NOT NULL '
                                   'ORDER BY Unix LIMIT 1').fetchone()
                if first and first[1]:
                    elapsed = float(first[0]) - parse_elapsed(first[1])
            rows = db.execute('SELECT Unix, VBUS, IBUS FROM table_1 ORDER BY Unix')
            trace = []
            for row in rows:
                if None in row:
                    continue
                t, v, a = map(float, row)
                if all(map(isfinite,(t,v,a))):
                    if elapsed is None:
                        elapsed = t
                    trace.append((round(t-elapsed,6),v,a,abs(v*a)))
    except (sqlite3.Error, TypeError, ValueError) as error:
        raise ValueError(f'{path.name} 缺少可读取的 Power-Z 时间、电压或电流记录。') from error
    if len(trace) < 12 or any(b[0]<=a[0] for a,b in zip(trace,trace[1:])):
        raise ValueError(f'{path.name} 的有效采样不足或时间重复。')
    return trace


def _events(rows, times, period, active):
    # Detect edges in native samples. The pre-edge reference scales with the
    # sampling cadence. Short interruptions may recover to a higher plateau.
    power = [r[3] for r in rows]
    first = next((i for i in range(len(power) - 2)
                  if min(power[i:i + 3]) >= active
                  and times[i + 2] - times[i] <= 3 * period), None)
    if first is None:
        return None, []
    events = []
    last = -1e30
    stride = max(1,round(0.35/period))
    for i in range(max(first+1,stride),len(rows)-1):
        t=times[i]
        if t-last < max(0.5,period) or times[i]-times[i-stride]>max(.75,period*stride*1.5):
            continue
        before = power[i-stride]
        if before<active or before-power[i] < max(.35,before*.09):
            continue
        start = max(first,bisect_left(times,t-max(.6,period*2)))
        if start >= i:
            continue
        before = median(power[start:i])
        if before-power[i] < max(.35,before*.09):
            continue
        # A transient at a commanded transition is distinct from the eventual load.
        settle = bisect_left(times,t+12)
        end = bisect_right(times,t+15)
        if settle >= end or times[settle]-t > 12+max(.5,period*1.5):
            settle=i
            end=min(len(rows),i+2)
        window=rows[settle:end]
        vp=median(r[1] for r in window)
        ap=median(r[2] for r in window)
        p_after=median(r[3] for r in window)
        if p_after < active:
            label='停载'
        elif p_after > before * .96:
            label='功率短时下探后恢复'
        else:
            label=f'降至约 {vp:.2f}V / {ap:.2f}A'
        events.append((i,label))
        last=t
        # Ignore the rest of a multi-sample edge.
        if label=='停载':
            last=t+max(1.,period)

    # A smooth taper has no unique edge. Mark the measured departure from a
    # stable plateau only when the next minute corroborates a sustained drop.
    next_check = times[first]+80
    last_taper = -1e30
    hard_times = [times[i] for i,_ in events]
    for i in range(first,len(rows)):
        t=times[i]
        if t<next_check:
            continue
        next_check=t+1
        if t-last_taper<120 or t+60>times[-1]:
            continue
        a,b=bisect_left(times,t-80),bisect_left(times,t-20)
        if b<=a or bisect_left(hard_times,t-80)!=bisect_right(hard_times,t+60):
            continue
        if abs(times[a]-(t-80))>max(1,period) or abs(times[b]-(t-20))>max(1,period):
            continue
        j,k=bisect_left(times,t+30),bisect_left(times,t+60)
        if k>=len(rows) or max(times[j]-t-30,times[k]-t-60)>max(1,period):
            continue
        base=power[b]
        if base<max(1,active*3) or abs(power[a]-base)>base*.004:
            continue
        if not (base-power[i]>=base*.006 and power[i]-power[j]>=base*.012
                and power[j]-power[k]>=base*.012):
            continue
        events.append((i,'持续减流起点（趋势估算）'))
        last_taper=t
    return first, sorted(events)


def analyze_meter(path):
    rows=read_trace(path)
    times=[r[0] for r in rows]
    period=median(b-a for a,b in zip(times,times[1:]))
    active=max(.2,max(r[3] for r in rows)*.025)
    first, edges=_events(rows,times,period,active)
    if first is None:
        raise ValueError('Power-Z 数据库中未找到有效带载区段。')
    events=[]
    table=[['DB 相对时间','功率变化','采样电压(V)','采样电流(A)','功率绝对值(W)']]
    for i,label in edges:
        t,v,a,p=rows[i]
        events.append({'seconds':t,'label':label})
        table.append([clock_text(t),label,round(v,4),round(a,4),round(p,4)])
    return {'events':events,'load_start':times[first],'event_table':table,
            'period':period,'samples':len(rows)}


def bind_meter(path, settings):
    """Build the saved DB binding from a readable, unchanged recording."""
    digest = image_digest(path)
    analysis = analyze_meter(path)
    if image_digest(path) != digest:
        raise ValueError('曲线 DB 正在变化，请保存记录后重新选择。')
    updated = dict(settings, db_path=str(path), db_hash=digest,
                   events=analysis['events'], load_start=analysis['load_start'])
    updated.pop('source', None)
    if settings.get('db_hash') != digest:
        for key in ('summary', 'table'):
            updated.pop(key, None)
    return updated, analysis


def switch_to_picture(settings):
    updated = dict(settings, source='picture')
    if updated.pop('db_path', None):
        for key in ('db_hash', 'events', 'load_start', 'summary', 'table'):
            updated.pop(key, None)
    return updated
