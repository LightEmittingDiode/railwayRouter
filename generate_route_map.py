import openpyxl, sys, csv, json, re, os, collections
from datetime import time as dt_time
sys.stdout.reconfigure(encoding='utf-8')

# 스크립트가 위치한 현재 폴더 경로를 자동으로 가져옵니다.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FILES = [
    os.path.join(BASE_DIR, 'src', 'KTX_timetable.xlsx'),
    os.path.join(BASE_DIR, 'src', 'ITXC_timetable.xlsx'),
    os.path.join(BASE_DIR, 'src', 'Local_timetable.xlsx'),
    os.path.join(BASE_DIR, 'src', 'SRT_timetable.xlsx'),
]

CSV_PATH = os.path.join(BASE_DIR, 'src', 'StationLocation_202404.csv')

S_PAT = re.compile(r'^S\d{4}$')
INVALID_EXACT = {
    '열차번호','열차종별','비 고','비고','종착역','시발역','열차',
    'KTX','SRT','무궁화','ITX','누리로','새마을','KTX-산천','KTX-청룡',
    'ITX-마음','ITX-새마을','KTX-이음','KTX_산천', '편성', '연계열번'
}

hangul_pattern = re.compile(r'[\uac00-\ud7a3]')
hanja_pattern = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf]')

def clean_station_name(v):
    if v is None: return None
    s = str(v).strip()
    if not s or s == '_': return None
    if S_PAT.match(s): return None
    if s in INVALID_EXACT: return None
    if hanja_pattern.search(s): return None  # 한자 제거
    
    # 통합 예외 처리
    if s in ('홍성(도착)', '홍성(출발)'):
        return '홍성'
        
    if not hangul_pattern.search(s): return None  # 순수 한글 필수
    
    # 헤더 메타데이터 제거
    if re.search(r'선|발|행|예시|년|월|일|보기|주말|평일|운행|정차|통과|시표|시간표', s) and len(s) > 4:
        return None
    if re.search(r'^[월화수목금토일]+$', s):
        return None
        
    s_clean = re.sub(r'\s+', '', s)
    if len(s_clean) > 10 or len(s_clean) < 2: return None
    return s_clean

import datetime

def time_to_mins(t):
    if isinstance(t, datetime.time):
        if t.hour == 0 and t.minute == 0:
            return None
        return t.hour * 60 + t.minute
    elif isinstance(t, str):
        parts = t.strip().split(':')
        if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
            h, m = int(parts[0]), int(parts[1])
            if h == 0 and m == 0:
                return None
            return h * 60 + m
    return None

def is_station_name(v):
    return clean_station_name(v) is not None

class RouteTrip:
    def __init__(self, train_no, train_type):
        self.train_no = str(train_no).strip()
        self.train_type = str(train_type).strip() if train_type else '열차'
        self.stops = []

    def add_station(self, name, time_mins=None):
        st = clean_station_name(name)
        if st:
            # Handle day rollover for times crossing midnight (like in train_router.py)
            if self.stops and time_mins is not None:
                prev_time = self.stops[-1]['time']
                if prev_time is not None:
                    t = time_mins
                    if t <= prev_time:
                        if (prev_time - t) > 18 * 60: 
                            t += 24 * 60
                    time_mins = t
            self.stops.append({'name': st, 'time': time_mins})

    def __repr__(self):
        return f"{self.train_no} ({self.train_type}): {self.stops}"

def filter_branches(station_times, ttype):
    # station_times is a list of (st_name, cell_val)
    stopped_names = set()
    for st_name, cell_val in station_times:
        if cell_val is not None:
            s_val = str(cell_val).strip()
            if s_val and re.search(r'\d', s_val) and '00:00:00' not in s_val:
                clean_st = st_name.split('(')[0] if '(' in st_name else st_name
                stopped_names.add(clean_st)
                
    to_remove = set()
    
    if 'SRT' in ttype:
        to_remove.update({'행신', '서울', '용산', '영등포', '수원', '광명'})
    else:
        to_remove.update({'수서', '동탄', '평택지제'})
        
    if '수원' in stopped_names or '영등포' in stopped_names:
        to_remove.update({'광명'})
    else:
        to_remove.update({'수원', '영등포'})
        
    if any(x in stopped_names for x in ['경산', '밀양', '물금', '구포']):
        to_remove.update({'경주', '울산'})
    else:
        to_remove.update({'경산', '밀양', '물금', '구포'})
        
    if any(x in stopped_names for x in ['서대전', '계룡', '논산']):
        to_remove.update({'공주'})
    else:
        to_remove.update({'서대전', '계룡', '논산'})
        
    if any(x in stopped_names for x in ['정동진', '묵호', '동해']):
        to_remove.update({'강릉'})
    else:
        to_remove.update({'정동진', '묵호', '동해'})
        
    filtered_times = []
    stopped_indices = []
    for st_name, cell_val in station_times:
        clean_st = st_name.split('(')[0] if '(' in st_name else st_name
        if clean_st not in to_remove:
            filtered_times.append((st_name, cell_val))
            has_time = False
            if cell_val is not None:
                s_val = str(cell_val).strip()
                if s_val and re.search(r'\d', s_val) and '00:00:00' not in s_val:
                    has_time = True
            if has_time:
                stopped_indices.append(len(filtered_times) - 1)
                
    return filtered_times, stopped_indices

all_trips = []

def parse_sheet_as_cols(ws, fname):
    """역명이 열(Column) 방향으로 나열된 시트 파싱"""
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return []
    trips = []
    
    for r_idx, row in enumerate(rows):
        has_train_no = False
        for val in row:
            if val is not None and '열차번호' in str(val).replace(' ', ''):
                has_train_no = True
                break
                
        if has_train_no:
            train_no_row = r_idx
            train_no_cols = []
            for c_idx, val in enumerate(row):
                if val is not None and '열차번호' in str(val).replace(' ', ''):
                    train_no_cols.append(c_idx)
                    
            for train_no_col in train_no_cols:
                station_cols = []
                for cc in range(train_no_col+1, len(row)):
                    cv = row[cc]
                    if cv is None: continue
                    sv_c = str(cv).strip().replace(' ', '')
                    if '열차번호' in sv_c or '비고' in sv_c:
                        break
                    if is_station_name(cv):
                        station_cols.append((cc, str(cv).strip()))
                        
                empty_count = 0
                for rr in range(train_no_row+1, len(rows)):
                    r_vals = rows[rr]
                    tno = r_vals[train_no_col] if train_no_col < len(r_vals) else None
                    if not tno or not str(tno).strip():
                        empty_count += 1
                        if empty_count >= 5: break
                        continue
                    empty_count = 0
                    tno_str = str(tno).strip()
                    if not re.search(r'\d', tno_str): continue
                    
                    ttype = '열차'
                    if train_no_col > 0:
                        tv = r_vals[train_no_col-1] if train_no_col-1 < len(r_vals) else None
                        if tv and is_station_name(tv) and str(tv).strip() in ('KTX','무궁화','ITX-마음','ITX-새마을','새마을','누리로','SRT','KTX-산천','KTX-청룡'):
                            ttype = str(tv).strip()
                    
                    trip = RouteTrip(tno_str, ttype)
                    
                    station_times = []
                    for c_st, st_name in station_cols:
                        cell_val = r_vals[c_st] if c_st < len(r_vals) else None
                        station_times.append((st_name, cell_val))
                        
                    filtered_times, stopped_indices = filter_branches(station_times, ttype)
                            
                    if not stopped_indices:
                        continue
                        
                    first_idx = stopped_indices[0]
                    last_idx = stopped_indices[-1]
                    
                    for i, (st_name, cell_val) in enumerate(filtered_times):
                        if first_idx <= i <= last_idx:
                            time_mins = time_to_mins(cell_val)
                            trip.add_station(st_name, time_mins)
                    
                    if len(trip.stops) >= 2:
                        trips.append(trip)
            return trips
    return trips

def parse_sheet_as_rows(ws, fname):
    """역명이 행(Row) 방향으로 나열된 시트 파싱"""
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return []
    trips = []
    
    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            if val is None: continue
            sv = str(val).strip()
            if '열차번호' in sv.replace(' ', ''):
                train_no_row = r_idx
                train_no_col = c_idx
                
                train_cols = []
                for cc in range(c_idx+1, len(row)):
                    tv = row[cc]
                    if tv is not None and str(tv).strip() and re.search(r'\d', str(tv)):
                        ttype = '열차'
                        if train_no_row > 0:
                            prev_row = rows[train_no_row-1]
                            pv = prev_row[cc] if cc < len(prev_row) else None
                            if pv and str(pv).strip() in ('KTX','무궁화','ITX-마음','ITX-새마을','새마을','누리로','SRT','KTX-산천','KTX-청룡','ITX'):
                                ttype = str(pv).strip()
                        train_cols.append((cc, str(tv).strip(), ttype))
                
                if not train_cols: continue
                
                station_rows = []
                empty_count = 0
                for rr in range(train_no_row+1, len(rows)):
                    r_vals = rows[rr]
                    st_name = r_vals[train_no_col] if train_no_col < len(r_vals) else None
                    if not is_station_name(st_name):
                        empty_count += 1
                        if empty_count >= 5: break
                        continue
                    empty_count = 0
                    station_rows.append((rr, str(st_name).strip()))
                
                for c_col, tno, ttype in train_cols:
                    trip = RouteTrip(tno, ttype)
                    station_times = []
                    for rr, st_name in station_rows:
                        r_vals = rows[rr]
                        cell_val = r_vals[c_col] if c_col < len(r_vals) else None
                        station_times.append((st_name, cell_val))
                        
                    filtered_times, stopped_indices = filter_branches(station_times, ttype)
                            
                    if not stopped_indices:
                        continue
                        
                    first_idx = stopped_indices[0]
                    last_idx = stopped_indices[-1]
                    
                    for i, (st_name, cell_val) in enumerate(filtered_times):
                        if first_idx <= i <= last_idx:
                            time_mins = time_to_mins(cell_val)
                            trip.add_station(st_name, time_mins)
                    
                    if len(trip.stops) >= 2:
                        trips.append(trip)
                return trips
    return trips

print("엑셀 파일 파싱 중 (무정차 포함)...")
for fpath in FILES:
    if not os.path.exists(fpath):
        print(f"  파일 없음: {fpath}")
        continue
    fname = os.path.basename(fpath)
    wb = openpyxl.load_workbook(fpath, data_only=True)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        parsed = parse_sheet_as_cols(ws, fname)
        if not parsed:
            parsed = parse_sheet_as_rows(ws, fname)
        all_trips.extend(parsed)
    print(f"  {fname}: 누적 {len(all_trips)}개 열차")

# ─────────────────────────────────────────
# 2. 구간 & 노선 타입 집계
# ─────────────────────────────────────────
def get_line_type(train_no, train_type):
    tt = str(train_type).strip()
    if tt in ('KTX','KTX-산천','KTX-청룡','열차'):
        try:
            no = int(re.sub(r'[^0-9]','',train_no))
            if 1 <= no <= 299: return 'KTX경부'
            if 400 <= no <= 499: return 'KTX호남'
            if 700 <= no <= 799: return 'KTX중부'
            if 800 <= no <= 899: return 'KTX강릉'
            return 'KTX기타'
        except: return 'KTX기타'
    if tt == 'SRT': return 'SRT'
    if tt in ('ITX-새마을',): return 'ITX-새마을'
    if tt in ('ITX-마음','ITX'): return 'ITX-마음'
    if tt == '새마을': return '새마을'
    if tt in ('무궁화',): return '무궁화'
    if tt == '누리로': return '누리로'
    return '무궁화'

LINE_CONFIG = {
    'KTX경부':   {'color':'#003B72','width':5,'label':'KTX 경부·동해선','dash':[]},
    'KTX호남':   {'color':'#003B72','width':5,'label':'KTX 호남선','dash':[]},
    'KTX강릉':   {'color':'#003B72','width':5,'label':'KTX 강릉선','dash':[]},
    'KTX중부':   {'color':'#003B72','width':5,'label':'KTX 중부내륙선','dash':[]},
    'KTX기타':   {'color':'#003B72','width':4,'label':'KTX 기타','dash':[]},
    'SRT':       {'color':'#59233E','width':4,'label':'SRT','dash':[]},
    'ITX-새마을':{'color':'#AC192D','width':3,'label':'ITX-새마을','dash':[]},
    'ITX-마음':  {'color':'#AC192D','width':3,'label':'ITX-마음','dash':[]},
    'ITX-청춘':  {'color':'#AC192D','width':3,'label':'ITX-청춘','dash':[]},
    '새마을':    {'color':'#005388','width':3,'label':'새마을','dash':[]},
    '무궁화':    {'color':'#E84E0F','width':2,'label':'무궁화','dash':[]},
    '누리로':    {'color':'#E84E0F','width':2,'label':'누리로','dash':[]},
    '통근열차':  {'color':'#7F8C8D','width':1,'label':'통근열차','dash':[]}
}

PRIORITY = ['KTX경부','KTX호남','KTX강릉','KTX중부','SRT','KTX기타','ITX-새마을','ITX-마음','새마을','무궁화','누리로']

def primary_type(types_set):
    for p in PRIORITY:
        if p in types_set: return p
    return '무궁화'

edge_map = collections.defaultdict(set)
node_lines = collections.defaultdict(set)

for t in all_trips:
    lt = get_line_type(t.train_no, t.train_type)
    for s in t.stops:
        node_lines[s['name']].add(lt)
    for i in range(len(t.stops)-1):
        a, b = t.stops[i]['name'], t.stops[i+1]['name']
        if a == b: continue
        canon = tuple(sorted([a, b]))
        edge_map[canon].add(lt)

print(f"\n총 역: {len(node_lines)}, 총 구간: {len(edge_map)}")

# 일영 확인
if '일영' in node_lines:
    print(f"일영역: {node_lines['일영']}")
else:
    print("⚠️  일영역이 없습니다!")

# ─────────────────────────────────────────
# 3. CSV 좌표 로드
# ─────────────────────────────────────────
station_coords = {}
with open(CSV_PATH, 'r', encoding='euc-kr', errors='ignore') as f:
    reader = csv.DictReader(f)
    for row in reader:
        name = row.get('역명','').strip()
        try:
            lat = float(row.get('위도',0))
            lon = float(row.get('경도',0))
            if lat and lon:
                station_coords[name] = {'lat':lat,'lon':lon}
        except: pass

# OSM 캐시로 누락 좌표 보완 (fetch_osm_coords.py 실행 후 생성됨)
OSM_CACHE_FILE = os.path.join(BASE_DIR, 'src', 'osm_coords_cache.json')
osm_used = 0
if os.path.exists(OSM_CACHE_FILE):
    with open(OSM_CACHE_FILE, 'r', encoding='utf-8') as f:
        osm_cache = json.load(f)
    for name, data in osm_cache.items():
        if name not in station_coords and data.get('lat') and data.get('lon'):
            station_coords[name] = {'lat': data['lat'], 'lon': data['lon']}
            osm_used += 1
    print(f"OSM 캐시 적용: {osm_used}개 역 좌표 보완")
else:
    print("ℹ️  OSM 캐시 없음 (fetch_osm_coords.py를 먼저 실행하면 좌표가 더 정확해집니다)")

# 동명이역 및 OSM 누락 역 수동 고정 좌표
station_coords['판교(경기)'] = {'lat': 37.39472, 'lon': 127.11153} # 경강선 판교역 (성남)
station_coords['양원'] = {'lat': 36.9392, 'lon': 129.1308} # 영동선 양원역 (봉화)
station_coords['송정'] = {'lat': 35.1812, 'lon': 129.2005} # 동해선 송정역 (부산)
station_coords['춘천'] = {'lat': 37.8847, 'lon': 127.7169}
station_coords['평내호평'] = {'lat': 37.6533, 'lon': 127.2436}
station_coords['평내호'] = {'lat': 37.6533, 'lon': 127.2436} # 평내호평 짤림 방지
station_coords['철암'] = {'lat': 37.1081, 'lon': 129.0289}
station_coords['용문'] = {'lat': 37.4816, 'lon': 127.5947}
station_coords['센텀'] = {'lat': 35.1685, 'lon': 129.1215}
station_coords['장흥'] = {'lat': 37.7208, 'lon': 126.9472}
station_coords['장동'] = {'lat': 34.7870, 'lon': 126.9940}
station_coords['강진'] = {'lat': 34.6432, 'lon': 126.7663}
station_coords['해남'] = {'lat': 34.5735, 'lon': 126.6021}
station_coords['영암'] = {'lat': 34.8001, 'lon': 126.6983}
station_coords['문경'] = {'lat': 36.5925, 'lon': 128.1611}
station_coords['수안보온천'] = {'lat': 36.8488, 'lon': 127.9942}
station_coords['연풍'] = {'lat': 36.8222, 'lon': 128.0063}
station_coords['삽교'] = {'lat': 36.6805, 'lon': 126.7583}
station_coords['신례원'] = {'lat': 36.7511, 'lon': 126.8362}
station_coords['여수'] = {'lat': 34.7555, 'lon': 127.7492}
station_coords['별내'] = {'lat': 37.6430, 'lon': 127.1260}
station_coords['살미'] = {'lat': 36.9024318, 'lon': 127.9602931}
station_coords['쌍룡'] = {'lat': 37.17464573, 'lon': 128.3284968}

# CSV에 좌표가 잘못 기재되어 삐죽하게 튀어나오는 역들 수동 교정
station_coords['나주'] = {'lat': 35.0139, 'lon': 126.7175}
station_coords['무안'] = {'lat': 34.9628, 'lon': 126.5176}
station_coords['다시'] = {'lat': 35.0165, 'lon': 126.6411}
station_coords['함평'] = {'lat': 35.0237, 'lon': 126.5391}
station_coords['군북'] = {'lat': 35.2810, 'lon': 128.3180}
station_coords['횡천'] = {'lat': 35.0860, 'lon': 127.7980}

if '편성' in station_coords: del station_coords['편성']

# ─────────────────────────────────────────
# 4. vis.js 노드/엣지 데이터 빌드
# ─────────────────────────────────────────
LAT_MIN,LAT_MAX = 33.0,38.8
LON_MIN,LON_MAX = 125.5,130.2
BASE_W,BASE_H = 2000,2400

def geo_to_xy(lat, lon):
    x = (lon - LON_MIN) / (LON_MAX - LON_MIN) * BASE_W
    y = (LAT_MAX - lat) / (LAT_MAX - LAT_MIN) * BASE_H
    return round(x,1), round(y,1)

ktx_types = {'KTX경부','KTX호남','KTX강릉','KTX중부','KTX기타','SRT'}
hub_stations = {'서울','용산','영등포','청량리','행신','수원','대전','동대구','부산','광명',
                '오송','익산','광주송정','목포','강릉','부전','포항','동해','여수엑스포','수서','동탄'}

station_neighbors = collections.defaultdict(set)
for (a, b) in edge_map:
    station_neighbors[a].add(b)
    station_neighbors[b].add(a)

station_services = collections.defaultdict(lambda: {
    'trains': set(), 'types': set(), 'first': None, 'last': None
})
for trip in all_trips:
    trip_key = f"{trip.train_type}:{trip.train_no}"
    for stop in trip.stops:
        service = station_services[stop['name']]
        service['trains'].add(trip_key)
        service['types'].add(trip.train_type)
        if stop['time'] is not None:
            minute = stop['time'] % (24 * 60)
            service['first'] = minute if service['first'] is None else min(service['first'], minute)
            service['last'] = minute if service['last'] is None else max(service['last'], minute)

nodes_data = []
for name, types in sorted(node_lines.items()):
    coord = station_coords.get(name)
    pt = primary_type(types)
    service = station_services[name]
    nodes_data.append({
        'id': name,
        'lat': coord['lat'] if coord else None,
        'lon': coord['lon'] if coord else None,
        'primary': pt,
        'types': [p for p in PRIORITY if p in types],
        'isHub': name in hub_stations,
        'isMajor': bool(ktx_types & types),
        'hasCoord': coord is not None,
        'neighbors': sorted(station_neighbors[name]),
        'degree': len(station_neighbors[name]),
        'trainCount': len(service['trains']),
        'firstTime': service['first'],
        'lastTime': service['last']
    })

edges_data = []
for (a,b), types_set in sorted(edge_map.items()):
    edges_data.append({'from':a,'to':b,'lines':[p for p in PRIORITY if p in types_set]})

missing_coord_names = sorted(n['id'] for n in nodes_data if not n['hasCoord'])
if missing_coord_names:
    print(f"⚠️  좌표 없는 역 {len(missing_coord_names)}개: {', '.join(missing_coord_names)}")

# ─────────────────────────────────────────
# 5. HTML Canvas 노선도 생성
# ─────────────────────────────────────────
line_cfg_js = {k:{**v} for k,v in LINE_CONFIG.items()}
nodes_json = json.dumps(nodes_data, ensure_ascii=False)
edges_json = json.dumps(edges_data, ensure_ascii=False)
priority_json = json.dumps(PRIORITY)
linecfg_json = json.dumps(line_cfg_js, ensure_ascii=False)

# 열차 시간표 데이터 (다익스트라 라우팅용) 생성
timetable_export = []
for trip in all_trips:
    stops_export = []
    for s in trip.stops:
        if s.get('time') is not None:
            stops_export.append([s['name'], s['time']])
    if len(stops_export) > 1:
        # Sort stops by time just in case
        stops_export.sort(key=lambda x: x[1])
        timetable_export.append([trip.train_no, trip.train_type, stops_export])
timetable_json = json.dumps(timetable_export, ensure_ascii=False)

html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Railway Router</title>
<style>
* {{ box-sizing:border-box; margin:0; padding:0; }}
html, body {{ width:100%; height:100%; }}
body {{ background:#08101f; font-family:'Pretendard','Noto Sans KR','Malgun Gothic',sans-serif; overflow:hidden; color:#e7edf7; }}
button, input, select {{ font:inherit; }}
button:focus-visible, input:focus-visible, select:focus-visible {{ outline:3px solid rgba(56,189,248,.42); outline-offset:2px; }}
#canvas-container {{ position:fixed; top:0; left:400px; right:0; bottom:0; cursor:grab; transition:left .18s ease; }}
#canvas-container:active {{ cursor:grabbing; }}
#mainCanvas {{ display:block; width:100%; height:100%; }}
#sidebar {{
  position:fixed; top:0; left:0; bottom:0; width:400px; border-right:1px solid #0f3460; border-left:none;
  background:rgba(10,20,38,.97); color:#e7edf7;
  display:flex; flex-direction:column;
  box-shadow:12px 0 36px rgba(0,0,0,.24);
}}
#sidebar-header {{
  padding:22px 20px 16px; background:linear-gradient(145deg,#10264a,#0a1426 72%);
  border-bottom:1px solid #0f3460;
}}
#sidebar-header h1 {{ font-size:20px; color:#f8fbff; font-weight:800; letter-spacing:-.04em; }}
#sidebar-header p {{ font-size:12px; color:#91a4c3; margin-top:7px; }}
#search-wrap {{ padding:10px 14px; border-bottom:1px solid #0f3460; }}
#search-input {{
  width:100%; padding:8px 12px;
  background:#0f3460; border:1px solid #1a4a80;
  border-radius:20px; color:#e0e0e0;
  font-size:13px; font-family:'Malgun Gothic',sans-serif; outline:none;
}}
#search-input::placeholder {{ color:#556; }}
#search-input:focus {{ border-color:#e94560; }}
#legend-area {{ padding:10px 14px; border-bottom:1px solid #0f3460; overflow-y:auto; max-height:260px; }}
#legend-area h3 {{ font-size:10px; color:#889; letter-spacing:1px; margin-bottom:8px; text-transform:uppercase; }}
.legend-row {{ display:flex; align-items:center; gap:8px; margin-bottom:6px; font-size:12px; color:#ccc; }}
.legend-line {{ height:4px; width:28px; border-radius:2px; flex-shrink:0; }}



#info-name {{ font-size:20px; font-weight:700; color:#e94560; margin-bottom:8px; }}
#info-badges {{ display:flex; flex-wrap:wrap; gap:5px; margin-bottom:8px; }}
.badge {{ padding:3px 8px; border-radius:10px; font-size:11px; font-weight:600; color:white; }}
#info-coord {{ font-size:11px; color:#556; }}
#controls {{ padding:10px 14px; border-top:1px solid #0f3460; display:flex; gap:6px; flex-wrap:wrap; }}
.ctrl-btn {{
  flex:1; min-width:70px;
  padding:7px 4px; background:#0f3460;
  border:1px solid #1a4a80; border-radius:6px;
  color:#aab; font-size:11px; cursor:pointer;
  transition:all 0.2s; font-family:'Malgun Gothic',sans-serif;
}}
.ctrl-btn:hover {{ background:#1a4a80; color:white; }}
.ctrl-btn.active {{ background:#e94560; color:white; border-color:#e94560; }}
#zoom-badge {{
  position:fixed; bottom:24px; right:24px;
  background:rgba(22,33,62,0.9); color:#aab;
  border:1px solid #0f3460; padding:4px 10px;
  border-radius:12px; font-size:11px; z-index:50;
}}
#map-controls {{ position:fixed; right:24px; top:24px; z-index:60; display:flex; gap:8px; padding:8px; border:1px solid rgba(148,163,184,.2); border-radius:14px; background:rgba(8,16,31,.82); box-shadow:0 10px 30px rgba(0,0,0,.25); backdrop-filter:blur(10px); }}
.map-btn {{ min-width:38px; height:38px; padding:0 11px; border:1px solid rgba(148,163,184,.2); border-radius:9px; background:#14233d; color:#d9e7f8; cursor:pointer; font-size:12px; font-weight:700; }}
.map-btn:hover {{ background:#1d3558; border-color:#38bdf8; }}
.map-btn.active {{ background:#0ea5e9; border-color:#38bdf8; color:#fff; }}
#tooltip {{
  position:fixed; pointer-events:none;
  background:rgba(22,33,62,0.95); border:1px solid #e94560;
  color:#e0e0e0; padding:6px 10px; border-radius:6px;
  font-size:12px; white-space:nowrap; display:none; z-index:100;
}}
#search-results {{
  position:absolute; top:100%; left:0; right:0;
  background:#0f3460; border:1px solid #1a4a80;
  border-radius:0 0 8px 8px; max-height:200px; overflow-y:auto;
  display:none; z-index:200;
}}
.search-result-item {{
  padding:7px 12px; cursor:pointer; font-size:12px; color:#ccc;
  border-bottom:1px solid #1a4a8044;
}}
.search-result-item:hover {{ background:#1a4a80; color:white; }}
#search-wrap {{ position:relative; }}
</style>
</head>
<body>

<div id="canvas-container"><canvas id="mainCanvas"></canvas></div>

<div id="map-controls" aria-label="지도 제어">
  <button class="map-btn" id="zoom-out-btn" title="축소" aria-label="지도 축소">−</button>
  <button class="map-btn" id="fit-btn" title="전체 보기">전체</button>
  <button class="map-btn" id="zoom-in-btn" title="확대" aria-label="지도 확대">＋</button>
  <button class="map-btn active" id="label-btn" title="역 이름 표시">역명</button>
  <button class="map-btn" id="lod-btn" title="주요 노선만 표시">간략</button>
</div>

<div id="sidebar">
  <div id="sidebar-header">
    <h1>🚄 전국 철도 노선도</h1>
    <p id="stats-text">Loading...</p>
  </div>
  <div id="search-wrap">
    <input type="text" id="search-input" placeholder="🔍 역 이름 검색..." autocomplete="off"/>
    <div id="search-results"></div>
  </div>
</div>
<div id="zoom-badge">100%</div>
<div id="tooltip"></div>

<script>
const NODES = {nodes_json};
const EDGES = {edges_json};

const TIMETABLE = {timetable_json};

const LINE_CFG = {linecfg_json};
const PRIORITY = {priority_json};

// 좌표 맵 구축
const nodePos = {{}};
const nodeData = {{}};
NODES.forEach(n => {{
  nodeData[n.id] = n;
  if (n.hasCoord && n.lat && n.lon) {{
    const lat=n.lat, lon=n.lon;
    const x = (lon - 125.5) / (130.2 - 125.5) * 2000;
    const y = (38.8 - lat) / (38.8 - 33.0) * 2400;
    nodePos[n.id] = {{x:Math.round(x*10)/10, y:Math.round(y*10)/10}};
  }}
}});

// 좌표 없는 역 → 인접역 기준 추정
let changed = true;
while (changed) {{
  changed = false;
  EDGES.forEach(e => {{
    const pA = nodePos[e.from], pB = nodePos[e.to];
    if (pA && !pB) {{ nodePos[e.to] = {{x:pA.x+15, y:pA.y+15}}; changed=true; }}
    else if (pB && !pA) {{ nodePos[e.from] = {{x:pB.x-15, y:pB.y-15}}; changed=true; }}
  }});
}}

// 범례
// legendEl removed
// legend area removed

document.getElementById('stats-text').textContent =
  `${{NODES.length}}개 역 · ${{EDGES.length}}개 구간 (무정차 포함)`;

// ── Canvas 엔진 ──
const canvas = document.getElementById('mainCanvas');
const ctx = canvas.getContext('2d');
const container = document.getElementById('canvas-container');

let camX=0, camY=0, scale=1;
let isDragging=false, dragStart={{x:0,y:0}}, camStart={{x:0,y:0}};
let highlightNode=null, searchQuery='';
let simpleLOD=false, showLabels=true;
let highlightedPathNodes=new Set(), highlightedPathEdges=new Set();

function resize() {{
  canvas.width = container.clientWidth;
  canvas.height = container.clientHeight;
  draw();
}}

function worldToScreen(wx, wy) {{
  return [wx*scale + canvas.width/2 + camX, wy*scale + canvas.height/2 + camY];
}}
function screenToWorld(sx, sy) {{
  return [(sx - canvas.width/2 - camX)/scale, (sy - canvas.height/2 - camY)/scale];
}}

function draw() {{
  ctx.clearRect(0, 0, canvas.width, canvas.height);

  // 배경
  const grad = ctx.createLinearGradient(0,0,0,canvas.height);
  grad.addColorStop(0,'#1a1a2e'); grad.addColorStop(1,'#0d0d1a');
  ctx.fillStyle = grad;
  ctx.fillRect(0,0,canvas.width,canvas.height);

  const showAllEdges = !simpleLOD || scale >= 0.35;
  const showAllNodes = scale >= 0.6;
  const showMajorLabel = scale >= 0.35 && showLabels;
  const showMinorLabel = scale >= 1.0 && showLabels;
  const showHubLabel = scale >= 0.2 && showLabels;

  // 구간 엣지 (덜 중요한 노선 먼저)
  const drawOrder = [...PRIORITY].reverse();
  drawOrder.forEach(lineType => {{
    const cfg = LINE_CFG[lineType];
    if (!cfg) return;
    const isKTX = lineType.startsWith('KTX') || lineType==='SRT';
    if (!showAllEdges && !isKTX) return;

    EDGES.forEach(e => {{
      if (!e.lines.includes(lineType)) return;
      const pA=nodePos[e.from], pB=nodePos[e.to];
      if (!pA||!pB) return;
      const [sx1,sy1]=worldToScreen(pA.x,pA.y);
      const [sx2,sy2]=worldToScreen(pB.x,pB.y);
      if (Math.max(sx1,sx2)<-50||Math.min(sx1,sx2)>canvas.width+50) return;
      if (Math.max(sy1,sy2)<-50||Math.min(sy1,sy2)>canvas.height+50) return;

      // 평행 오프셋
      const lineIdx = e.lines.indexOf(lineType);
      const totalLines = e.lines.length;
      let offset = 0;
      if (totalLines > 1) {{
        const spread = Math.min(cfg.width * scale * 1.4, 7);
        offset = (lineIdx - (totalLines-1)/2) * spread;
      }}
      const dx=sx2-sx1, dy=sy2-sy1, len=Math.sqrt(dx*dx+dy*dy);
      const nx = len>0 ? -dy/len*offset : 0;
      const ny = len>0 ?  dx/len*offset : 0;

      const edgeId = e.from < e.to ? `${{e.from}}-${{e.to}}` : `${{e.to}}-${{e.from}}`;
      const isPath = highlightedPathEdges.has(edgeId);
      const isHL = isPath || (highlightNode && (e.from===highlightNode||e.to===highlightNode));
      const isSQ = searchQuery && (e.from.includes(searchQuery)||e.to.includes(searchQuery));

      ctx.save();
      ctx.beginPath();
      ctx.moveTo(sx1+nx, sy1+ny);
      ctx.lineTo(sx2+nx, sy2+ny);
      ctx.strokeStyle = isPath ? '#67e8f9' : cfg.color;
      ctx.lineWidth = Math.max(cfg.width*scale*(isHL?2.35:1), isKTX?1.5:0.8);
      ctx.globalAlpha = highlightedPathEdges.size?(isPath?1:.1):searchQuery?(isSQ?0.95:0.08):(highlightNode?(isHL?0.95:0.15):0.85);
      if (cfg.dash&&cfg.dash.length) ctx.setLineDash(cfg.dash.map(d=>d*Math.max(scale,0.4)));
      else ctx.setLineDash([]);
      ctx.stroke();
      ctx.restore();
    }});
  }});

  // 역 노드
  NODES.forEach(n => {{
    const pos = nodePos[n.id];
    if (!pos) return;
    const isMinor = !n.isMajor && !n.isHub;
    if (!showAllNodes && isMinor) return;

    const [sx,sy] = worldToScreen(pos.x,pos.y);
    if (sx<-40||sx>canvas.width+40||sy<-40||sy>canvas.height+40) return;

    const cfg = LINE_CFG[n.primary]||LINE_CFG['무궁화'];
    const isSelected = n.id===highlightNode;
    const isPath = highlightedPathNodes.has(n.id);
    const isSearched = searchQuery && n.id.includes(searchQuery);
    const baseR = n.isHub?10:n.isMajor?7:4;
    const r = Math.max(baseR*Math.min(scale,1.5), n.isHub?4:n.isMajor?2.5:1.5);

    const alpha = highlightedPathNodes.size?(isPath?1:.14):searchQuery?(isSearched?1:0.1):(highlightNode?(isSelected?1:0.24):1);

    ctx.save();
    ctx.globalAlpha = alpha;

    if (isSelected || isPath) {{
      ctx.beginPath(); ctx.arc(sx,sy,r+6,0,Math.PI*2);
      ctx.strokeStyle=isPath?'#67e8f9':'#FFD700'; ctx.lineWidth=2.5; ctx.stroke();
    }}

    ctx.beginPath(); ctx.arc(sx,sy,r,0,Math.PI*2);
    if (n.isHub || n.isMajor) {{
      ctx.fillStyle='#1a1a2e'; ctx.fill();
      ctx.strokeStyle = isPath?'#67e8f9':isSelected?'#FFD700':cfg.color;
      ctx.lineWidth = Math.max(r*0.55,2);
      ctx.stroke();
    }} else {{
      ctx.fillStyle = cfg.color; ctx.fill();
    }}

    const showLabel = n.isHub ? showHubLabel : n.isMajor ? showMajorLabel : showMinorLabel;
    if (showLabel || isSelected || isSearched) {{
      const fs = n.isHub?Math.max(14*scale,11):n.isMajor?Math.max(12*scale,9):Math.max(10*scale,8);
      ctx.font = `${{n.isHub?'700':n.isMajor?'600':'400'}} ${{fs}}px 'Malgun Gothic',sans-serif`;
      ctx.textAlign='center'; ctx.textBaseline='top';
      const lx=sx, ly=sy+r+2;
      ctx.strokeStyle='#0d0d1a'; ctx.lineWidth=3.5; ctx.lineJoin='round';
      ctx.strokeText(n.id,lx,ly);
      ctx.fillStyle = isPath?'#cffafe':isSelected?'#FFD700':isSearched?'#FFD700':n.isHub?'#fff':n.isMajor?'#ddd':'#aaa';
      ctx.fillText(n.id,lx,ly);
    }}
    ctx.restore();
  }});

  document.getElementById('zoom-badge').textContent = Math.round(scale*100)+'%';
}}

// 드래그
container.addEventListener('mousedown', e => {{
  isDragging=true;
  dragStart={{x:e.clientX,y:e.clientY}};
  camStart={{x:camX,y:camY}};
}});
container.addEventListener('mousemove', e => {{
  if (isDragging) {{
    camX=camStart.x+(e.clientX-dragStart.x);
    camY=camStart.y+(e.clientY-dragStart.y);
    draw(); return;
  }}
  // 호버 툴팁
  const [wx,wy]=screenToWorld(e.clientX,e.clientY);
  let found=null, bestD=Infinity;
  NODES.forEach(n => {{
    const pos=nodePos[n.id]; if(!pos) return;
    const r=Math.max((n.isHub?10:n.isMajor?7:4)/scale,10);
    const d=Math.sqrt((pos.x-wx)**2+(pos.y-wy)**2);
    if(d<r&&d<bestD){{found=n;bestD=d;}}
  }});
  const tip=document.getElementById('tooltip');
  if(found){{
    tip.style.display='block';
    tip.style.left=(e.clientX+14)+'px'; tip.style.top=(e.clientY-10)+'px';
    const cfg=LINE_CFG[found.primary]||{{}};
    tip.innerHTML=`<strong style="color:${{cfg.color||'#fff'}}">${{found.id}}역</strong><br>${{found.types.map(t=>LINE_CFG[t]?.label||t).join(' · ')}}<br><span style="color:#91a4c3">인접 ${{found.degree}}개 · 운행 ${{found.trainCount}}편</span>`;
    canvas.style.cursor='pointer';
  }} else {{
    tip.style.display='none';
    canvas.style.cursor=isDragging?'grabbing':'grab';
  }}
}});
container.addEventListener('mouseup', e => {{
  if(!isDragging) return;
  const moved=Math.abs(e.clientX-dragStart.x)+Math.abs(e.clientY-dragStart.y);
  isDragging=false;
  if(moved<6) {{
    const [wx,wy]=screenToWorld(e.clientX,e.clientY);
    let found=null, bestD=Infinity;
    NODES.forEach(n=>{{
      const pos=nodePos[n.id]; if(!pos) return;
      const r=Math.max((n.isHub?10:n.isMajor?7:4)/scale,12);
      const d=Math.sqrt((pos.x-wx)**2+(pos.y-wy)**2);
      if(d<r&&d<bestD){{found=n;bestD=d;}}
    }});
    highlightNode = found?.id||null;
    
    
    if(found){{
      showStationDetail(found);
      const pos=nodePos[found.id];
      if(pos) {{
        const [tx,ty]=worldToScreen(pos.x,pos.y);
        animateCam(camX+(canvas.width*0.35-tx), camY+(canvas.height*0.5-ty));
      }}
    }} else {{
      
    }}
    draw();
  }}
}});
container.addEventListener('mouseleave',()=>{{
  isDragging=false;
  document.getElementById('tooltip').style.display='none';
}});
container.addEventListener('wheel',e=>{{
  e.preventDefault();
  const f=e.deltaY>0?0.85:1.18;
  const ns=Math.max(0.12,Math.min(6,scale*f));
  const mx=e.clientX, my=e.clientY;
  camX=mx-(mx-camX)*ns/scale;
  camY=my-(my-camY)*ns/scale;
  scale=ns; draw();
}},{{passive:false}});

// 검색
const searchInput=document.getElementById('search-input');
const searchResults=document.getElementById('search-results');
searchInput.addEventListener('input',e=>{{
  searchQuery=e.target.value.trim();
  highlightNode=null;
  
  

  if(searchQuery.length>=1){{
    const matched=NODES.filter(n=>n.id.includes(searchQuery)&&nodePos[n.id]).slice(0,12);
    searchResults.innerHTML='';
    matched.forEach(n=>{{
      const item=document.createElement('div');
      item.className='search-result-item';
      const cfg=LINE_CFG[n.primary]||{{}};
      item.innerHTML=`<span style="color:${{cfg.color||'#fff'}};font-weight:600">${{n.id}}</span> <span style="color:#556;font-size:10px">${{n.types.map(t=>LINE_CFG[t]?.label||t).join(', ')}}</span>`;
      item.onclick=()=>selectStation(n);
      searchResults.appendChild(item);
    }});
    searchResults.style.display=matched.length?'block':'none';
  }} else {{
    searchResults.style.display='none';
  }}
  draw();
}});
document.addEventListener('click',e=>{{
  if(!e.target.closest('#search-wrap')) searchResults.style.display='none';
}});

function selectStation(n) {{
  highlightNode=n.id;
  searchQuery='';
  searchInput.value=n.id;
  searchResults.style.display='none';
  
  showStationDetail(n);
  
  const pos=nodePos[n.id];
  if(pos){{
    if(scale<1.5) scale=2;
    const [tx,ty]=worldToScreen(pos.x,pos.y);
    animateCam(camX+(canvas.width*0.35-tx), camY+(canvas.height*0.5-ty));
  }}
  draw();
}}

function animateCam(tx,ty){{
  const sx=camX,sy=camY,dur=500,start=performance.now();
  function step(now){{
    const t=Math.min((now-start)/dur,1);
    const e=t<0.5?2*t*t:-1+(4-2*t)*t;
    camX=sx+(tx-sx)*e; camY=sy+(ty-sy)*e;
    draw(); if(t<1) requestAnimationFrame(step);
  }}
  requestAnimationFrame(step);
}}

function fitAll(){{
  const ps=Object.values(nodePos);
  if(!ps.length)return;
  const xs=ps.map(p=>p.x),ys=ps.map(p=>p.y);
  const minX=Math.min(...xs),maxX=Math.max(...xs);
  const minY=Math.min(...ys),maxY=Math.max(...ys);
  const ns=Math.min(canvas.width*0.88/(maxX-minX),canvas.height*0.88/(maxY-minY));
  scale=ns;
  const hg = nodePos['황간'];
  const targetX = hg ? hg.x : (minX+maxX)/2;
  const targetY = hg ? hg.y : (minY+maxY)/2;
  animateCam(-targetX * ns, -targetY * ns);
}}

function toggleLOD(){{
  simpleLOD=!simpleLOD;
  const btn=document.getElementById('lod-btn'); if(!btn)return;
  btn.classList.toggle('active',simpleLOD);
  btn.textContent='간략';
  draw();
}}

function toggleLabels(){{
  showLabels=!showLabels;
  const btn=document.getElementById('label-btn'); if(!btn)return;
  btn.classList.toggle('active',!showLabels);
  btn.textContent='역명';
  draw();
}}

function zoomAtCenter(factor) {{
  scale=Math.max(.12,Math.min(6,scale*factor));
  draw();
}}

document.getElementById('fit-btn').addEventListener('click',fitAll);
document.getElementById('zoom-in-btn').addEventListener('click',()=>zoomAtCenter(1.22));
document.getElementById('zoom-out-btn').addEventListener('click',()=>zoomAtCenter(.82));
document.getElementById('label-btn').addEventListener('click',toggleLabels);
document.getElementById('lod-btn').addEventListener('click',toggleLOD);

window.addEventListener('resize',resize);
resize();
setTimeout(fitAll,150);
</script>
</body>
</html>"""

custom_css = """
/* Routing UI CSS */
#route-panel {
  padding: 14px;
  border-bottom: 1px solid #0f3460;
  flex: 1;
  overflow-y: auto;
}
#route-panel h3 {
  font-size: 14px; color: #4facfe; margin-bottom: 12px;
}
.waypoint-wrap {
  display: flex; flex-direction: column; gap: 6px; margin-bottom: 10px;
}
.waypoint-input {
  width: 100%; padding: 8px 12px;
  background: #0f3460; border: 1px solid #1a4a80;
  border-radius: 6px; color: #e0e0e0;
  font-size: 13px; outline: none;
}
.waypoint-input:focus { border-color: #4facfe; }
.waypoint-row { display: flex; gap: 6px; align-items: center; }
.remove-btn { 
  background: none; border: none; color: #e94560; cursor: pointer;
  font-size: 16px; padding: 0 4px;
}
.add-btn {
  width: 100%; padding: 6px; background: transparent;
  border: 1px dashed #1a4a80; color: #889; border-radius: 6px;
  cursor: pointer; font-size: 12px; margin-bottom: 10px;
}
.add-btn:hover { border-color: #4facfe; color: #4facfe; }
.route-options {
  display: flex; gap: 8px; margin-bottom: 12px;
}
.route-options input, .route-options select {
  flex: 1; padding: 6px; background: #0f3460; border: 1px solid #1a4a80;
  border-radius: 4px; color: #e0e0e0; font-size: 12px; outline: none;
}
.find-btn {
  width: 100%; padding: 10px; background: linear-gradient(135deg, #4facfe, #00f2fe);
  border: none; border-radius: 20px; color: #fff; font-weight: bold;
  cursor: pointer; box-shadow: 0 4px 15px rgba(79,172,254,0.3);
}
.find-btn:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(79,172,254,0.4); }
#route-result {
  margin-top: 14px; font-size: 12px; color: #ccc; 
}
.route-step {
  margin-bottom: 8px; padding-left: 12px; border-left: 2px solid #4facfe;
}
.route-step-time { color: #889; font-size: 10px; }
.route-step-train { color: #e94560; font-weight: bold; }
.route-error { color: #e94560; text-align: center; padding: 10px 0; }

/* Readability overrides */
#search-wrap { order: 1; padding: 12px 16px; background: #0b172a; }
#search-input { min-height: 42px; padding: 10px 14px; border-radius: 10px; background: #111f36; border-color: #263a59; color: #f8fbff; }
#search-input::placeholder { color: #7183a1; }
#search-results { top: calc(100% - 8px); left: 16px; right: 16px; border-radius: 10px; background: #13233c; box-shadow: 0 18px 38px rgba(0,0,0,.35); }
.search-result-item { padding: 10px 12px; }
#route-panel { order: 2; padding: 16px; border-bottom: 0; flex: 1; scrollbar-width: thin; scrollbar-color: #2a4267 transparent; }
#route-panel h3 { font-size: 13px; color: #91a4c3; margin-bottom: 10px; letter-spacing: .02em; }
.waypoint-wrap { position: relative; gap: 8px; }
.waypoint-input { min-height: 44px; padding: 10px 12px 10px 38px; background: #111f36; border-color: #263a59; border-radius: 10px; color: #f8fbff; font-size: 14px; }
.waypoint-input:focus { border-color: #38bdf8; background:#142642; }
.waypoint-row { position:relative; }
.waypoint-row::before { position:absolute; left:14px; z-index:1; width:9px; height:9px; border-radius:50%; content:''; background:#38bdf8; box-shadow:0 0 0 4px rgba(56,189,248,.12); }
.waypoint-row:nth-child(2)::before { background:#f472b6; box-shadow:0 0 0 4px rgba(244,114,182,.12); }
.remove-btn { width:32px; height:32px; flex:0 0 32px; background:#26182a; border:1px solid #4c284f; border-radius:8px; color:#fb7185; }
.route-tools { display:grid; grid-template-columns:1fr auto; gap:8px; margin-bottom:10px; }
.add-btn { min-height:36px; margin-bottom:0; padding:7px 10px; background:#0e1a2e; border-color:#314665; color:#9db0cc; border-radius:9px; }
.swap-btn { width:38px; border:1px solid #314665; border-radius:9px; background:#14233d; color:#c7d5e8; cursor:pointer; font-size:16px; }
.route-options { display:grid; grid-template-columns:1fr 1fr; gap:8px; }
.route-options input, .route-options select { width:100%; min-height:40px; padding:8px 10px; background:#111f36; border-color:#263a59; border-radius:9px; color:#e7edf7; }
.find-btn { min-height:42px; padding:10px 14px; background:linear-gradient(135deg,#0284c7,#06b6d4); border-radius:10px; font-weight:800; box-shadow:0 8px 22px rgba(14,165,233,.18); }
#reset-route-btn { min-height:42px; border-radius:10px !important; background:#14233d !important; border-color:#314665 !important; }
#route-result { color:#ced8e8; }
.route-step { padding:10px 12px !important; border-radius:0 9px 9px 0; background:#101d32; }
.route-error { color:#fb7185; }
#station-detail { margin-top:16px; padding:15px; border:1px solid #263a59; border-radius:13px; background:linear-gradient(145deg,#111f36,#0c1729); }
#station-detail.empty { color:#7f91ad; }
.station-kicker { color:#38bdf8; font-size:10px; font-weight:800; letter-spacing:.12em; text-transform:uppercase; }
#station-detail h2 { margin-top:5px; color:#fff; font-size:21px; letter-spacing:-.04em; }
.station-sub { margin-top:5px; color:#8fa2bf; font-size:11px; line-height:1.55; }
.station-metrics { display:grid; grid-template-columns:repeat(3,1fr); gap:7px; margin-top:13px; }
.station-metric { padding:9px 6px; border-radius:9px; background:#0a1527; text-align:center; }
.station-metric strong { display:block; color:#f8fbff; font-size:14px; }
.station-metric span { display:block; margin-top:3px; color:#7183a1; font-size:9px; }
.station-badges { display:flex; flex-wrap:wrap; gap:5px; margin-top:12px; }
.station-neighbors { display:flex; flex-wrap:wrap; gap:6px; margin-top:11px; }
.neighbor-chip { padding:6px 9px; border:1px solid #304766; border-radius:999px; background:#14233d; color:#c9d7e9; cursor:pointer; font-size:11px; }
.neighbor-chip:hover { border-color:#38bdf8; color:#fff; }
.badge { padding:4px 8px; border-radius:999px; font-size:10px; }
@media (max-width:760px) {
  #canvas-container { left:0 !important; bottom:46vh; }
  #sidebar { top:auto; width:100% !important; height:46vh; border-right:0; border-top:1px solid #263a59; border-radius:18px 18px 0 0; }
  #sidebar-header { padding:12px 16px 10px; }
  #sidebar-header h1 { font-size:17px; }
  #sidebar-header p { margin-top:3px; }
  #search-wrap { padding:9px 12px; }
  #route-panel { padding:12px; }
  #resizer { display:none; }
  #map-controls { top:12px; right:12px; gap:5px; padding:6px; }
  .map-btn { min-width:34px; height:34px; padding:0 8px; }
  #zoom-badge { right:14px; bottom:calc(46vh + 12px); }
  #tooltip { display:none !important; }
}
"""

ui_html = """
  <div id="route-panel">
    <h3>경로 검색</h3>
    <div id="waypoints" class="waypoint-wrap">
      <div class="waypoint-row">
        <input type="text" class="waypoint-input" placeholder="출발역" list="station-options" autocomplete="off">
      </div>
      <div class="waypoint-row">
        <input type="text" class="waypoint-input" placeholder="도착역" list="station-options" autocomplete="off">
      </div>
    </div>
    <datalist id="station-options"></datalist>
    <div class="route-tools">
      <button id="add-waypoint-btn" class="add-btn">+ 경유지 추가</button>
      <button id="swap-route-btn" class="swap-btn" title="출발역과 도착역 바꾸기" aria-label="출발역과 도착역 바꾸기">⇅</button>
    </div>
    
    <div class="route-options">
      <input type="time" id="route-time" value="05:00">
      <select id="route-mode">
        <option value="TIME">최소 시간</option>
        <option value="TRANSFER">최소 환승</option>
      </select>
    </div>
    
    <div class="btn-group" style="display:flex; gap:8px;">
      <button id="find-route-btn" class="find-btn" style="flex:2;">경로 찾기</button>
      <button id="reset-route-btn" class="reset-btn" style="flex:1; background:#16213e; color:#e0e0e0; border:1px solid #1a4a80; border-radius:20px; font-weight:bold; cursor:pointer;">초기화</button>
    </div>
    <div id="route-result"></div>
    <section id="station-detail" class="empty" aria-live="polite">
      <div class="station-kicker">Station details</div>
      <h2>역을 선택하세요</h2>
      <p class="station-sub">지도에서 역을 클릭하거나 위 검색창에서 역 이름을 찾아보세요.</p>
    </section>
  </div>
"""

dijkstra_js = """
// ----------------------------------------------------
// 🚂 Javascript Dijkstra Routing Algorithm
// ----------------------------------------------------
let stationDepartures = {};
const MAX_HOURS = 168;
const TRANSFER_MINS = 10;

function buildRoutingGraph() {
  stationDepartures = {};
  for (const trip of TIMETABLE) {
    const tno = trip[0];
    const ttype = trip[1];
    const stops = trip[2];
    for (let i = 0; i < stops.length - 1; i++) {
      const st = stops[i][0];
      const t = stops[i][1];
      if (!stationDepartures[st]) stationDepartures[st] = [];
      stationDepartures[st].push({
        dep_time: t,
        trip_no: tno,
        trip_type: ttype,
        stops: stops,
        stop_idx: i
      });
    }
  }
}

function minsToTime(m) {
  const day = Math.floor(m / (24 * 60));
  const rem = m % (24 * 60);
  const h = Math.floor(rem / 60);
  const min = rem % 60;
  const dStr = day > 0 ? `(+${day}일) ` : '';
  return `${dStr}${h.toString().padStart(2,'0')}:${min.toString().padStart(2,'0')}`;
}

class PriorityQueue {
  constructor(compare) { this.data = []; this.compare = compare; }
  push(val) {
    this.data.push(val);
    this.data.sort(this.compare);
  }
  pop() { return this.data.shift(); }
  isEmpty() { return this.data.length === 0; }
}

function runDijkstra(startStation, startTimeMins, targetStation, mode) {
  const pq = new PriorityQueue((a, b) => {
    if (a.c1 !== b.c1) return a.c1 - b.c1;
    if (a.c2 !== b.c2) return a.c2 - b.c2;
    return a.seq - b.seq;
  });
  
  const visited = new Map();
  const bestKnown = new Map();
  
  const initC1 = mode === 'TIME' ? startTimeMins : 0;
  const initC2 = mode === 'TIME' ? 0 : startTimeMins;
  
  bestKnown.set(`${startStation}_null`, [initC1, initC2]);
  
  let seqCounter = 0;
  pq.push({
    c1: initC1, c2: initC2, seq: seqCounter++,
    curTime: startTimeMins, curStation: startStation, curTrain: null, path: []
  });
  
  let bestArrival = null;
  let bestPath = null;
  const maxTimeLimit = startTimeMins + MAX_HOURS * 60;
  
  while (!pq.isEmpty()) {
    const { c1, c2, curTime, curStation, curTrain, path } = pq.pop();
    
    if (curTime > maxTimeLimit) continue;
    
    if (curStation === targetStation) {
      if (!bestArrival || (c1 < bestArrival[0] || (c1 === bestArrival[0] && c2 < bestArrival[1]))) {
        bestArrival = [c1, c2];
        bestPath = path;
        if (mode === 'TIME') break;
      }
      continue;
    }
    
    const stateKey = `${curStation}_${curTrain}`;
    if (visited.has(stateKey)) {
      const v = visited.get(stateKey);
      if (v[0] <= c1 && v[1] <= c2) continue;
    }
    visited.set(stateKey, [c1, c2]);
    
    const deps = stationDepartures[curStation] || [];
    for (const dep of deps) {
      const isSameTrain = (curTrain === dep.trip_no);
      let reqTime = curTime;
      if (!isSameTrain && curTrain !== null) {
        reqTime += TRANSFER_MINS;
      }
      
      const timeOfDay = reqTime % 1440;
      let actualDepTime = Math.floor(reqTime / 1440) * 1440 + dep.dep_time;
      if (dep.dep_time < timeOfDay) {
        actualDepTime += 1440;
      }
      
      if (actualDepTime - curTime > MAX_HOURS * 60) continue;
      
      if (actualDepTime >= reqTime) {
        for (let nextIdx = dep.stop_idx + 1; nextIdx < dep.stops.length; nextIdx++) {
          const nextStation = dep.stops[nextIdx][0];
          const arrTimeRaw = dep.stops[nextIdx][1];
          
          const arrTimeOfDay = actualDepTime % 1440;
          let arrT = Math.floor(actualDepTime / 1440) * 1440 + arrTimeRaw;
          if (arrTimeRaw < arrTimeOfDay) {
            arrT += 1440;
          }
          
          let transfers = mode === 'TIME' ? c2 : c1;
          if (!isSameTrain && curTrain !== null) transfers++;
          
          const newC1 = mode === 'TIME' ? arrT : transfers;
          const newC2 = mode === 'TIME' ? transfers : arrT;
          
          const nextState = `${nextStation}_${dep.trip_no}`;
          if (bestKnown.has(nextState)) {
            const bk = bestKnown.get(nextState);
            if (bk[0] <= newC1 && bk[1] <= newC2) continue;
          }
          bestKnown.set(nextState, [newC1, newC2]);
          
          const newPath = [...path];
          if (!isSameTrain) {
            newPath.push({
              trip_no: dep.trip_no, trip_type: dep.trip_type,
              b_st: curStation, b_t: actualDepTime, b_idx: dep.stop_idx,
              a_st: nextStation, a_t: arrT, a_idx: nextIdx, stops: dep.stops
            });
          } else {
            const lastLeg = newPath[newPath.length - 1];
            newPath[newPath.length - 1] = {
              ...lastLeg,
              a_st: nextStation, a_t: arrT, a_idx: nextIdx
            };
          }
          
          pq.push({
            c1: newC1, c2: newC2, seq: seqCounter++,
            curTime: arrT, curStation: nextStation, curTrain: dep.trip_no, path: newPath
          });
        }
      }
    }
  }
  
  return { cost: bestArrival, path: bestPath };
}

function findFixedRoute(waypoints, startTimeStr, mode) {
  const parts = startTimeStr.split(':');
  let currentTime = parseInt(parts[0]) * 60 + parseInt(parts[1]);
  
  let currentStation = waypoints[0];
  let fullPath = [];
  let totalTransfers = 0;
  
  for (let i = 1; i < waypoints.length; i++) {
    const nextStation = waypoints[i];
    const res = runDijkstra(currentStation, currentTime, nextStation, mode);
    if (!res.cost) {
      return { error: `${currentStation} ➔ ${nextStation} 구간 연결 노선이 없습니다.` };
    }
    
    let transfers = mode === 'TIME' ? res.cost[1] : res.cost[0];
    let arrTime = mode === 'TIME' ? res.cost[0] : res.cost[1];
    
    totalTransfers += transfers;
    currentTime = arrTime;
    currentStation = nextStation;
    fullPath = fullPath.concat(res.path);
  }
  
  return { path: fullPath, totalTransfers, finalTime: currentTime };
}

let physicalAdj = {};
function buildPhysicalGraph() {
  physicalAdj = {};
  EDGES.forEach(e => {
    const u = e.from, v = e.to;
    if (!physicalAdj[u]) physicalAdj[u] = [];
    if (!physicalAdj[v]) physicalAdj[v] = [];
    const edgeId = u < v ? `${u}-${v}` : `${v}-${u}`;
    physicalAdj[u].push({ nxt: v, edgeId: edgeId });
    physicalAdj[v].push({ nxt: u, edgeId: edgeId });
  });
}

function getPhysicalPath(start, target) {
  if (start === target) return { nodes: [start], edges: [] };
  const queue = [[start, [start], []]];
  const visited = new Set([start]);
  while (queue.length > 0) {
    const [curr, nodePath, edgePath] = queue.shift();
    if (curr === target) {
      return { nodes: nodePath, edges: edgePath };
    }
    const neighbors = physicalAdj[curr] || [];
    for (const { nxt, edgeId } of neighbors) {
      if (!visited.has(nxt)) {
        visited.add(nxt);
        queue.push([nxt, [...nodePath, nxt], [...edgePath, edgeId]]);
      }
    }
  }
  const fallbackEdge = start < target ? `${start}-${target}` : `${target}-${start}`;
  return { nodes: [start, target], edges: [fallbackEdge] };
}

function escapeHtml(value) {
  return String(value).replace(/[&<>'"]/g, ch => ({
    '&':'&amp;', '<':'&lt;', '>':'&gt;', "'":'&#39;', '"':'&quot;'
  })[ch]);
}

function formatClock(minute) {
  if (minute === null || minute === undefined) return '—';
  const h = Math.floor(minute / 60) % 24;
  const m = minute % 60;
  return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
}

function showStationDetail(station) {
  const panel = document.getElementById('station-detail');
  if (!panel || !station) return;
  panel.classList.remove('empty');
  const badges = station.types.map(type => {
    const cfg = LINE_CFG[type] || { color:'#475569', label:type };
    return `<span class="badge" style="background:${cfg.color}">${escapeHtml(cfg.label)}</span>`;
  }).join('');
  const neighbors = station.neighbors.map(name =>
    `<button class="neighbor-chip" data-station="${escapeHtml(name)}">${escapeHtml(name)}</button>`
  ).join('');
  const coord = station.hasCoord
    ? `${station.lat.toFixed(4)}, ${station.lon.toFixed(4)}`
    : '인접 노드를 기준으로 지도 위치 추정';
  panel.innerHTML = `
    <div class="station-kicker">Station details</div>
    <h2>${escapeHtml(station.id)}역</h2>
    <p class="station-sub">${coord}</p>
    <div class="station-metrics">
      <div class="station-metric"><strong>${station.trainCount}</strong><span>운행 열차</span></div>
      <div class="station-metric"><strong>${formatClock(station.firstTime)}</strong><span>첫 운행</span></div>
      <div class="station-metric"><strong>${formatClock(station.lastTime)}</strong><span>마지막 운행</span></div>
    </div>
    <div class="station-badges">${badges}</div>
    <p class="station-sub" style="margin-top:12px">직접 연결된 역 ${station.degree}개</p>
    <div class="station-neighbors">${neighbors || '<span class="station-sub">연결 정보 없음</span>'}</div>`;
  panel.querySelectorAll('.neighbor-chip').forEach(button => {
    button.addEventListener('click', () => {
      const next = NODES.find(n => n.id === button.dataset.station);
      if (next) selectStation(next);
    });
  });
}

document.addEventListener('DOMContentLoaded', () => {
  buildRoutingGraph();
  buildPhysicalGraph();
  const stationOptions = document.getElementById('station-options');
  if (stationOptions) {
    stationOptions.innerHTML = NODES.filter(n => nodePos[n.id])
      .map(n => `<option value="${escapeHtml(n.id)}"></option>`).join('');
  }
  function setCurrentTime() {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, '0');
    const mm = String(now.getMinutes()).padStart(2, '0');
    const timeInput = document.getElementById('route-time');
    if (timeInput) timeInput.value = `${hh}:${mm}`;
  }
  setCurrentTime();
  
  const waypointsDiv = document.getElementById('waypoints');
  
  document.getElementById('add-waypoint-btn').addEventListener('click', () => {
    const row = document.createElement('div');
    row.className = 'waypoint-row';
    row.innerHTML = `
      <input type="text" class="waypoint-input" placeholder="경유지" list="station-options" autocomplete="off">
      <button class="remove-btn">×</button>
    `;
    row.querySelector('.remove-btn').addEventListener('click', () => row.remove());
    waypointsDiv.insertBefore(row, waypointsDiv.lastElementChild);
  });

  document.getElementById('swap-route-btn').addEventListener('click', () => {
    const inputs = document.querySelectorAll('.waypoint-input');
    if (inputs.length < 2) return;
    const first = inputs[0].value;
    inputs[0].value = inputs[inputs.length - 1].value;
    inputs[inputs.length - 1].value = first;
  });
  
  document.getElementById('find-route-btn').addEventListener('click', () => {
    const inputs = Array.from(document.querySelectorAll('.waypoint-input')).map(el => el.value.trim()).filter(v => v);
    if (inputs.length < 2) {
      alert("출발역과 도착역을 입력해주세요.");
      return;
    }
    
    for (const st of inputs) {
      if (!NODES.find(n => n.id === st)) {
        alert(`${st} 역을 지도에서 찾을 수 없습니다.`);
        return;
      }
    }
    
    const timeStr = document.getElementById('route-time').value;
    const mode = document.getElementById('route-mode').value;
    
    const res = findFixedRoute(inputs, timeStr, mode);
    const resultDiv = document.getElementById('route-result');
    
    if (res.error) {
      resultDiv.innerHTML = `<div class="route-error">❌ ${res.error}</div>`;
      highlightPath(new Set(), new Set());
      return;
    }
    
    const [startHour, startMinute] = timeStr.split(':').map(Number);
    const duration = res.finalTime - (startHour * 60 + startMinute);
    const durationText = `${Math.floor(duration / 60)}시간 ${duration % 60}분`;
    let htmlStr = `<div style="padding:12px; margin-bottom:10px; border:1px solid #285070; border-radius:10px; background:#0d2136; color:#fff;"><b>${minsToTime(res.finalTime)} 도착</b><div style="margin-top:4px; color:#91a4c3; font-size:11px">${durationText} · 환승 ${res.totalTransfers}회</div></div>`;
    
    let prevArrTime = null;
    let pathNodes = new Set();
    let pathEdges = new Set();
    
    res.path.forEach(leg => {
      if (prevArrTime !== null) {
        htmlStr += `<div class="route-step-time">환승 대기 ${leg.b_t - prevArrTime}분</div>`;
      }
      
      const isKTXNum = (String(leg.trip_no).trim().length <= 3 && !isNaN(String(leg.trip_no).trim()));
      const displayTrainType = isKTXNum ? 'KTX' : leg.trip_type;
      
      // Line border color
      let lineBorderColor = '#4facfe';
      if (displayTrainType.includes('KTX')) lineBorderColor = '#003B72';
      else if (displayTrainType.includes('SRT')) lineBorderColor = '#59233E';
      else if (displayTrainType.includes('ITX')) lineBorderColor = '#AC192D';
      else if (displayTrainType.includes('새마을')) lineBorderColor = '#005388';
      else if (displayTrainType.includes('무궁화')) lineBorderColor = '#E84E0F';

      // Bright text color for Dark Mode readability
      let trainTextColor = '#38bdf8';
      if (displayTrainType.includes('KTX')) trainTextColor = '#38bdf8';       // Bright Sky Blue
      else if (displayTrainType.includes('SRT')) trainTextColor = '#f472b6';   // Bright Pink/Magenta
      else if (displayTrainType.includes('ITX')) trainTextColor = '#f87171';   // Bright Red
      else if (displayTrainType.includes('새마을')) trainTextColor = '#38bdf8'; // Bright Blue
      else if (displayTrainType.includes('무궁화')) trainTextColor = '#fb923c'; // Bright Orange
      
      htmlStr += `<div class="route-step" style="border-left: 3px solid ${lineBorderColor}; margin-bottom:10px; padding-left:10px;">
        <div class="route-step-train" style="color:${trainTextColor}; font-size:13px; font-weight:700;">[${minsToTime(leg.b_t)}] ${leg.b_st} ➔ [${minsToTime(leg.a_t)}] ${leg.a_st}</div>
        <div class="route-step-time" style="color:#aaa; font-size:11px; margin-top:2px;">${displayTrainType} #${leg.trip_no}</div>
      </div>`;
      prevArrTime = leg.a_t;
      
      for (let i = leg.b_idx; i < leg.a_idx; i++) {
        const st1 = leg.stops[i][0];
        const st2 = leg.stops[i+1][0];
        const phys = getPhysicalPath(st1, st2);
        phys.nodes.forEach(n => pathNodes.add(n));
        phys.edges.forEach(e => pathEdges.add(e));
      }
      const directPhys = getPhysicalPath(leg.b_st, leg.a_st);
      directPhys.nodes.forEach(n => pathNodes.add(n));
      directPhys.edges.forEach(e => pathEdges.add(e));
    });
    
    resultDiv.innerHTML = htmlStr;
    highlightPath(pathNodes, pathEdges);
  });

  document.getElementById('reset-route-btn').addEventListener('click', () => {
    const waypointsDiv = document.getElementById('waypoints');
    waypointsDiv.innerHTML = `
      <div class="waypoint-row">
        <input type="text" class="waypoint-input" placeholder="출발역" list="station-options" autocomplete="off">
      </div>
      <div class="waypoint-row">
        <input type="text" class="waypoint-input" placeholder="도착역" list="station-options" autocomplete="off">
      </div>
    `;
    document.getElementById('route-result').innerHTML = '';
    pathNodes = new Set();
    pathEdges = new Set();
    highlightNode = null;
    searchQuery = '';
    const searchInput = document.getElementById('search-input');
    if (searchInput) searchInput.value = '';
    highlightPath(new Set(), new Set());
    setCurrentTime();
    fitAll();
  });
});

function highlightPath(nodesSet, edgesSet) {
  highlightedPathNodes = nodesSet || new Set();
  highlightedPathEdges = edgesSet || new Set();
  if (highlightedPathNodes.size > 1) {
    const positions = [...highlightedPathNodes].map(name => nodePos[name]).filter(Boolean);
    if (positions.length > 1) {
      const xs = positions.map(p => p.x), ys = positions.map(p => p.y);
      const minX = Math.min(...xs), maxX = Math.max(...xs);
      const minY = Math.min(...ys), maxY = Math.max(...ys);
      const width = Math.max(maxX - minX, 80), height = Math.max(maxY - minY, 80);
      scale = Math.max(.12, Math.min(2.4, Math.min((canvas.width - 120) / width, (canvas.height - 120) / height)));
      animateCam(-((minX + maxX) / 2) * scale, -((minY + maxY) / 2) * scale);
    }
  }
  draw();
}
"""

html = html.replace('</style>', custom_css + '</style>')
html = html.replace('<div id="search-wrap">', ui_html + '\n  <div id="search-wrap">')
html = html.replace('</script>\n</body>', dijkstra_js + '\n</script>\n</body>')

# ---------------------------------------------------------
# [추가 기능 1] 사이드바 좌우 크기 조절(리사이저)
# ---------------------------------------------------------
resizer_css = """
#sidebar { z-index: 100; }
#resizer {
  position: fixed; top: 0; left: 400px; bottom: 0; width: 6px;
  cursor: ew-resize; z-index: 150; background: transparent;
  transition: background-color 0.2s;
}
#resizer:hover, #resizer.resizing { background-color: #4facfe; }
"""
html = html.replace('</style>', resizer_css + '</style>')
html = html.replace('<div id="zoom-badge">', '<div id="resizer"></div>\n<div id="zoom-badge">')

resizer_js = """
const sidebarEl = document.getElementById('sidebar');
const resizerEl = document.getElementById('resizer');
const canvasContainerEl = document.getElementById('canvas-container');
let isResizingPanel = false;

resizerEl.addEventListener('mousedown', (e) => {
  isResizingPanel = true;
  resizerEl.classList.add('resizing');
  document.body.style.cursor = 'ew-resize';
  e.preventDefault();
});

document.addEventListener('mousemove', (e) => {
  if (!isResizingPanel) return;
  let newWidth = e.clientX;
  if (newWidth < 250) newWidth = 250; 
  if (newWidth > window.innerWidth * 0.6) newWidth = window.innerWidth * 0.6; 
  sidebarEl.style.width = newWidth + 'px';
  resizerEl.style.left = newWidth + 'px';
  canvasContainerEl.style.left = newWidth + 'px';
  resize();
});

document.addEventListener('mouseup', () => {
  if (isResizingPanel) {
    isResizingPanel = false;
    resizerEl.classList.remove('resizing');
    document.body.style.cursor = 'default';
  }
});
"""
html = html.replace('</script>\n</body>', resizer_js + '\n</script>\n</body>')

# ---------------------------------------------------------
# [추가 기능 2] 일정 캡처 및 이미지 내보내기 (교체할 코드)
# ---------------------------------------------------------
# 1. 이미지 캡처용 html2canvas 라이브러리 CDN 추가
html = html.replace('</head>', '<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>\n</head>')

# 2. 버튼 HTML 추가
export_btn_html = """
<button id="export-itinerary-btn" class="find-btn" style="margin-top: 15px; display: none; background: linear-gradient(135deg, #11998e, #38ef7d);">📸 일정 이미지로 저장</button>
"""
html = html.replace('<div id="route-result"></div>', '<div id="route-result"></div>\n' + export_btn_html)

# 3. 이미지 저장 로직 및 버튼 이벤트 JS 주입
export_js = """
document.getElementById('find-route-btn').addEventListener('click', () => {
  const exportBtn = document.getElementById('export-itinerary-btn');
  const routeResultDiv = document.getElementById('route-result');
  setTimeout(() => {
    if (routeResultDiv.innerHTML.trim() !== '' && !routeResultDiv.innerHTML.includes('route-error')) {
      exportBtn.style.display = 'block';
    } else {
      exportBtn.style.display = 'none';
    }
  }, 150);
});

document.getElementById('reset-route-btn').addEventListener('click', () => {
  document.getElementById('export-itinerary-btn').style.display = 'none';
});

document.getElementById('export-itinerary-btn').addEventListener('click', () => {
  if (typeof html2canvas === 'undefined') {
    alert('이미지 저장 도구를 불러오는 중입니다. 잠시 후 다시 시도해주세요.');
    return;
  }
  
  const routeResultDiv = document.getElementById('route-result');
  const exportBtn = document.getElementById('export-itinerary-btn');
  exportBtn.style.display = 'none'; // 캡처 시 버튼 숨김
  
  const originalOverflow = routeResultDiv.style.overflow;
  const originalMaxHeight = routeResultDiv.style.maxHeight;
  routeResultDiv.style.overflow = 'visible';
  routeResultDiv.style.maxHeight = 'none';

  html2canvas(routeResultDiv, {
    backgroundColor: '#16213e',
    scale: 2,
    logging: false,
    // 화면에 안 보이는 가상 복제본에만 제목을 주입하는 핵심 로직
    onclone: function(clonedDoc) {
      const clonedResultDiv = clonedDoc.getElementById('route-result');
      
      // 현재 입력된 출발/경유/도착역 값들을 가져와서 '➔' 로 연결
      const inputs = Array.from(document.querySelectorAll('.waypoint-input'))
                          .map(el => el.value.trim())
                          .filter(v => v);
      const routeSummary = inputs.join(' ➔ ');
      
      // 캡처용 큰 글씨 제목 요소 생성 (개별 열차 13px보다 큰 18px 적용)
      const headerEl = clonedDoc.createElement('div');
      headerEl.innerHTML = `<div style="color:#ffffff; font-size:18px; font-weight:800; text-align:center; padding-bottom:12px; margin-bottom:15px; border-bottom:1px dashed #4facfe;">${routeSummary}</div>`;
      
      // 복제된 결과창 맨 위에 제목 삽입
      clonedResultDiv.insertBefore(headerEl, clonedResultDiv.firstChild);
    }
  }).then(canvas => {
    routeResultDiv.style.overflow = originalOverflow;
    routeResultDiv.style.maxHeight = originalMaxHeight;
    exportBtn.style.display = 'block';

    const link = document.createElement('a');
    link.download = '나의_기차_일정.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
  }).catch(err => {
    console.error('캡처 에러:', err);
    alert('이미지 생성 중 오류가 발생했습니다.');
    routeResultDiv.style.overflow = originalOverflow;
    routeResultDiv.style.maxHeight = originalMaxHeight;
    exportBtn.style.display = 'block';
  });
});
"""
html = html.replace('</script>\n</body>', export_js + '\n</script>\n</body>')

out = os.path.join(BASE_DIR, 'route_map.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"\n✅ 노선도 생성 완료: {out}")
print(f"   총 {len(nodes_data)}개 역, {len(edges_data)}개 구간 (무정차 포함)")
