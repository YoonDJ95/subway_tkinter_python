import pandas as pd
import tkinter as tk
import copy
from collections import deque

# 엑셀 파일을 시트별로 불러오기
excel_file = r'C:\\py\\프로젝트\\subway.xlsx'
sheets = pd.read_excel(excel_file, sheet_name=None)

# 노선과 환승역 시트 분리
lines_df = {name: df for name, df in sheets.items() if name != '환승역'}
transfer_df = sheets.get('환승역', pd.DataFrame())

# 그래프 초기화
landscape = {}
colors = {
    '1호선': '#F66130',  # RGB(246, 97, 48)
    '2호선': '#27AF1D',  # RGB(39, 175, 29)
    '3호선': '#B58941',  # RGB(181, 137, 65)
    '4호선': '#286CD3',  # RGB(40, 108, 211)
    '동해선': '#4DCAF8',  # RGB(77, 202, 248)
    '부김선': '#AF49CC'   # RGB(175, 73, 204)
}
station_colors = {}  # 역별 색상 저장

# 노선 데이터로 그래프 엣지 추가
for sheet_name, data in lines_df.items():
    color = colors.get(sheet_name, '#000000')  # 기본 색상 검정
    for i in range(len(data) - 1):
        station1 = data.iloc[i]['지하철명']
        station2 = data.iloc[i + 1]['지하철명']
        x1, y1 = data.iloc[i]['X'], data.iloc[i]['Y']
        x2, y2 = data.iloc[i + 1]['X'], data.iloc[i + 1]['Y']

        if station1 not in landscape:
            landscape[station1] = {}
        if station2 not in landscape:
            landscape[station2] = {}

        landscape[station1][station2] = 2  # 정류장 간 거리는 2로 설정
        landscape[station2][station1] = 2  # 양방향으로 설정

# 환승역 추가 (다중 그래프 엣지 추가)
for i in range(len(transfer_df)):
    station1 = transfer_df.iloc[i]['노선1']
    station2 = transfer_df.iloc[i]['노선2']
    if station1 not in landscape:
        landscape[station1] = {}
    if station2 not in landscape:
        landscape[station2] = {}
    landscape[station1][station2] = 4  # 환승역은 거리 4로 설정
    landscape[station2][station1] = 4  # 양방향으로 설정

# tkinter 창 생성
root = tk.Tk()
root.title("지하철 노선도")

# 캔버스 크기 설정
canvas_width = 1680
canvas_height = 900
canvas = tk.Canvas(root, width=canvas_width, height=canvas_height, bg='white')
canvas.pack()

# 역 클릭 이벤트 처리 함수
clicked_stations = []

def on_click(event):
    x, y = event.x, event.y
    for station, coord in station_positions.items():
        if abs(x - coord[0]) < 10 and abs(y - coord[1]) < 10:
            clicked_stations.append(station)
            if len(clicked_stations) == 1:
                start_entry.delete(0, tk.END)
                start_entry.insert(0, station)
            elif len(clicked_stations) == 2:
                end_entry.delete(0, tk.END)
                end_entry.insert(0, station)
                start = clicked_stations[0]
                end = clicked_stations[1]
                draw_shortest_path(start, end)
            return

canvas.bind("<Button-1>", on_click)

# 역의 좌표와 이름을 저장
station_positions = {}

def draw_station(data, color):
    for index, row in data.iterrows():
        x, y = row['X'], row['Y']
        name = row['지하철명']
        station_positions[name] = (x, y)
        canvas.create_oval(x-5, y-5, x+5, y+5, fill=color)
        canvas.create_text(x, y-10, text=name, fill=color)

def draw_all_lines():
    for sheet_name, data in lines_df.items():
        color = colors[sheet_name]
        for i in range(len(data) - 1):
            station1 = data.iloc[i]['지하철명']
            station2 = data.iloc[i + 1]['지하철명']
            x1, y1 = data.iloc[i]['X'], data.iloc[i]['Y']
            x2, y2 = data.iloc[i + 1]['X'], data.iloc[i + 1]['Y']
            canvas.create_line(x1, y1, x2, y2, fill=color, width=2)

# 각 시트별로 역 그리기 및 모든 노선 연결
for sheet_name, data in lines_df.items():
    draw_station(data, colors[sheet_name])
draw_all_lines()

# 다익스트라 알고리즘 적용 함수
def visitPlace(visit, routing):
    routing[visit]['visited'] = 1
    for togo, betweenDist in landscape[visit].items():
        toDist = routing[visit]['shortestDist'] + betweenDist
        if routing[togo]['shortestDist'] > toDist:
            routing[togo]['shortestDist'] = toDist
            routing[togo]['route'] = copy.deepcopy(routing[visit]['route'])
            routing[togo]['route'].append(visit)

def find_shortest_path(start, end):
    routing = {}
    for place in landscape.keys():
        routing[place] = {'shortestDist': float('inf'), 'route': [], 'visited': 0}

    routing[start]['shortestDist'] = 0

    visitPlace(start, routing)

    while True:
        minDist = float('inf')
        toVisit = ''
        for name, search in routing.items():
            if 0 < search['shortestDist'] < minDist and not search['visited']:
                minDist = search['shortestDist']
                toVisit = name
        if toVisit == '':
            break
        visitPlace(toVisit, routing)

    return routing[end]['route'] + [end], routing[end]['shortestDist']

# 최단 경로를 그리는 함수
def draw_shortest_path(start, end):
    path, distance = find_shortest_path(start, end)
    if not path:
        return

    canvas.delete("path")  # 기존 경로 삭제
    canvas.delete("station")  # 기존 역 삭제

    # 노선과 역의 색상 업데이트
    for sheet_name, color in colors.items():
        for i in range(len(lines_df[sheet_name]) - 1):
            station1 = lines_df[sheet_name].iloc[i]['지하철명']
            station2 = lines_df[sheet_name].iloc[i + 1]['지하철명']
            if station1 in path or station2 in path:
                x1, y1 = lines_df[sheet_name].iloc[i]['X'], lines_df[sheet_name].iloc[i]['Y']
                x2, y2 = lines_df[sheet_name].iloc[i + 1]['X'], lines_df[sheet_name].iloc[i + 1]['Y']
                canvas.create_line(x1, y1, x2, y2, fill=color, width=2, tags="path")

    # 역 표시
    for station in path:
        x, y = station_positions[station]
        canvas.create_oval(x-5, y-5, x+5, y+5, fill='red', tags="station")
        canvas.create_text(x, y-10, text=station, fill='red', tags="station")

    # 경로와 거리 출력
    print(f"최단 경로: {start} -> {end}는 {path}이며 거리: {distance}")

    # 시간 계산
    time = len(path) * 2  # 각 정류장마다 2분
    for i in range(len(path) - 1):
        if (path[i], path[i + 1]) in landscape:
            if landscape[path[i]][path[i + 1]] == 4:
                time += 2  # 환승역에서 추가 4분

    time_label.config(text=f"총 여행 시간: {time} 분")

# 출발역과 도착역 입력창 및 버튼
def set_stations():
    start = start_entry.get()
    end = end_entry.get()
    if start and end:
        print(f"최단 경로를 계산합니다: {start} -> {end}")
        draw_shortest_path(start, end)

# 선택 초기화 함수
def reset_selection():
    global clicked_stations
    clicked_stations = []
    canvas.delete("path")
    canvas.delete("station")
    start_entry.delete(0, tk.END)
    end_entry.delete(0, tk.END)
    draw_all_lines()

start_label = tk.Label(root, text="출발역:")
start_label.pack(side=tk.LEFT, padx=5)
start_entry = tk.Entry(root)
start_entry.pack(side=tk.LEFT, padx=5)

end_label = tk.Label(root, text="도착역:")
end_label.pack(side=tk.LEFT, padx=5)
end_entry = tk.Entry(root)
end_entry.pack(side=tk.LEFT, padx=5)

set_button = tk.Button(root, text="경로 찾기", command=set_stations)
set_button.pack(side=tk.LEFT, padx=5)

reset_button = tk.Button(root, text="초기화", command=reset_selection)
reset_button.pack(side=tk.LEFT, padx=5)

time_label = tk.Label(root, text="총 여행 시간: 0 분")
time_label.pack(side=tk.LEFT, padx=5)

root.mainloop()
