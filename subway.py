import pandas as pd
import tkinter as tk
import copy

# 엑셀 파일을 시트별로 불러오기
excel_file = r'C:\\subway_tkinter\\subway_tkinter\\subway.xlsx'
sheets = pd.read_excel(excel_file, sheet_name=None)

# 노선과 환승역 시트 분리
lines_df = {name: df for name, df in sheets.items() if name not in ['환승역','호선정보']}
transfer_df = sheets.get('환승역', pd.DataFrame())

# 호선 정보 불러오기
line_info_df = pd.read_excel(excel_file, sheet_name='호선정보')  # '호선정보' 시트에서 데이터 읽기
line_info = line_info_df.set_index('지하철명').to_dict()['노선']



# 그래프 초기화
landscape = {}
colors = {
    '1호선': '#F66130',
    '2호선': '#27AF1D',
    '3호선': '#B58941',
    '4호선': '#286CD3',
    '동해선': '#4DCAF8',
    '부김선': '#AF49CC',
    '환승역': '#FFFFFF'
}
station_colors = {}
line_mapping = {}  # 각 역의 노선 저장

# 노선 데이터로 그래프 엣지 추가
for sheet_name, data in lines_df.items():
    color = colors.get(sheet_name, '#000000')
    for i in range(len(data) - 1):
        station1 = data.iloc[i]['지하철명']
        station2 = data.iloc[i + 1]['지하철명']
        x1, y1 = data.iloc[i]['X'], data.iloc[i]['Y']
        x2, y2 = data.iloc[i + 1]['X'], data.iloc[i + 1]['Y']

        if station1 not in landscape:
            landscape[station1] = {}
        if station2 not in landscape:
            landscape[station2] = {}

        landscape[station1][station2] = 1
        landscape[station2][station1] = 1
        
        # 노선 매핑 추가
        if station1 not in line_mapping:
            line_mapping[station1] = set()
        if station2 not in line_mapping:
            line_mapping[station2] = set()
        line_mapping[station1].add(sheet_name)
        line_mapping[station2].add(sheet_name)
        
        # 역에 대한 노선별 색상 저장 (여러 노선에 속한 역일 경우 첫 번째 노선 색상 사용)
        if station1 not in station_colors:
            station_colors[station1] = color
        if station2 not in station_colors:
            station_colors[station2] = color

# 환승역 추가
for i in range(len(transfer_df)):
    station1 = transfer_df.iloc[i]['노선1']
    station2 = transfer_df.iloc[i]['노선2']
    if station1 not in landscape:
        landscape[station1] = {}
    if station2 not in landscape:
        landscape[station2] = {}
    landscape[station1][station2] = 2
    landscape[station2][station1] = 2

# tkinter 창 생성
root = tk.Tk()
root.title("지하철 노선도")

# 출발역과 도착역을 설정하여 경로 찾기
def set_stations():
    start = start_entry.get()
    end = end_entry.get()
    if start and end:
        draw_shortest_path(start, end)

# 리셋 버튼의 콜백 함수
def reset_selection():
    global clicked_stations
    clicked_stations = []
    
    # 출발역과 도착역 입력 필드 비우기
    start_entry.delete(0, tk.END)
    end_entry.delete(0, tk.END)
    
    # 총 여행 시간 및 거리 초기화
    time_label.config(text="총 여행 시간: 0 분")
    details_text.delete(1.0, tk.END)  # 기존 텍스트 삭제
    # 초기 맵 다시 그리기
    draw_map()
    
    
    

# 우측 영역에 경로 세부정보를 표시할 프레임 추가
info_frame = tk.Frame(root, padx=10, pady=10, bg='lightgrey', width=300)
info_frame.pack(side=tk.RIGHT, fill=tk.Y)

details_label = tk.Label(info_frame, text="경로 세부정보", font=("Arial", 12, "bold"), bg='lightgrey')
details_label.pack(pady=(0, 10))

details_text = tk.Text(info_frame, width=40, height=20, wrap=tk.WORD, padx=5, pady=5)
details_text.pack(expand=True)

# 입출력 및 버튼 배치
controls_frame = tk.Frame(root, padx=10, pady=10)
controls_frame.pack(side=tk.TOP, fill=tk.X)

start_label = tk.Label(controls_frame, text="출발역:")
start_label.pack(side=tk.LEFT, padx=5)
start_entry = tk.Entry(controls_frame)
start_entry.pack(side=tk.LEFT, padx=5)

end_label = tk.Label(controls_frame, text="도착역:")
end_label.pack(side=tk.LEFT, padx=5)
end_entry = tk.Entry(controls_frame)
end_entry.pack(side=tk.LEFT, padx=5)

set_button = tk.Button(controls_frame, text="경로 찾기", command=set_stations)
set_button.pack(side=tk.LEFT, padx=5)

reset_button = tk.Button(controls_frame, text="리셋", command=reset_selection)
reset_button.pack(side=tk.LEFT, padx=5)

time_label = tk.Label(controls_frame, text="총 여행 시간: 0 분")
time_label.pack(side=tk.LEFT, padx=5)

canvas_width = 1680
canvas_height = 900
canvas = tk.Canvas(root, width=canvas_width, height=canvas_height, bg='white')
canvas.pack()

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

station_positions = {}

# 노선과 역을 그리는 함수 (초기 화면 및 리셋 시 사용)
def draw_map(hidden_lines=None, highlighted_stations=None):
    canvas.delete("all")  # 기존 모든 요소 삭제
    if hidden_lines is None:
        hidden_lines = set()

    if highlighted_stations is None:
        highlighted_stations = set()

    # 환승역을 세트로 만듭니다.
    transfer_stations = set(transfer_df['지하철명'].unique())

    # 역을 표시한 세트를 만듭니다.
    displayed_stations = set()

    # 노선 그리기
    for sheet_name, data in lines_df.items():
        if sheet_name not in hidden_lines:
            color = colors[sheet_name]  # 각 노선의 색상
            for i in range(len(data) - 1):
                station1 = data.iloc[i]['지하철명']
                station2 = data.iloc[i + 1]['지하철명']
                x1, y1 = data.iloc[i]['X'], data.iloc[i]['Y']
                x2, y2 = data.iloc[i + 1]['X'], data.iloc[i + 1]['Y']
                canvas.create_line(x1, y1, x2, y2, fill=color, width=2, tags="line")

    # 역 그리기
    for sheet_name, data in lines_df.items():
        for index, row in data.iterrows():
            x, y = row['X'], row['Y']
            name = row['지하철명']
            station_positions[name] = (x, y)

            # 역 아이콘 그리기
            if name == '부전':
                if name not in displayed_stations:
                    canvas.create_oval(x-30, y-8, x+10, y+8, fill='white', outline='black', tags="station")
                    displayed_stations.add(name)
            elif name == '벡스코(시립미술관)':
                if name not in displayed_stations:
                    canvas.create_oval(x-8, y-20, x+8, y+10, fill='white', outline='black', tags="station")
                    displayed_stations.add(name)
            elif name in transfer_stations:
                if name not in displayed_stations:
                    canvas.create_oval(x-5, y-5, x+5, y+5, fill='white', outline='black', tags="station")
                    displayed_stations.add(name)
            else:
                # 일반 역의 경우
                station_color = station_colors.get(name, 'black') if name not in highlighted_stations else colors.get(sheet_name, 'black')
                canvas.create_oval(x-5, y-5, x+5, y+5, fill=station_color, tags="station")
                canvas.create_text(x, y-15, text=name, fill=station_color, tags="station_name")

    # 역 이름을 선과 아이콘 위에 표시
    name_positions = {}
    for station, coord in station_positions.items():
        x, y = coord
        offset = 0
        while (x, y) in name_positions.values():
            offset += 20
            y += offset
        name_positions[station] = (x, y)
        if station in transfer_stations:
            canvas.create_text(x, y-15, text=station, fill='black', tags="station_name")
        else:
            station_color = station_colors.get(station, 'black') if station not in highlighted_stations else colors.get(sheet_name, 'black')
            canvas.create_text(x, y-15, text=station, fill=station_color, tags="station_name")

# 최단 경로 찾기 (다익스트라 알고리즘)
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

# 경로 그리기
def draw_shortest_path(start, end):
    path, distance = find_shortest_path(start, end)
    if not path:
        return

    # 선택한 경로 이외의 모든 요소 삭제
    canvas.delete("line")  # 모든 노선 삭제
    canvas.delete("station")  # 모든 역 삭제
    canvas.delete("station_name")  # 모든 역 삭제
    canvas.delete("station_oval")  # 모든 역 삭제

    used_lines = set()
    # 경로 상의 노선만 사용하여 노선 리스트 작성
    for i in range(len(path) - 1):
        station1 = path[i]
        station2 = path[i + 1]
        for sheet_name, data in lines_df.items():
            for j in range(len(data) - 1):
                line_station1 = data.iloc[j]['지하철명']
                line_station2 = data.iloc[j + 1]['지하철명']
                if (line_station1 == station1 and line_station2 == station2) or (line_station1 == station2 and line_station2 == station1):
                    used_lines.add(sheet_name)
                    break

    # 경로 상의 노선만 다시 그리기
    for sheet_name, color in colors.items():
        if sheet_name in used_lines:  # 경로에 포함된 노선만 그리기
            data = lines_df[sheet_name]
            for i in range(len(data) - 1):
                station1 = data.iloc[i]['지하철명']
                station2 = data.iloc[i + 1]['지하철명']
                if (station1 in path and station2 in path) or (station1 in path and station2 in path):
                    x1, y1 = data.iloc[i]['X'], data.iloc[i]['Y']
                    x2, y2 = data.iloc[i + 1]['X'], data.iloc[i + 1]['Y']
                    canvas.create_line(x1, y1, x2, y2, fill=color, width=2, tags="path")

    # 경로 상의 역만 다시 표시
    for station in path:
        x, y = station_positions[station]
        
        if station == '부전':
            canvas.create_oval(x-10, y-8, x+30, y+8, fill='red', outline='black', tags="station")
        elif station == '벡스코(시립미술관)':
            canvas.create_oval(x-8, y-10, x+8, y+20, fill='red', outline='black', tags="station")
        else:
            canvas.create_oval(x-5, y-5, x+5, y+5, fill='red', tags="station")
        
        canvas.create_text(x, y-15, text=station, fill='red', tags="station")
        
    # 거리와 시간 계산
    total_distance = 0  # 총 거리 (Km)
    total_time = 0  # 총 시간 (분)
    prev_line = None
    segment_distances = []

    for i in range(len(path) - 1):
        station1 = path[i]
        station2 = path[i + 1]
        line1 = line_mapping[station1].intersection(line_mapping[station2]).pop()

        if prev_line is None or prev_line == line1:
            # 동일 노선 이동: 2km, 2분씩 증가
            segment_distances.append(2)
            total_distance += 2
            total_time += 2
        else:
            # 환승: 4km, 4분 추가
            segment_distances.append(6)  # 기본 2km 이동 + 4km 환승 거리
            total_distance += 6  # 기본 2km 이동 + 4km 환승 거리
            total_time += 6  # 기본 2분 이동 + 4분 환승 시간

        prev_line = line1

    # 총 이동 거리 및 시간 출력
    print(f"총 이동 거리: {total_distance} km")
    print(f"총 이동 시간: {total_time} 분")

    # UI에 시간 정보 업데이트
    time_label.config(text=f"총 여행 시간: {total_time} 분, 총 이동 거리: {total_distance} km")

    # 경로 세부정보 업데이트
    update_details(path, {
        'total_distance': total_distance,
        'total_time': total_time,
        'segment_distances': segment_distances
    })

# 경로 세부정보 업데이트 함수
def update_details(path, distances):
    details_text.delete(1.0, tk.END)  # 기존 텍스트 삭제

    # 경로 세부정보 텍스트 생성
    details = "최단 경로: {}\n".format(" -> ".join(path))
    details += "총 이동 거리: {} km\n".format(distances['total_distance'])
    details += "총 여행 시간: {} 분\n".format(distances['total_time'])

    details_text.insert(tk.END, details)




draw_map()
root.mainloop()
