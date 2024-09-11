# Excel 읽기, 쓰기
import pandas as pd
import tkinter as tk
from tkinter import *
import copy
# 이미지
from PIL import Image, ImageTk
# 자동완성 기능
import ctypes   
import re
from datetime import datetime

# 엑셀 파일을 시트별로 불러오기
excel_file = r'C:\\subway_tkinter\\subway_tkinter\\subway.xlsx'
sheets = pd.read_excel(excel_file, sheet_name=None)

# 노선과 환승역 시트 분리
lines_df = {name: df for name, df in sheets.items() if name not in ['환승역','호선정보']}
transfer_df = sheets.get('환승역', pd.DataFrame())
transfer_stations = transfer_df['지하철명'].tolist()

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
root.configure(bg='#000000')
# 입력기 관련 라이브러리 불러옴
imm32 = ctypes.WinDLL('imm32')

# 이미지 로드 및 크기 조정 함수
def load_image(file_path, size):
    try:
        image = Image.open(file_path)
        image = image.resize(size, Image.LANCZOS)  # 이미지 크기 조정
        return ImageTk.PhotoImage(image)
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# 이미지 로드
start_image_path = r"image/start.png"
end_image_path = r"image/end.png"
trans_image_path = r"image/transportation.png"
bu_image_path = r"image/transportation_bu.png"
bex_image_path = r"image/transportation_bex.png"
here_path = r"image/here.png"
top_banner_path = r"image/bar.png"
banner_path = r"image/banner.png"
search_path = r"image/검색.png"
reset_path = r"image/reset.png"

#출발, 도착 아이콘 위치
x_offset = 10
y_offset = 50

# 캔버스 사이즈
canvas_x = 1600
canvas_y = 900

# 이미지 사이즈
canvas_size = (canvas_x,canvas_y)
start_end = (50,50)
train_size = (20,20)
search_reset = (50,50)

# 이미지 사이즈 재가공
start_image = load_image(start_image_path, start_end)
end_image = load_image(end_image_path, start_end)
transfer_image = load_image(trans_image_path, train_size)  # 환승역 이미지 로드
transfer_image_bu = load_image(bu_image_path, (40, 20))  # 환승역 이미지 로드
transfer_image_bex = load_image(bex_image_path, (20, 40))  # 환승역 이미지 로드
here_image = load_image(here_path, train_size)
top_banner_image = load_image(top_banner_path, (1920,150))
banner_image = load_image(banner_path, canvas_size)
search_image = load_image(search_path, search_reset)
reset_image = load_image(reset_path, search_reset)

def add_image(x, y, image):
   canvas.create_image(x, y, image=image, anchor=tk.NW, tags="icon")

def update_time():
    # 현재 시간을 가져옴
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 레이블에 현재 시간 업데이트
    time_label.config(text=current_time)
    # 1초마다 갱신
    root.after(1000, update_time)
    
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
    
    canvas.delete("icon")
    
    # 출발역과 도착역 입력 필드 비우기
    start_entry.delete(0, tk.END)
    end_entry.delete(0, tk.END)
    
    # 총 여행 시간 및 거리 초기화
    time_label.config(text="총 여행 시간: 0 분")
    # 초기 맵 다시 그리기
    draw_map()

station_list = list(line_info.keys()) 

# 프레임 시작 -----------------------------------------------------------------------
# 상단 프레임 1
# controls_frame을 캔버스로 대체하여 배경 이미지 설정
controls_frame = tk.Canvas(root, width=top_banner_image.width(), height=top_banner_image.height(), highlightthickness=0)  # highlightthickness로 캔버스 테두리 제거
controls_frame.pack(side=tk.TOP, fill=tk.X, padx=0, pady=0)

# 캔버스에 이미지 배경 설정
controls_frame.create_image(0, 0, anchor=tk.NW, image=top_banner_image)

# 이미지 객체를 계속 참조하기 위해 저장
controls_frame.image = top_banner_image

# 출발지 입력창 (절대 좌표로 배치)
start_entry = tk.Entry(controls_frame, font=("Nanum Gothic", 15), bd=0, highlightthickness=0)  # 테두리와 하이라이트 제거
start_entry.place(x=520, y=68)

# 도착지 입력창 (절대 좌표로 배치)
end_entry = tk.Entry(controls_frame, font=("Nanum Gothic", 15), bd=0, highlightthickness=0)  # 테두리 제거
end_entry.place(x=1040, y=68)

# 검색 버튼 (이미지로 대체, 절대 좌표로 배치)
set_button = tk.Button(controls_frame, image=search_image, bd=0, highlightthickness=0,bg="#65418E", command=set_stations)  # 버튼 테두리 제거
set_button.place(x=1400, y=53)

# 리셋 버튼 (이미지로 대체, 절대 좌표로 배치)
reset_button = tk.Button(controls_frame, image=reset_image, bd=0, highlightthickness=0, bg="#65418E", command=reset_selection)  # 버튼 테두리 제거
reset_button.place(x=1470, y=53)

# 검색 자동완성 박스
search_frame = tk.Frame(root)
search_listbox = tk.Listbox(search_frame)
search_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
search_scrollbar = tk.Scrollbar(search_frame, orient=tk.VERTICAL, command=search_listbox.yview)
search_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
search_listbox.config(yscrollcommand=search_scrollbar.set)
search_frame.pack_forget()

# 포커스 인/아웃 시 검색 자동완성 리스트박스 보여주기
start_entry.bind("<FocusIn>", lambda event: entry_focus_in(start_entry, search_listbox))
end_entry.bind("<FocusIn>", lambda event: entry_focus_in(end_entry, search_listbox))

# 상단 프레임 2
button_frame = tk.Frame(root, bg="white")
button_frame.pack(side=tk.TOP, fill=tk.X)

# 호선별 버튼을 위한 프레임 설정
line_label = tk.Label(button_frame, text="노선정보 ▶ ", fg="blue", bg="white")
line_label.pack(side=tk.LEFT, padx=5)
line_buttons_frame = tk.Frame(button_frame, bg="white")
line_buttons_frame.pack(side=tk.LEFT, fill=tk.X, padx=5)

# 편의시설 버튼 프레임 설정
category_btn_frame = tk.Frame(button_frame, bg="white")
category_btn_frame.pack(side=tk.RIGHT, fill=tk.X, padx=5)
category_label = tk.Label(button_frame, text="편의시설 ▶ ", fg="blue", bg="white")
category_label.pack(side=tk.RIGHT, padx=5)


# 우측 캔버스 프레임
info_canvas = tk.Canvas(root, width=300, bg='#000000',bd=0,highlightthickness=0)
info_canvas.pack(side=tk.RIGHT, fill=tk.Y)
#채워넣을곳




### 하단프레임 ###
bottom_frame= tk.Frame(root, bg="black",bd=0,highlightthickness=0)
bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5)
# 현재 시간을 표시할 레이블
time_label = tk.Label(bottom_frame, font=('Helvetica', 16), fg='white', bg='black')
time_label.pack(side=tk.LEFT, fill=tk.X, padx=5)

# 캔버스 생성
canvas = tk.Canvas(root, width=canvas_x+20, height=canvas_y,bg='#000000',bd=0,highlightthickness=0)
canvas.pack(fill=tk.BOTH, expand=True)
# 프레임 끝-----------------------------------------------------------------------

# 이미지 로드 및 추가
if banner_image:
    canvas.create_image(0, 0, anchor=tk.NW, image=banner_image)
    canvas.image = banner_image  # 이미지 참조 유지
    print("Image loaded and banner to canvas.")
else:
    print("Failed to load image.")

clicked_stations = []  # 선택한 역을 저장할 리스트

def on_click(event):
    x, y = event.x, event.y
    for station, coord in station_positions.items():
        if abs(x - coord[0]) < 10 and abs(y - coord[1]) < 10:
            if len(clicked_stations) == 2:
                # 초기화 상태로 돌아가고, 이전에 추가한 이미지를 제거
                remove_images()  # 이전에 추가한 출발점과 도착점 이미지를 제거하는 함수
                clicked_stations.clear()  # 클릭한 역 목록을 초기화
                start_entry.delete(0, tk.END)
                end_entry.delete(0, tk.END)
            
            clicked_stations.append(station)
            
            if len(clicked_stations) == 1:
                start_entry.delete(0, tk.END)
                start_entry.insert(0, station)
                add_image(coord[0] - x_offset, coord[1] - y_offset, start_image)  # 조정된 위치에 출발점 이미지 추가
            elif len(clicked_stations) == 2:
                end_entry.delete(0, tk.END)
                end_entry.insert(0, station)
                add_image(coord[0] - x_offset, coord[1] - y_offset, end_image)  # 조정된 위치에 도착점 이미지 추가
                start = clicked_stations[0]
                end = clicked_stations[1]
                draw_shortest_path(start, end)
            return

def remove_images():
    # 여기에 출발점과 도착점 이미지를 제거하는 코드 추가
    clear_canvas()
    draw_map()
    pass

station_positions = {}

# 노선과 역을 그리는 함수 (초기 화면 및 리셋 시 사용)
def draw_map(hidden_lines=None, highlighted_stations=None):
    clear_canvas()
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
                canvas.create_line(x1, y1, x2, y2, fill=color, width=4, tags="line")

    # 역 그리기
    for sheet_name, data in lines_df.items():
        for index, row in data.iterrows():
            x, y = row['X'], row['Y']
            name = row['지하철명']
            station_positions[name] = (x, y)

            # 역 아이콘 그리기
            if name == '부전':
                if name not in displayed_stations:
                    canvas.create_image(x-10, y, image=transfer_image_bu, anchor=tk.CENTER, tags="station")
                    displayed_stations.add(name)
            elif name == '벡스코(시립미술관)':
                if name not in displayed_stations:
                    canvas.create_image(x, y, image=transfer_image_bex, anchor=tk.CENTER, tags="station")
                    displayed_stations.add(name)
            elif name in transfer_stations:
                if name not in displayed_stations:
                    canvas.create_image(x, y, image=transfer_image, anchor=tk.CENTER, tags="station")
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

# 엔트리 활성화시
def entry_focus_in(entry, listbox):
    entry.bind("<KeyRelease>", lambda event: key_release_handler(event, entry, listbox))
    
    entry.bind("<KeyPress-Up>", lambda event: move_listbox_selection(search_listbox, entry, -1))
    entry.bind("<KeyPress-Down>", lambda event: move_listbox_selection(search_listbox, entry, 1))
    entry.bind("<Return>", lambda event: select_from_listbox(entry, listbox))
    
    search_listbox.bind("<Motion>", lambda event: update_selection_on_mouse_move(event, listbox))
    search_listbox.bind("<ButtonRelease-1>", lambda event: handle_click(event,entry, listbox))

    entry.bind("<FocusOut>", lambda event: search_frame.place_forget())

# IME에서 조합 중인 문자열을 가져오는 함수
def get_ime_composition_string(hwnd):
    hIMC = imm32.ImmGetContext(hwnd)
    if not hIMC:
        return None
    buffer_size = imm32.ImmGetCompositionStringW(hIMC, 8, None, 0)
    if buffer_size > 0:
        buffer = ctypes.create_unicode_buffer(buffer_size // 2)
        imm32.ImmGetCompositionStringW(hIMC, 8, buffer, buffer_size)
        return buffer.value
    return None

def key_release_handler(event, entry, listbox):
    input_text = entry.get()

    if event.keysym == "BackSpace" and len(input_text) == 0:
        search_frame.place_forget()
        return
    
    if len(input_text) == 0: 
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        composition = get_ime_composition_string(hwnd)
        if composition:
            if re.fullmatch(r'[가-힣0-9]+', composition) and event.keysym not in ("Up", "Down", "Return"):
                update_autocomplete_list(entry, listbox, station_list, composition)
                adjust_listbox_size(entry, listbox)
            else:
                return
    else:
        if event.keysym not in ("Up", "Down", "Return"):
            update_autocomplete_list(entry, listbox, station_list, input_text)

def update_autocomplete_list(entry, listbox, station_list, search_text):
    input_text = search_text.lower()
    listbox.delete(0, tk.END)

    if input_text:
        starts_with = [station for station in station_list if station.lower().startswith(input_text)]
        starts_with.sort()

        contains = [station for station in station_list if input_text in station.lower() and not station.lower().startswith(input_text)]
        contains.sort()

        matching_stations = starts_with + contains

        if matching_stations:
            for station in matching_stations:
                listbox.insert(tk.END, station)

            max_height = 50
            listbox_height = min(len(matching_stations), max_height)
            listbox.config(height=listbox_height)

            adjust_listbox_size(entry, listbox)
            listbox.select_set(0)
        else:
            search_frame.place_forget()
    else:
        search_frame.place_forget()

def move_listbox_selection(listbox, entry, direction):
    current_selection = listbox.curselection()
    if current_selection:
        index = current_selection[0]
        new_index = index + direction
        
        if 0 <= new_index < listbox.size():
            listbox.select_clear(index)
            listbox.select_set(new_index)
            listbox.activate(new_index)
            listbox.see(new_index)
    
    adjust_listbox_size(entry, listbox)

# 리스트박스 크기 자동 조절 및 자동 배치
def adjust_listbox_size(entry, listbox):
    listbox_height = min(listbox.size(), 10)
    listbox.config(height=listbox_height)
    search_frame.place(x=entry.winfo_x(), y=entry.winfo_y()+ entry.winfo_height() )
    search_frame.lift()  # 리스트박스를 최상단으로 올리기

# 검색 방향키 선택시
def select_from_listbox(entry, listbox):
    if listbox.size() > 0:
        selected_station = listbox.get(tk.ACTIVE)
        entry.delete(0, tk.END)
        entry.insert(0, selected_station)
        search_frame.place_forget()

select_listbox_item_on_mouse_move = None

# 검색 마우스 클릭 이벤트
def handle_click(event, entry, listbox):
    global select_listbox_item_on_mouse_move
    if select_listbox_item_on_mouse_move is not None:
        entry.delete(0, tk.END)
        entry.insert(0, select_listbox_item_on_mouse_move)
        search_frame.place_forget()

# 마우스가 움직일 때 선택 항목을 바꿈
def update_selection_on_mouse_move(event, listbox):
    global select_listbox_item_on_mouse_move
    index = event.widget.nearest(event.y)
    listbox.select_clear(0, tk.END)
    listbox.select_set(index)
    listbox.activate(index)
    select_listbox_item_on_mouse_move=listbox.get(index)

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

def clear_canvas():
    canvas.delete("facility")  # 태그가 "facility"인 아이콘만 삭제
    canvas.delete("line")  # 모든 노선 삭제
    canvas.delete("path")
    canvas.delete("station")  # 모든 역 삭제
    canvas.delete("station_name")  # 모든 역 삭제
    canvas.delete("station_oval")  # 모든 역 삭제
    canvas.delete("icon")

# 경로 그리기
def draw_shortest_path(start, end):
    clear_canvas()
    path, distance = find_shortest_path(start, end)
    if not path:
        return

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
                    canvas.create_line(x1, y1, x2, y2, fill=color, width=4, tags="path")

    # 경로 상의 역만 다시 표시
    for station in path:
        x, y = station_positions[station]
        
        if station in transfer_stations:
            if station == '부전':
                canvas.create_image(x+10, y, image=transfer_image_bu, anchor=tk.CENTER, tags="station")
            elif station == '벡스코(시립미술관)':
                canvas.create_image(x, y+10, image=transfer_image_bex, anchor=tk.CENTER, tags="station")
            else:
                canvas.create_image(x, y, image=transfer_image, anchor=tk.CENTER, tags="station")
        else:
            canvas.create_oval(x-5, y-5, x+5, y+5, fill='red', tags="station")
        
        canvas.create_text(x, y-15, text=station, fill='red', tags="station")
    
    # 경로 상의 아이콘 다시 추가
    for station in path:
        x, y = station_positions[station]      
        if station == start:
            add_image(x-x_offset, y-y_offset, start_image)
        elif station == end:
            add_image(x-x_offset, y-y_offset, end_image)

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

      
# 체크박스 상태를 저장할 변수들
facility_vars = {
    '엘레베이터': IntVar(),
    '휠체어리프트': IntVar(),
    '환승주차장': IntVar(),
    '자전거보관소': IntVar(),
    '물품보관함': IntVar(),
    '자동사진기': IntVar(),
    '도시철도경찰대': IntVar(),
    '섬식형': IntVar(),
    '반대방향': IntVar()
}

# 시트에서 역 정보 읽기
stations = []
for sheet_name, data in sheets.items():
    if '지하철명' in data.columns and 'X' in data.columns and 'Y' in data.columns:
        for index, row in data.iterrows():
            station_info = {
                'name': row['지하철명'],
                'x': row['X'],
                'y': row['Y'],
                'facilities': {
                    '엘레베이터': row.get('엘레베이터', 0),
                    '휠체어리프트': row.get('휠체어리프트', 0),
                    '환승주차장': row.get('환승주차장', 0),
                    '자전거보관소': row.get('자전거보관소', 0),
                    '물품보관함': row.get('물품보관함', 0),
                    '자동사진기': row.get('자동사진기', 0),
                    '도시철도경찰대': row.get('도시철도경찰대', 0),
                    '섬식형': row.get('섬식형', 0),
                    '반대방향': row.get('반대방향', 0),
                }
            }
            stations.append(station_info)

line_images_paths = {
    '1호선': r"image/1호선.png",
    '2호선': r"image/2호선.png",
    '3호선': r"image/3호선.png",
    '4호선': r"image/4호선.png",
    '동해선': r"image/동해선.png",
    '부김선': r"image/부김선.png",
    '전체역': r"image/전체 역 보기.png"
}

line_images = {}
for line, path in line_images_paths.items():
    line_images[line] = load_image(path, (60, 30))

# 이미지 버튼
for line, img in line_images.items():
    # 이미지 버튼 생성
    button = tk.Label(line_buttons_frame, image=img, bg="white")
    button.pack(side=tk.LEFT, padx=0)
    
        # 이미지 클릭 시 함수 호출
    if line == '전체역':
        button.bind("<Button-1>", lambda e: draw_map())
    else:
        button.bind("<Button-1>", lambda e, l=line: show_line(l))

def show_line(line_name):
    clear_canvas()  # 기존 노선 및 역 삭제
    
    if line_name:
        data = lines_df[line_name]
        color = colors[line_name]
        
        # 선택된 노선의 라인만 그리기
        for i in range(len(data) - 1):
            station1 = data.iloc[i]['지하철명']
            station2 = data.iloc[i + 1]['지하철명']
            x1, y1 = data.iloc[i]['X'], data.iloc[i]['Y']
            x2, y2 = data.iloc[i + 1]['X'], data.iloc[i + 1]['Y']
            canvas.create_line(x1, y1, x2, y2, fill=color, width=4, smooth=True, tags="line")

        
        # 환승역을 세트로 만듭니다.
        transfer_stations = set(transfer_df['지하철명'].unique())
            # 역을 표시한 세트를 만듭니다.
        displayed_stations = set()
        
        # 선택된 노선의 역만 그리기
        for index, row in data.iterrows():
            x, y = row['X'], row['Y']
            name = row['지하철명']
            if name in transfer_stations:
                if name not in displayed_stations:
                    canvas.create_image(x, y, image=transfer_image, anchor=tk.CENTER, tags="station")
                    displayed_stations.add(name)
            else:
                station_color = station_colors.get(name, 'black')
                canvas.create_oval(x-5, y-5, x+5, y+5, fill=station_color, tags="station")
            canvas.create_text(x, y-15, text=name, fill=color, tags="station_name")
    else:
        draw_map()  # 선택된 노선이 없으면 모든 노선 그리기
        

def show_facilities(selected_facility=None):
    """
    선택된 편의시설을 화면에 표시합니다.
    """
    canvas.delete("facility")  # 태그가 "facility"인 아이콘만 삭제
    draw_map()
    for station in stations:
        x, y = station['x'], station['y']
        # 모든 시설 표시 또는 선택된 시설이 있는 경우만 표시
        if selected_facility is None:
            for facility in facility_vars.keys():
                if station['facilities'].get(facility, 0) == 1:
                    # 모든 시설 아이콘을 표시
                    canvas.create_image(x, y-15, image=here_image, anchor=tk.CENTER, tags="facility")
                    break
        else:
            if station['facilities'].get(selected_facility, 0) == 1:
                # 선택된 시설 아이콘을 표시
                    canvas.create_image(x, y-15, image=here_image, anchor=tk.CENTER, tags="facility")

def show_tooltip(event, text):
    tooltip = tk.Toplevel()  # 툴팁을 새로운 작은 창으로 만듦
    tooltip.wm_overrideredirect(True)  # 창 테두리 없이 만듦
    tooltip.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")  # 마우스 근처에 툴팁 배치
    label = tk.Label(tooltip, text=text, background="white", relief="solid", borderwidth=1)
    label.pack()

    # 툴팁을 기억해두기 위해 이벤트와 연결
    event.widget.tooltip = tooltip

def hide_tooltip(event):
    if hasattr(event.widget, 'tooltip'):
        event.widget.tooltip.destroy()  # 툴팁 숨기기
    
    
def create_facility_buttons():
    """
    편의시설 버튼을 생성하고 프레임에 추가합니다.
    """
    global facility_vars  # 전역 변수로 facility_vars를 사용
    facility_vars = {}  # facility_vars 초기화

    # 편의시설에 해당하는 이미지 딕셔너리 예시
    facility_images_paths = {
        '엘레베이터': r"image/엘리베이터.png",
        '휠체어리프트': r"image/휠체어리프트.png",
        '환승주차장': r"image/환승주차장.png",
        '자전거보관소': r"image/자전거 보관소.png",
        '물품보관함': r"image/물품보관함.png",
        '자동사진기': r"image/자동사진기.png",
        '도시철도경찰대': r"image/도시철도경찰대.png",
        '섬식형': r"image/섬식형.png",
        '반대방향': r"image/반대방향.png"
    }
    facility_images = {}
    for facility, path in facility_images_paths.items():
        facility_images[facility] = load_image(path, (30, 30))

    # 편의시설 버튼 생성
    for facility, img in facility_images.items():
        var = tk.IntVar()  # 체크박스의 상태를 관리할 변수
        facility_vars[facility] = var
        button = tk.Label(category_btn_frame, image=img, bg="white")
        button.image = img  # 유지되도록 참조를 유지합니다.
        button.pack(side=tk.LEFT, padx=5, pady=2)

        # 이미지 클릭 시 함수 호출
        button.bind("<Button-1>", lambda e, f=facility: show_facilities(f))
        
        # 툴팁 생성
        button.bind("<Enter>", lambda e, f=facility: show_tooltip(e, f))
        button.bind("<Leave>", hide_tooltip)

canvas.bind("<Button-1>", on_click)
update_time()
create_facility_buttons()  # 편의시설 버튼 생성
draw_map()
root.mainloop()