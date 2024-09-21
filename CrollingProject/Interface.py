import pandas as pd
import tkinter as tk
from tkinter import ttk
import random as rd
from tkinter import messagebox
# 엑셀 파일 경로 입력

df = pd.read_excel("C:/Users/Public/Documents/RawData.xlsx")
df['평점'] = pd.to_numeric(df['평점'], errors='coerce').fillna(0.00)
df['개봉연도'] = pd.to_numeric(df['개봉연도'], errors='coerce').fillna(0).astype(int)
# def
# window.mainloop


def show_selected_movies(country, genre):
    global netflix_var, watcha_var, tiving_var, window, tree, selected_movies, rating_var, release_var
    # 선택된 국가와 장르에 해당하는 영화 데이터 필터링
    selected_movies = df[(df['국가'] == country) & (df['장르'] == genre)]
    # Tkinter 창 생성
    window = tk.Tk()
    window.geometry("1000x1000")
    window.title(f"{country} - {genre} 영화 목록")

    # 체크박스 변수 생성
    netflix_var = tk.BooleanVar()
    watcha_var = tk.BooleanVar()
    tiving_var = tk.BooleanVar()
    rating_var = tk.BooleanVar()
    release_var = tk.BooleanVar()
    # recommand_var = tk.BooleanVar(value=False)

    # 프레임 생성1 (OTT)
    checkbox_frame = tk.Frame(window)
    checkbox_frame.pack(side="top", pady=10)  # 체크박스 프레임을 위에 배치하고 간격을 추가

    tk.Checkbutton(checkbox_frame, text="Netflix", variable=netflix_var,
                   command=OTT_checked).pack(side="left", padx=20)
    tk.Checkbutton(checkbox_frame, text="Watcha", variable=watcha_var,
                   command=OTT_checked).pack(side="left", padx=20)
    tk.Checkbutton(checkbox_frame, text="Tiving", variable=tiving_var,
                   command=OTT_checked).pack(side="left", padx=20)

    # 프레임 생성2 (정렬 기준)
    checkbox_frame2 = tk.Frame(window)
    checkbox_frame2.pack(side="top", pady=10)  # 체크박스 프레임을 위에 배치하고 간격을 추가
    tk.Checkbutton(checkbox_frame2, text="평점순", variable=rating_var,
                   command=on_rating_checked).pack(side="left", padx=20)
    tk.Checkbutton(checkbox_frame2, text="최근 개봉일순", variable=release_var,
                   command=on_release_checked).pack(side="left", padx=20)

    tk.Button(checkbox_frame2, text="오늘의 영화 추천",
              fg="black", command=today_movie).pack(side="left", padx=20)

    # 표 생성
    tree = ttk.Treeview(window)

    # 열 제목 설정
    tree["columns"] = ("제목", "평점", "개봉연도", "OTT")
    tree.column("#0", width=0, stretch=tk.NO)  # 첫 번째 열(인덱스 열)은 보이지 않도록 설정

    # 나머지 열 설정
    for col in tree["columns"]:
        # 특정 열 크기 변경
        if col == "제목":
            # "제목" 열만 폭을 250으로 설정, 문자열 가운데 정렬
            tree.column(col, anchor="center", width=250)
        else:
            # 나머지 열은 100으로 설정, 문자열 가운데 정렬
            tree.column(col, anchor="center", width=100)
        tree.heading(col, text=col)

    # 정렬된 목록 보여주기

        # 데이터 추가 (제목과 평점만 사용)
    for index, row in selected_movies.iterrows():
        tree.insert("", "end", values=(
            row["제목"], row["평점"], row["개봉연도"], row["OTT"]))

    # 스크롤바 추가
    y_scrollbar = ttk.Scrollbar(window, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=y_scrollbar.set)
    y_scrollbar.pack(side="right", fill="y")

    # 표 출력
    tree.pack(fill="both", expand=True)

    # Tkinter 창 실행
    window.mainloop()


# 콤보박스에서 선택된 값 받아오기
def on_combobox_selected():
    selected_country = country_combo.get()
    selected_genre = genre_combo.get()
    window.destroy()
    show_selected_movies(selected_country, selected_genre)


def OTT_checked():
    global tree
    tree.delete(*tree.get_children())
    for index, row in selected_movies.iterrows():
        if netflix_var.get() and row["OTT"] == "Netflix":
            if rating_var.get() or release_var.get():
                rating_var.set(value=False)
                release_var.set(value=False)
            tree.insert("", "end", values=(
                row["제목"], row["평점"], row["개봉연도"], row["OTT"]))

        elif watcha_var.get() and row["OTT"] == "Watcha":
            if rating_var.get() or release_var.get():
                rating_var.set(value=False)
                release_var.set(value=False)
            tree.insert("", "end", values=(
                row["제목"], row["평점"], row["개봉연도"], row["OTT"]))

        elif tiving_var.get() and row["OTT"] == "Tving":
            if rating_var.get() or release_var.get():
                rating_var.set(value=False)
                release_var.set(value=False)
            tree.insert("", "end", values=(
                row["제목"], row["평점"], row["개봉연도"], row["OTT"]))


def on_rating_checked():
    global tree
    if rating_var.get():
        release_var.set(value=False)
        # 현재 트리뷰의 아이템을 리스트로 가져옵니다.
        items = [(tree.set(child, "평점"), child)
                 for child in tree.get_children("")]

    # 아이템을 정렬합니다.
        items.sort(reverse=True)
        for index, (val, child) in enumerate(items):
            tree.move(child, "", index)
        tree.see(tree.get_children("")[0])


def on_release_checked():
    global tree
    if release_var.get():
        rating_var.set(value=False)
        # 현재 트리뷰의 아이템을 리스트로 가져옵니다.
        items = [(tree.set(child, "개봉연도"), child)
                 for child in tree.get_children("")]

    # 아이템을 정렬합니다.
        items.sort(reverse=True)
        for index, (val, child) in enumerate(items):
            tree.move(child, "", index)
        tree.see(tree.get_children("")[0])


def today_movie():

    top_rated_movies = [child for child in tree.get_children("") if float(tree.item(child, "values")[1]) >= 7.0]
    if top_rated_movies:
        random_movie = rd.choice(top_rated_movies)
        movie_title = tree.item(random_movie, "values")[0]
        # 선택된 영화의 인덱스를 찾아서 해당 아이템을 포커스하도록 스크롤 조정
        index = tree.index(random_movie)
        if index:
            # 해당 영화의 아이템을 선택 상태로 만들어서 보여줌
            tree.selection_set(random_movie)
            tree.see(random_movie)
    else:
        messagebox.showerror("에러","추천 할 영화가 존재하지 않음")
        

# 중복 없이 국가와 장르 목록 가져오기
countries = df['국가'].unique().tolist()
genres = df['장르'].unique().tolist()


# Tkinter 윈도우 생성
global window
window = tk.Tk()
window.title("네이버 영화 크롤링 파일 열기")


# 국가 선택 콤보박스
country_label = tk.Label(window, text="국가 선택:")
country_label.grid(row=0, column=0, padx=10, pady=5)
country_combo = ttk.Combobox(window, values=countries)
country_combo.grid(row=0, column=1, padx=10, pady=5)
country_combo.current(0)
country_combo.set("선택")

# 장르 선택 콤보박스
genre_label = tk.Label(window, text="장르 선택:")
genre_label.grid(row=1, column=0, padx=10, pady=5)
genre_combo = ttk.Combobox(window, values=genres)
genre_combo.grid(row=1, column=1, padx=10, pady=5)
genre_combo.current(0)
genre_combo.set("선택")

# 선택완료 버튼
submit_button = tk.Button(window, text="선택완료", command=on_combobox_selected)
submit_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)


window.mainloop()
