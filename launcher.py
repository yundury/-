"""
config.py를 매번 열어서 고치지 않아도 되도록, 지역/키워드/희망 리뷰수를
입력하는 작은 창(프로그램)입니다. "실행"을 누르면 naver_restaurant_scraper.py의
크롤링 코드가 그대로 실행됩니다.

실행 방법: python launcher.py
"""

import importlib
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext

import config
import naver_restaurant_scraper as scraper


class _QueueWriter:
    """print() 출력을 화면(텍스트 상자)으로 보내기 위한 도우미."""

    def __init__(self, q):
        self.q = q

    def write(self, text):
        if text:
            self.q.put(text)
        return len(text)

    def flush(self):
        pass


def _run_scraper_in_thread(district, groups_text, brand, min_reviews, log_queue, stop_event):
    """실제 크롤링을 별도 스레드에서 실행한다 (창이 멈추지 않도록).

    config.py의 값들을 입력받은 값으로 덮어쓴 뒤, 기존 크롤링 코드(main())를
    그대로 호출한다. stop_event는 "중지" 버튼을 눌렀을 때 크롤링 쪽에 신호를
    보내서, 지금까지 모은 것만이라도 저장하고 멈추게 한다.

    희망지역/키워드/브랜드/기준리뷰수는 전부 비워둬도 되는 선택 입력이다.

    맨 처음에 config.py를 다시 읽어온다 - GUI 창을 계속 켜둔 채로 config.py의
    NOTION_ENABLED 같은 값을 고친 경우에도, 새로 실행할 때는 방금 저장한 값을
    반영해야 하기 때문이다 (안 그러면 프로그램이 켜져 있는 동안 맨 처음 읽은
    값을 계속 쓰게 된다).
    """
    importlib.reload(config)
    config.DISTRICT = district
    config.PRODUCT_GROUPS = [g.strip() for g in groups_text.replace(",", "\n").split("\n") if g.strip()]
    config.BRAND = brand
    config.MIN_REVIEW_COUNT = min_reviews

    name_parts = [p.replace("/", "_") for p in (district, "_".join(config.PRODUCT_GROUPS), brand) if p]
    base_name = "_".join(name_parts) if name_parts else "검색결과"
    config.OUTPUT_FILE = f"{base_name}_리뷰{min_reviews}개이상.xlsx"

    old_stdout = sys.stdout
    sys.stdout = _QueueWriter(log_queue)
    output_path = None
    try:
        output_path = scraper.main(stop_event=stop_event)
    except Exception as e:
        log_queue.put(f"\n오류가 발생했습니다: {e}\n")
    finally:
        sys.stdout = old_stdout
        log_queue.put(("__DONE__", output_path))


# "포근한 화이트"(A) 시안의 여백/모양 + "미니멀 카드"(C) 시안의 색감을 섞은 색상표
BG_COLOR = "#f7f7f5"        # 창 배경 (C의 옅은 그레이)
CARD_BG = "#ffffff"         # 카드(폼) 배경
CARD_BORDER = "#e2e2de"     # 카드 테두리
TITLE_COLOR = "#16181d"     # C의 잉크색
LABEL_COLOR = "#53565c"
NOTE_COLOR = "#8b8c86"      # 입력창 아래 안내 문구
INPUT_BORDER = "#e2e2de"    # 입력창 테두리 (평소)
ACCENT_COLOR = "#3554d1"    # C의 포인트 인디고블루
ACCENT_HOVER = "#2c46b3"
ACCENT_DISABLED = "#a9b6ea"
STOP_FILL = "#f7e8e5"       # 중지 버튼 배경 (옅은 러스트 톤)
STOP_FILL_HOVER = "#f0dbd6"
STOP_TEXT = "#a33f2b"       # 중지 버튼 글자색 (C의 러스트레드)
LOG_TEXT = "#4b4c48"
TITLE_TEXT = "🦅 독수리오형제 Project"


class App:
    def __init__(self, root):
        self.root = root
        root.title("네이버 맛집 리뷰 검색기")
        root.configure(bg=BG_COLOR)

        card = tk.Frame(
            root, bg=CARD_BG, padx=40, pady=32,
            highlightbackground=CARD_BORDER, highlightthickness=1,
        )
        card.pack(padx=30, pady=30)
        card.grid_columnconfigure(1, weight=1)

        title = tk.Label(
            card, text=TITLE_TEXT, font=("맑은 고딕", 15, "bold"),
            bg=CARD_BG, fg=TITLE_COLOR,
        )
        title.grid(row=0, column=0, columnspan=2, pady=(0, 24))

        self.district_entry = self._add_field(card, 1, "희망지역", "ex. 서울, 부산, 강남구 등")
        self.groups_entry = self._add_field(card, 3, "키워드", "ex. 한식 맛집, 새로오픈한맛집 등")
        self.brand_entry = self._add_field(card, 5, "브랜드", "ex. 스타벅스, 교촌치킨 등 (특정 브랜드만 찾을 때)")
        self.review_entry = self._add_field(card, 7, "기준리뷰수", "*기준 리뷰수 이상 리뷰가 달린 브랜드만 수집합니다. (비워두면 전부 수집)")

        button_row = tk.Frame(card, bg=CARD_BG)
        button_row.grid(row=9, column=1, sticky="e", pady=(16, 0))

        self.stop_button = tk.Button(
            button_row, text="중지", command=self.stop,
            bg=STOP_FILL, fg=STOP_TEXT, font=("맑은 고딕", 11, "bold"),
            padx=24, pady=8, relief="flat", bd=0,
            activebackground=STOP_FILL_HOVER, activeforeground=STOP_TEXT,
            cursor="hand2", state="disabled",
        )
        self.stop_button.pack(side="left", padx=(0, 8))
        self.stop_button.bind("<Enter>", lambda e: self._set_stop_hover(True))
        self.stop_button.bind("<Leave>", lambda e: self._set_stop_hover(False))

        self.run_button = tk.Button(
            button_row, text="실행", command=self.start,
            bg=ACCENT_COLOR, fg="white", font=("맑은 고딕", 11, "bold"),
            padx=24, pady=8, relief="flat", bd=0,
            activebackground=ACCENT_HOVER, activeforeground="white",
            cursor="hand2",
        )
        self.run_button.pack(side="left")
        self.run_button.bind("<Enter>", lambda e: self._set_button_hover(True))
        self.run_button.bind("<Leave>", lambda e: self._set_button_hover(False))

        self.stop_event = None

        self.log_box = scrolledtext.ScrolledText(
            root, width=100, height=22, state="disabled",
            bg=CARD_BG, fg=LOG_TEXT, relief="flat",
            highlightbackground=CARD_BORDER, highlightthickness=1,
            font=("맑은 고딕", 10),
        )
        self.log_box.pack(padx=30, pady=(0, 30), fill="both", expand=True)

        self.log_queue = queue.Queue()
        self.root.after(200, self._poll_queue)

    def _add_field(self, parent, row, label_text, note_text=None):
        label = tk.Label(
            parent, text=label_text, bg=CARD_BG, fg=LABEL_COLOR,
            font=("맑은 고딕", 10, "bold"), anchor="w", justify="left",
        )
        label.grid(row=row, column=0, sticky="w", pady=(8, 0))

        entry = tk.Entry(
            parent, font=("맑은 고딕", 11),
            highlightbackground=INPUT_BORDER, highlightcolor=ACCENT_COLOR,
            highlightthickness=1, relief="flat", bd=6,
        )
        entry.grid(row=row, column=1, pady=(8, 0), sticky="ew")

        if note_text:
            note = tk.Label(
                parent, text=note_text, bg=CARD_BG, fg=NOTE_COLOR, font=("맑은 고딕", 9),
            )
            note.grid(row=row + 1, column=1, sticky="w", pady=(0, 4))

        return entry

    def _append_log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_button_hover(self, hovering):
        if str(self.run_button["state"]) == "disabled":
            return
        self.run_button.configure(bg=ACCENT_HOVER if hovering else ACCENT_COLOR)

    def _set_stop_hover(self, hovering):
        if str(self.stop_button["state"]) == "disabled":
            return
        self.stop_button.configure(bg=STOP_FILL_HOVER if hovering else STOP_FILL)

    def _poll_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item[:1] == ("__DONE__",):
                    output_path = item[1]
                    self.run_button.configure(state="normal", text="실행", bg=ACCENT_COLOR)
                    self.stop_button.configure(state="disabled", bg=STOP_FILL)
                    self.stop_event = None
                    self._append_log("\n===== 완료! 결과 엑셀 파일을 확인하세요 =====\n")
                    self._notify_done(output_path)
                else:
                    self._append_log(item)
        except queue.Empty:
            pass
        self.root.after(200, self._poll_queue)

    def _notify_done(self, output_path):
        if not output_path or not os.path.exists(output_path):
            messagebox.showinfo("완료", "검색이 끝났습니다.\n(저장된 결과 파일을 찾지 못했어요 - 조건에 맞는 가게가 없었을 수 있어요.)")
            return

        filename = os.path.basename(output_path)
        if messagebox.askyesno("완료", f"검색이 끝났습니다!\n\n'{filename}' 파일을 지금 확인하시겠어요?"):
            self._open_file(output_path)

    def _open_file(self, path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showwarning("열기 실패", f"파일을 여는 데 실패했습니다: {e}")

    def start(self):
        district = self.district_entry.get().strip()
        groups_text = self.groups_entry.get().strip()
        brand = self.brand_entry.get().strip()
        review_text = self.review_entry.get().strip()

        if not district and not groups_text and not brand:
            messagebox.showwarning("입력 필요", "희망지역 / 키워드 / 브랜드 중 하나는 입력해주세요.")
            return

        if review_text:
            try:
                min_reviews = int(review_text.replace(",", ""))
            except ValueError:
                messagebox.showwarning("입력 오류", "기준 리뷰 수는 숫자로 입력해주세요.")
                return
        else:
            min_reviews = 0

        self.run_button.configure(state="disabled", text="실행 중...", bg=ACCENT_DISABLED)
        self.stop_button.configure(state="normal", bg=STOP_FILL)
        summary = " / ".join(p for p in (district, groups_text, brand) if p)
        self._append_log(f"\n===== '{summary}' (리뷰 {min_reviews}개 이상) 검색 시작 =====\n")

        self.stop_event = threading.Event()
        thread = threading.Thread(
            target=_run_scraper_in_thread,
            args=(district, groups_text, brand, min_reviews, self.log_queue, self.stop_event),
            daemon=True,
        )
        thread.start()

    def stop(self):
        if self.stop_event is None:
            return
        self.stop_event.set()
        self.stop_button.configure(state="disabled", bg=STOP_FILL)
        self._append_log("\n===== 중지 요청함 - 지금까지 모은 것만 저장하고 곧 끝납니다 =====\n")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
