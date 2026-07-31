"""
config.py를 매번 열어서 고치지 않아도 되도록, 지역/키워드/희망 리뷰수를
입력하는 작은 창(프로그램)입니다. "실행"을 누르면 naver_restaurant_scraper.py의
크롤링 코드가 그대로 실행됩니다.

실행 방법: python launcher.py
"""

import queue
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


def _run_scraper_in_thread(district, groups_text, min_reviews, log_queue, stop_event):
    """실제 크롤링을 별도 스레드에서 실행한다 (창이 멈추지 않도록).

    config.py의 값들을 입력받은 값으로 덮어쓴 뒤, 기존 크롤링 코드(main())를
    그대로 호출한다. stop_event는 "중지" 버튼을 눌렀을 때 크롤링 쪽에 신호를
    보내서, 지금까지 모은 것만이라도 저장하고 멈추게 한다.
    """
    config.DISTRICT = district
    config.PRODUCT_GROUPS = [g.strip() for g in groups_text.replace(",", "\n").split("\n") if g.strip()]
    config.MIN_REVIEW_COUNT = min_reviews

    safe_district = district.replace("/", "_")
    safe_groups = "_".join(config.PRODUCT_GROUPS).replace("/", "_")
    config.OUTPUT_FILE = f"{safe_district}_{safe_groups}_리뷰{min_reviews}개이상.xlsx"

    old_stdout = sys.stdout
    sys.stdout = _QueueWriter(log_queue)
    try:
        scraper.main(stop_event=stop_event)
    except Exception as e:
        log_queue.put(f"\n오류가 발생했습니다: {e}\n")
    finally:
        sys.stdout = old_stdout
        log_queue.put("__DONE__")


# 디자이너가 만든 HTML/CSS 시안(모던 플랫 스타일)에 맞춘 색상표
BG_COLOR = "#f4f6f8"       # 창 배경 (은은한 밝은 회색)
CARD_BG = "#ffffff"        # 카드(폼) 배경
CARD_BORDER = "#e9ecef"    # 카드 테두리 (그림자 대신 은은한 선으로 구분)
TITLE_COLOR = "#111111"
LABEL_COLOR = "#495057"
INPUT_BORDER = "#dee2e6"   # 입력창 테두리 (평소)
ACCENT_COLOR = "#3b82f6"   # 포인트 블루 (포커스/버튼)
ACCENT_HOVER = "#2563eb"   # 버튼 위에 마우스 올렸을 때
STOP_COLOR = "#ef4444"     # 중지 버튼 (빨강)
STOP_HOVER = "#dc2626"
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

        self.district_entry = self._add_field(card, 1, "희망지역", config.DISTRICT)
        self.groups_entry = self._add_field(card, 2, "키워드", ", ".join(config.PRODUCT_GROUPS))
        self.review_entry = self._add_field(card, 3, "기준 리뷰수", str(config.MIN_REVIEW_COUNT))

        button_row = tk.Frame(card, bg=CARD_BG)
        button_row.grid(row=4, column=1, sticky="e", pady=(20, 0))

        self.stop_button = tk.Button(
            button_row, text="중지", command=self.stop,
            bg=STOP_COLOR, fg="white", font=("맑은 고딕", 11, "bold"),
            padx=24, pady=8, relief="flat", bd=0,
            activebackground=STOP_HOVER, activeforeground="white",
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
            bg=CARD_BG, fg="#212529", relief="flat",
            highlightbackground=CARD_BORDER, highlightthickness=1,
            font=("맑은 고딕", 10),
        )
        self.log_box.pack(padx=30, pady=(0, 30), fill="both", expand=True)

        self.log_queue = queue.Queue()
        self.root.after(200, self._poll_queue)

    def _add_field(self, parent, row, label_text, default_value):
        label = tk.Label(
            parent, text=label_text, bg=CARD_BG, fg=LABEL_COLOR,
            font=("맑은 고딕", 10, "bold"), anchor="w", width=10,
        )
        label.grid(row=row, column=0, sticky="w", pady=8)

        entry = tk.Entry(
            parent, font=("맑은 고딕", 11),
            highlightbackground=INPUT_BORDER, highlightcolor=ACCENT_COLOR,
            highlightthickness=1, relief="flat", bd=6,
        )
        entry.insert(0, default_value)
        entry.grid(row=row, column=1, pady=8, sticky="ew")
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
        self.stop_button.configure(bg=STOP_HOVER if hovering else STOP_COLOR)

    def _poll_queue(self):
        try:
            while True:
                text = self.log_queue.get_nowait()
                if text == "__DONE__":
                    self.run_button.configure(state="normal", text="실행", bg=ACCENT_COLOR)
                    self.stop_button.configure(state="disabled", bg=STOP_COLOR)
                    self.stop_event = None
                    self._append_log("\n===== 완료! 결과 엑셀 파일을 확인하세요 =====\n")
                else:
                    self._append_log(text)
        except queue.Empty:
            pass
        self.root.after(200, self._poll_queue)

    def start(self):
        district = self.district_entry.get().strip()
        groups_text = self.groups_entry.get().strip()
        review_text = self.review_entry.get().strip()

        if not district or not groups_text:
            messagebox.showwarning("입력 필요", "지역과 키워드를 입력해주세요.")
            return

        try:
            min_reviews = int(review_text.replace(",", ""))
        except ValueError:
            messagebox.showwarning("입력 오류", "희망 리뷰 수는 숫자로 입력해주세요.")
            return

        self.run_button.configure(state="disabled", text="실행 중...", bg="#93c5fd")
        self.stop_button.configure(state="normal", bg=STOP_COLOR)
        self._append_log(f"\n===== '{district} / {groups_text}' (리뷰 {min_reviews}개 이상) 검색 시작 =====\n")

        self.stop_event = threading.Event()
        thread = threading.Thread(
            target=_run_scraper_in_thread,
            args=(district, groups_text, min_reviews, self.log_queue, self.stop_event),
            daemon=True,
        )
        thread.start()

    def stop(self):
        if self.stop_event is None:
            return
        self.stop_event.set()
        self.stop_button.configure(state="disabled", bg=STOP_COLOR)
        self._append_log("\n===== 중지 요청함 - 지금까지 모은 것만 저장하고 곧 끝납니다 =====\n")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
