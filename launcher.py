"""
config.py를 매번 열어서 고치지 않아도 되도록, 지역/검색 메뉴/희망 리뷰수를
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


def _run_scraper_in_thread(district, groups_text, min_reviews, log_queue):
    """실제 크롤링을 별도 스레드에서 실행한다 (창이 멈추지 않도록).

    config.py의 값들을 입력받은 값으로 덮어쓴 뒤, 기존 크롤링 코드(main())를
    그대로 호출한다.
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
        scraper.main()
    except Exception as e:
        log_queue.put(f"\n오류가 발생했습니다: {e}\n")
    finally:
        sys.stdout = old_stdout
        log_queue.put("__DONE__")


BLUE = "#4472C4"
TITLE_TEXT = "🦅 독수리오형제 Project 🦅"


class App:
    def __init__(self, root):
        self.root = root
        root.title("네이버 맛집 리뷰 검색기")
        root.configure(bg="white")

        outer = tk.Frame(root, bg="white", highlightbackground=BLUE, highlightthickness=2, padx=25, pady=20)
        outer.pack(padx=15, pady=15)

        title = tk.Label(outer, text=TITLE_TEXT, font=("맑은 고딕", 16, "bold"), bg="white", fg="black")
        title.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        self.district_entry = self._add_field(outer, 1, "희망지역", config.DISTRICT)
        self.groups_entry = self._add_field(outer, 2, "검색메뉴", ", ".join(config.PRODUCT_GROUPS))
        self.review_entry = self._add_field(outer, 3, "기준 리뷰수", str(config.MIN_REVIEW_COUNT))

        self.run_button = tk.Button(
            outer, text="실행", command=self.start,
            bg=BLUE, fg="white", font=("맑은 고딕", 11, "bold"),
            width=10, relief="flat", activebackground="#375a99", activeforeground="white",
        )
        self.run_button.grid(row=4, column=1, sticky="e", pady=(15, 0))

        self.log_box = scrolledtext.ScrolledText(root, width=100, height=24, state="disabled")
        self.log_box.pack(padx=15, pady=(0, 15), fill="both", expand=True)

        self.log_queue = queue.Queue()
        self.root.after(200, self._poll_queue)

    def _add_field(self, parent, row, label_text, default_value):
        label = tk.Label(
            parent, text=label_text, bg=BLUE, fg="white",
            font=("맑은 고딕", 11, "bold"), width=12, height=2, relief="flat",
        )
        label.grid(row=row, column=0, padx=(0, 12), pady=6)

        entry = tk.Entry(
            parent, width=40, font=("맑은 고딕", 11),
            highlightbackground=BLUE, highlightcolor=BLUE, highlightthickness=1, relief="solid", bd=1,
        )
        entry.insert(0, default_value)
        entry.grid(row=row, column=1, pady=6, sticky="w")
        return entry

    def _append_log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _poll_queue(self):
        try:
            while True:
                text = self.log_queue.get_nowait()
                if text == "__DONE__":
                    self.run_button.configure(state="normal", text="실행")
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
            messagebox.showwarning("입력 필요", "지역과 검색 메뉴를 입력해주세요.")
            return

        try:
            min_reviews = int(review_text.replace(",", ""))
        except ValueError:
            messagebox.showwarning("입력 오류", "희망 리뷰 수는 숫자로 입력해주세요.")
            return

        self.run_button.configure(state="disabled", text="실행 중...")
        self._append_log(f"\n===== '{district} / {groups_text}' (리뷰 {min_reviews}개 이상) 검색 시작 =====\n")

        thread = threading.Thread(
            target=_run_scraper_in_thread,
            args=(district, groups_text, min_reviews, self.log_queue),
            daemon=True,
        )
        thread.start()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
