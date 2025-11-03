import customtkinter as ctk
import xlwings as xw
import tkinter as tk
import win32gui
import win32con
import win32com.client
import time
import sys

# ===== DPI-Aware cho Windows =====
if sys.platform == "win32":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

class ExcelTextBoxFinder(ctk.CTk):
    """
    v11: Quay lại giao diện v9 (Tách biệt nút Xoá/Tìm)
    - Thêm nút Trợ giúp vào footer
    """
    def __init__(self):
        super().__init__()

        # ======= Theme / Window =======
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.title("Excel TextBox Finder") # THAY ĐỔI: v11
        self.geometry("600x750") # Quay lại kích thước v9
        self.minsize(600, 600)

        # ======= i18n =======
        self.lang_var = tk.StringVar(value="vi")
        self.i18n = {
            "vi": {
                "title": "Excel TextBox Finder",
                "reload": "Tải lại DS Excel",
                "mode_partial": "Tương đối",
                "mode_exact": "Tuyệt đối",
                "btn_add": "Thêm vào Danh sách",
                "btn_find_selected": "Tìm mục đã chọn",
                "btn_find_all": "Tìm tất cả",
                "keywords": "Danh sách Từ khoá",
                "results": "Kết quả tìm",
                "btn_remove_sel": "Xoá mục đã chọn",
                "btn_clear": "Xoá hết",
                "btn_goto": "Đi đến",
                "status_ready": "Sẵn sàng",
                "lang": "Ngôn ngữ",
                "workbook": "Workbook",
                "selected_info": "Đã chọn",
                "no_excel": "(Không có file nào đang mở)",
                "choose_excel": "Vui lòng chọn một file Excel.",
                "empty_kw": "Danh sách từ khoá đang trống.",
                "select_kw": "Chọn ít nhất 1 từ khoá ở cột trái để tìm.",
                "added": "Đã thêm {added}/{total} từ khoá.",
                "deleted_sel": "Đã xoá mục đã chọn.",
                "deleted_all": "Đã xoá toàn bộ danh sách từ khoá.",
                "scanning": "Đang quét các file Excel...",
                "found_books": "Tìm thấy {n} file.",
                "done_total": "Tổng {n} kết quả.",
                "error": "Lỗi: {e}",
                "help_btn": "Trợ giúp", # MỚI
                "help_title": "Hướng dẫn sử dụng", # MỚI
                "help_content": ( # MỚI
                    "Chào mừng bạn đến với Trình tìm TextBox Excel!\n\n"
                    "QUY TRÌNH SỬ DỤNG:\n\n"
                    "1. Tải và Chọn Workbook:\n"
                    "   - Bấm 'Tải lại DS Excel' để quét các file đang mở.\n"
                    "   - Chọn 1 file từ danh sách bên cạnh.\n\n"
                    "2. Thêm Từ khoá:\n"
                    "   - Gõ (hoặc dán) danh sách từ khoá vào ô 'Thêm vào Danh sách', mỗi từ 1 dòng.\n"
                    "   - Bấm nút 'Thêm vào Danh sách'. Các từ khoá sẽ xuất hiện ở cột 'Danh sách Từ khoá' bên dưới.\n\n"
                    "3. Tìm kiếm:\n"
                    "   - Chọn 'Tương đối' (mặc định) hoặc 'Tuyệt đối'.\n"
                    "   - Để tìm: Bấm 'Tìm tất cả' (tìm theo mọi từ trong danh sách) hoặc...\n"
                    "   - ...Chọn 1 (hoặc nhiều, giữ Ctrl/Shift) từ khoá rồi bấm 'Tìm theo mục đã chọn'.\n\n"
                    "4. Xem Kết quả:\n"
                    "   - Kết quả sẽ hiện ở cột 'Kết quả tìm'.\n\n"
                    "5. Đi đến Vị trí:\n"
                    "   - Bấm 1 lần vào kết quả bạn muốn.\n"
                    "   - Bấm nút 'Đi đến & Chọn' (hoặc nháy đúp chuột vào kết quả).\n"
                    "   - Ứng dụng sẽ tự động kích hoạt Excel, chuyển đến đúng Sheet, cuộn và chọn TextBox đó cho bạn.\n\n"
                    "--- GHI CHÚ ---\n"
                    "- Nút Xoá: Dùng 2 nút màu Đỏ/Cam để quản lý danh sách từ khoá.\n"
                    "- Ngôn ngữ: Đổi ngôn ngữ ở thanh dưới cùng."
                )
            },
            "en": {
                "title": "Excel TextBox Finder",
                "reload": "Reload Excel List",
                "mode_partial": "Contains",
                "mode_exact": "Exact match",
                "btn_add": "Add to List",
                "btn_find_selected": "Find Selected",
                "btn_find_all": "Find All",
                "keywords": "Keyword List",
                "results": "Search Results",
                "btn_remove_sel": "Remove Selected",
                "btn_clear": "Clear All",
                "btn_goto": "GO",
                "status_ready": "Ready.",
                "lang": "Language",
                "workbook": "Workbook",
                "selected_info": "Selected",
                "no_excel": "Ready. (No workbooks open)",
                "choose_excel": "Please select an Excel workbook.",
                "empty_kw": "Keyword list is empty.",
                "select_kw": "Select at least one keyword on the left.",
                "added": "Added {added}/{total} keywords.",
                "deleted_sel": "Removed selected items.",
                "deleted_all": "Cleared all keywords.",
                "scanning": "Scanning Excel workbooks...",
                "found_books": "Ready. Found {n} file(s).",
                "done_total": "Done. Total {n} results.",
                "error": "Error: {e}",
                "help_btn": "Help", # MỚI
                "help_title": "User Guide", # MỚI
                "help_content": ( # MỚI
                    "Welcome to the Excel TextBox Finder!\n\n"
                    "WORKFLOW:\n\n"
                    "1. Load and Select Workbook:\n"
                    "   - Click 'Reload Excel List' to scan open files.\n"
                    "   - Select a file from the dropdown list.\n\n"
                    "2. Add Keywords:\n"
                    "   - Type (or paste) your keywords into the 'Add to List' box, one per line.\n"
                    "   - Click the 'Add to List' button. Keywords will appear in the 'Keyword List' below.\n\n"
                    "3. Search:\n"
                    "   - Choose 'Contains (partial)' (default) or 'Exact match'.\n"
                    "   - To search: Click 'Find All' (searches all keywords) or...\n"
                    "   - ...Select one (or more, hold Ctrl/Shift) keywords, then click 'Find Selected'.\n\n"
                    "4. View Results:\n"
                    "   - Results will appear in the 'Search Results' column.\n\n"
                    "5. Go To Location:\n"
                    "   - Click once on the desired result.\n"
                    "   - Click the 'Go To & Select' button (or just double-click the result).\n"
                    "   - The app will automatically activate Excel, go to the correct Sheet, scroll, and select the TextBox for you.\n\n"
                    "--- NOTES ---\n"
                    "- Delete Buttons: Use the Red/Orange buttons to manage your keyword list.\n"
                    "- Language: Change the language on the bottom bar."
                )
            },
            "ja": {
                "title": "Excel テキストボックス検索",
                "reload": "Excel一覧を再読込",
                "mode_partial": "含む（部分一致）",
                "mode_exact": "完全一致",
                "btn_add": "リストに追加",
                "btn_find_selected": "選択した語で検索",
                "btn_find_all": "すべて検索",
                "keywords": "キーワードリスト",
                "results": "検索結果",
                "btn_remove_sel": "選択を削除",
                "btn_clear": "すべて削除",
                "btn_goto": "選択",
                "status_ready": "Ready",
                "lang": "言語",
                "workbook": "ブック",
                "selected_info": "選択済み",
                "no_excel": "準備完了。（開いているブックはありません）",
                "choose_excel": "Excel ブックを選択してください。",
                "empty_kw": "キーワードリストは空です。",
                "select_kw": "左のリストから1つ以上選択してください。",
                "added": "{total}行中 {added}件を追加（重複/コメントは除外）。",
                "deleted_sel": "選択項目を削除しました。",
                "deleted_all": "すべてのキーワードを削除しました。",
                "scanning": "ブックをスキャン中…",
                "found_books": "準備完了。{n}件のブック。",
                "done_total": "完了。合計 {n} 件。",
                "error": "エラー: {e}",
                "help_btn": "ヘルプ", # MỚI
                "help_title": "使い方", # MỚI
                "help_content": ( # MỚI
                    "Excel テキストボックス検索へようこそ！\n\n"
                    "使用手順:\n\n"
                    "1. ブックの読み込みと選択:\n"
                    "   - 「Excel一覧を再読込」をクリックして、開いているファイルをスキャンします。\n"
                    "   - 横のリストからファイルを1つ選択します。\n\n"
                    "2. キーワードの追加:\n"
                    "   - 「リストに追加」のボックスにキーワードを1行に1つずつ入力（または貼り付け）します。\n"
                    "   - 「リストに追加」ボタンをクリックします。キーワードが下の「キーワードリスト」に表示されます。\n\n"
                    "3. 検索:\n"
                    "   - 「含む（部分一致）」（デフォルト）または「完全一致」を選択します。\n"
                    "   - 検索方法: 「すべて検索」（リスト内の全キーワードで検索）をクリックするか...\n"
                    "   - ...キーワードを1つ（またはCtrl/Shiftを押しながら複数）選択し、「選択した語で検索」をクリックします。\n\n"
                    "4. 結果の表示:\n"
                    "   - 検索結果が「検索結果」の欄に表示されます。\n\n"
                    "5. 場所へ移動:\n"
                    "   - 該当する結果を1回クリックします。\n"
                    "   - 「移動して選択」ボタンをクリックします（または結果をダブルクリックします）。\n"
                    "   - ツールが自動的にExcelをアクティブにし、正しいシートに移動、スクロールして該当のテキストボックスを選択します。\n\n"
                    "--- 補足 ---\n"
                    "- 削除ボタン: 赤/オレンジ色のボタンでキーワードリストを管理します。\n"
                    "- 言語: フッターバーで言語を変更できます。"
                )
            }
        }
        
        # ======= Dữ liệu =======
        self.search_results = []
        self.help_win = None # Biến cho cửa sổ help

        # ======= Layout tổng =======
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # Main content

        # ======= TOP BAR (Layout v9) =======
        self.top_frame = ctk.CTkFrame(self, corner_radius=12)
        self.top_frame.grid(row=0, column=0, padx=12, pady=(12,8), sticky="ew")
        self.top_frame.grid_columnconfigure(1, weight=5) # Combo box
        self.top_frame.grid_columnconfigure(2, weight=1) # Radio 1
        self.top_frame.grid_columnconfigure(3, weight=1) # Radio 2

        self.reload_btn = ctk.CTkButton(self.top_frame, text="", command=self.refresh_workbooks, height=36)
        self.reload_btn.grid(row=0, column=0, padx=8, pady=10)

        self.workbook_combo = ctk.CTkComboBox(self.top_frame, values=["[Chưa có file nào]"], state="readonly", height=36)
        self.workbook_combo.grid(row=0, column=1, padx=8, pady=10, sticky="ew")

        self.partial_radio = ctk.CTkRadioButton(self.top_frame, text="", variable=tk.StringVar(), value="partial")
        self.exact_radio   = ctk.CTkRadioButton(self.top_frame, text="", variable=self.partial_radio._variable, value="exact")
        self.partial_radio._variable.set("partial")
        self.partial_radio.grid(row=0, column=2, padx=(12,4), pady=10, sticky="w")
        self.exact_radio.grid(row=0, column=3, padx=(4,12), pady=10, sticky="w")

        # ======= ADD AREA (Layout v9) =======
        self.add_frame = ctk.CTkFrame(self, corner_radius=12)
        self.add_frame.grid(row=1, column=0, padx=12, pady=(0,8), sticky="ew")
        self.add_frame.grid_columnconfigure(0, weight=1)

        self.kw_text = ctk.CTkTextbox(self.add_frame, height=120)
        self.kw_text.grid(row=0, column=0, columnspan=3, padx=12, pady=(12,8), sticky="ew")
        
        self.add_btn = ctk.CTkButton(self.add_frame, text="", command=self.add_keywords_bulk, height=38)
        self.add_btn.grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 12), sticky="ew")

        # ======= MAIN SPLIT (Keywords | Results) (Layout v9) =======
        self.bottom_frame = ctk.CTkFrame(self, corner_radius=12)
        self.bottom_frame.grid(row=2, column=0, padx=12, pady=(0,8), sticky="nsew")
        self.bottom_frame.grid_columnconfigure(0, weight=1, uniform="cols")
        self.bottom_frame.grid_columnconfigure(1, weight=1, uniform="cols")
        self.bottom_frame.grid_rowconfigure(1, weight=1)

        self.lbl_kw = ctk.CTkLabel(self.bottom_frame, text="", anchor="w")
        self.lbl_kw.grid(row=0, column=0, padx=12, pady=(12,4), sticky="ew")
        self.lbl_rs = ctk.CTkLabel(self.bottom_frame, text="", anchor="w")
        self.lbl_rs.grid(row=0, column=1, padx=12, pady=(12,4), sticky="ew")

        # --- Left column (keywords) ---
        self.left_frame = ctk.CTkFrame(self.bottom_frame, corner_radius=10)
        self.left_frame.grid(row=1, column=0, padx=(12,6), pady=(0,12), sticky="nsew")
        
        self.left_frame.grid_rowconfigure(0, weight=0) # Hàng nút Xoá (trên)
        self.left_frame.grid_rowconfigure(1, weight=1) # Listbox (giữa, co dãn)
        self.left_frame.grid_rowconfigure(2, weight=0) # Hàng nút Tìm (dưới)
        self.left_frame.grid_columnconfigure(0, weight=1)

        # Hàng 1: Nút Xoá
        self.kw_delete_btns = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.kw_delete_btns.grid(row=0, column=0, padx=10, pady=(10,0), sticky="ew")
        self.kw_delete_btns.grid_columnconfigure((0,1), weight=1)
        
        # Nút Xoá (vẫn giữ height=34 như v9)
        self.remove_btn = ctk.CTkButton(self.kw_delete_btns, text="", 
                                        fg_color="#9B1C1C", hover_color="#7A1515", 
                                        command=self.remove_selected_keywords, height=34) 
        self.remove_btn.grid(row=0, column=0, padx=6, pady=4, sticky="ew")
        
        self.clear_btn = ctk.CTkButton(self.kw_delete_btns, text="", 
                                       fg_color="#B45309", hover_color="#92400E", 
                                       command=self.clear_keywords, height=34)
        self.clear_btn.grid(row=0, column=1, padx=6, pady=4, sticky="ew")

        # Hàng 2: Listbox
        self.kw_listbox = tk.Listbox(self.left_frame, bg="#1f1f1f", fg="white",
                                     borderwidth=0, highlightthickness=0,
                                     selectbackground="#1f6AA5", activestyle="none",
                                     selectmode=tk.EXTENDED)
        self.kw_listbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        # Hàng 3: Nút Tìm
        self.kw_find_btns = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.kw_find_btns.grid(row=2, column=0, padx=10, pady=(0,10), sticky="ew")
        self.kw_find_btns.grid_columnconfigure((0,1), weight=1)

        self.find_sel_btn = ctk.CTkButton(self.kw_find_btns, text="", 
                                          command=lambda: self.find_selected_keywords(focus_list=True), 
                                          height=34,
                                          fg_color="#565B5E", hover_color="#464A4D") # Xám
        self.find_sel_btn.grid(row=0, column=0, padx=6, pady=4, sticky="ew")
        
        self.find_all_btn = ctk.CTkButton(self.kw_find_btns, text="", 
                                          command=self.find_all_keywords, 
                                          height=34,
                                          fg_color="#0A8754", hover_color="#086644") # Xanh lá
        self.find_all_btn.grid(row=0, column=1, padx=6, pady=4, sticky="ew")

        # --- Right column (results) ---
        self.right_frame = ctk.CTkFrame(self.bottom_frame, corner_radius=10)
        self.right_frame.grid(row=1, column=1, padx=(6,12), pady=(0,12), sticky="nsew")
        self.right_frame.grid_rowconfigure(0, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)

        self.results_listbox = tk.Listbox(self.right_frame, bg="#1f1f1f", fg="white",
                                          borderwidth=0, highlightthickness=0,
                                          selectbackground="#1f6AA5", activestyle="none")
        self.results_listbox.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.results_listbox.bind("<Double-Button-1>", lambda e: self.go_to_selection())

        self.goto_btn = ctk.CTkButton(self.right_frame, text="", command=self.go_to_selection, height=34)
        self.goto_btn.grid(row=1, column=0, padx=10, pady=(0,10), sticky="ew")

        # ======= FOOTER BAR (Layout v9 + Help) =======
        self.footer = ctk.CTkFrame(self, corner_radius=0)
        self.footer.grid(row=3, column=0, padx=0, pady=0, sticky="ew")
        self.footer.grid_columnconfigure(0, weight=1) # Status label

        self.status_label = ctk.CTkLabel(self.footer, text="", text_color="gray")
        self.status_label.grid(row=0, column=0, padx=12, pady=8, sticky="w")

        # Cột cho các nút bên phải
        self.lang_label = ctk.CTkLabel(self.footer, text="")
        self.lang_label.grid(row=0, column=1, padx=(12, 0), pady=8, sticky="e")
        
        self.lang_menu = ctk.CTkOptionMenu(self.footer, values=["Tiếng Việt","English","日本語"],
                                            command=self.on_change_language, height=34)
        self.lang_menu.grid(row=0, column=2, padx=(6, 12), pady=8, sticky="e")

        # THAY ĐỔI: Thêm nút Help
        self.help_btn = ctk.CTkButton(self.footer, text="", command=self.show_help_popup, 
                                      height=34, width=50)
        self.help_btn.grid(row=0, column=3, padx=(0, 12), pady=8, sticky="e")

        self.copyright_label = ctk.CTkLabel(self.footer, text="© KNT15083", text_color="#8b8b8b")
        self.copyright_label.grid(row=0, column=4, padx=(0, 12), pady=8, sticky="e")

        # ======= Khởi tạo ngôn ngữ & dữ liệu =======
        self.apply_language()
        self.refresh_workbooks()
        self.status_label.configure(text=self.t("status_ready"))

    # =================== Thêm Hàm Hiển thị Help ===================
    def show_help_popup(self):
        """Mở cửa sổ Toplevel để hiển thị hướng dẫn sử dụng."""
        if self.help_win is not None and self.help_win.winfo_exists():
            self.help_win.focus()
            return

        self.help_win = ctk.CTkToplevel(self)
        self.help_win.title(self.t("help_title"))
        self.help_win.geometry("600x500")
        
        self.help_win.transient(self) # Giữ ở trên
        self.help_win.grab_set()      # Chặn cửa sổ chính
        self.help_win.resizable(False, True)

        textbox = ctk.CTkTextbox(self.help_win, corner_radius=0, wrap="word", 
                                 font=("Arial", 13) if sys.platform == "win32" else ("Arial", 14))
        textbox.pack(expand=True, fill="both", padx=10, pady=10)
        
        help_text = self.t("help_content")
        textbox.insert("1.0", help_text)
        textbox.configure(state="disabled") # Chỉ đọc

        self.help_win.after(50, self.help_win.focus_force)

    # =================== i18n helpers ===================
    def t(self, key):
        lang = self.lang_var.get()
        return self.i18n.get(lang, self.i18n["vi"]).get(key, key)

    def on_change_language(self, _value):
        mapping = {"Tiếng Việt": "vi", "English": "en", "日本語": "ja"}
        self.lang_var.set(mapping.get(_value, "vi"))
        self.apply_language()

    def apply_language(self):
        self.title(self.t("title") + " 1.0") # THAY ĐỔI: v11
        self.reload_btn.configure(text=self.t("reload"))
        self.partial_radio.configure(text=self.t("mode_partial"))
        self.exact_radio.configure(text=self.t("mode_exact"))

        self.add_btn.configure(text=self.t("btn_add"))
        self.find_sel_btn.configure(text=self.t("btn_find_selected"))
        self.find_all_btn.configure(text=self.t("btn_find_all"))

        self.lbl_kw.configure(text=self.t("keywords"))
        self.lbl_rs.configure(text=self.t("results"))

        self.remove_btn.configure(text=self.t("btn_remove_sel"))
        self.clear_btn.configure(text=self.t("btn_clear"))
        self.goto_btn.configure(text=self.t("btn_goto"))

        self.lang_label.configure(text=self.t("lang") + ":")
        self.help_btn.configure(text=self.t("help_btn")) # THAY ĐỔI

    # ================= Excel helpers =================
    def refresh_workbooks(self):
        self.status_label.configure(text=self.t("scanning"), text_color="gray")
        try:
            names = [b.name for b in xw.books]
            if names:
                self.workbook_combo.configure(values=names)
                cur = self.workbook_combo.get()
                self.workbook_combo.set(cur if cur in names else names[0])
                self.status_label.configure(text=self.t("found_books").format(n=len(names)), text_color="gray")
            else:
                self.workbook_combo.configure(values=["[No Workbooks]"])
                self.workbook_combo.set("[No Workbooks]")
                self.status_label.configure(text=self.t("no_excel"), text_color="gray")
        except Exception as e:
            self.workbook_combo.configure(values=["[Error]"])
            self.workbook_combo.set("[Error]")
            self.status_label.configure(text=self.t("error").format(e=e), text_color="red")

    def _ensure_book(self, name):
        try:
            return xw.Book(name)
        except Exception:
            for b in xw.books:
                if b.name == name:
                    return b
        return None

    def _bring_excel_to_front(self, app):
        try:
            app.visible = True
            hwnd = app.api.Hwnd
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.05)
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('%') 
            except Exception:
                pass
            win32gui.SetForegroundWindow(hwnd)
            return True
        except Exception:
            return False

    def _scroll_shape_into_view(self, app, sheet, shape, pad_rows=5, pad_cols=2):
        try:
            tl = shape.api.TopLeftCell
            sheet.api.Activate()
            app.api.Goto(tl, True)
            try:
                win = app.api.ActiveWindow
                r = max(1, tl.Row - pad_rows)
                c = max(1, tl.Column - pad_cols)
                if hasattr(win, "ScrollRow") and win.ScrollRow != r:
                    win.ScrollRow = r
                if hasattr(win, "ScrollColumn") and win.ScrollColumn != c:
                    win.ScrollColumn = c
            except Exception:
                pass
            return True
        except Exception:
            return False

    def _match_text(self, text, kw, mode):
        if text is None:
            return False
        return (text == kw) if mode == "exact" else (kw in text)

    # =V================ Từ khoá (thêm/xoá) =================
    def add_keywords_bulk(self):
        raw = self.kw_text.get("1.0", "end").strip("\n")
        lines = [ln.strip() for ln in raw.splitlines()]
        lines = [ln for ln in lines if ln and not ln.startswith("#")]

        if not lines:
            self.status_label.configure(text=self.t("empty_kw"), text_color="orange")
            return

        existing = set(self.kw_listbox.get(0, tk.END))
        added = 0
        for s in lines:
            if s not in existing:
                self.kw_listbox.insert(tk.END, s)
                existing.add(s)
                added += 1

        self.status_label.configure(text=self.t("added").format(added=added, total=len(lines)), text_color="gray")
        self.kw_text.delete("1.0", "end")

    def remove_selected_keywords(self):
        sel = self.kw_listbox.curselection()
        if not sel:
            self.status_label.configure(text=self.t("select_kw"), text_color="orange")
            return
        for i in reversed(sel):
            self.kw_listbox.delete(i)
        self.status_label.configure(text=self.t("deleted_sel"), text_color="gray")

    def clear_keywords(self):
        self.kw_listbox.delete(0, tk.END)
        self.status_label.configure(text=self.t("deleted_all"), text_color="gray")

    # ================= Tìm kiếm =================
    def _iter_shapes(self, book):
        for sheet in book.sheets:
            for shape in sheet.shapes:
                yield sheet, shape

    def _append_result(self, book_name, sheet_name, shape_name, text, kw_tag):
        disp = (text[:70] + '...') if text and len(text) > 70 else (text or "")
        line = f"[kw: {kw_tag}] {sheet_name} | {shape_name}: '{disp}'"
        self.results_listbox.insert(tk.END, line)
        self.search_results.append((book_name, sheet_name, shape_name, kw_tag))

    def _find_core(self, keywords):
        self.results_listbox.delete(0, tk.END)
        self.search_results.clear()

        book_name = self.workbook_combo.get()
        if not book_name or book_name.startswith("["):
            self.status_label.configure(text=self.t("choose_excel"), text_color="red")
            return 0

        mode = self.partial_radio._variable.get()
        found = 0

        try:
            book = self._ensure_book(book_name)
            if not book:
                raise RuntimeError(self.t("choose_excel"))

            for sheet, shape in self._iter_shapes(book):
                try:
                    text = getattr(shape, "text", None)
                    if text is None:
                        continue
                    for kw in keywords:
                        if self._match_text(text, kw, mode):
                            self._append_result(book.name, sheet.name, shape.name, text, kw)
                            found += 1
                except Exception:
                    continue

            self.status_label.configure(text=self.t("done_total").format(n=found), text_color="gray")
            return found

        except Exception as e:
            self.status_label.configure(text=self.t("error").format(e=e), text_color="red")
            return 0

    def find_selected_keywords(self, focus_list=True):
        sel = self.kw_listbox.curselection()
        if not sel:
            self.status_label.configure(text=self.t("select_kw"), text_color="orange")
            if focus_list:
                self.kw_listbox.focus_set()
            return
        kws = [self.kw_listbox.get(i) for i in sel]
        self._find_core(kws)

    def find_all_keywords(self):
        count = self.kw_listbox.size()
        if count == 0:
            self.status_label.configure(text=self.t("empty_kw"), text_color="orange")
            self.kw_listbox.focus_set()
            return
        kws = [self.kw_listbox.get(i) for i in range(count)]
        self._find_core(kws)

    # ================= Chọn & focus (kết quả, phải) =================
    def go_to_selection(self):
        try:
            sel = self.results_listbox.curselection()
            if not sel:
                return

            idx = sel[0]
            book_name, sheet_name, shape_name, kw_tag = self.search_results[idx]

            book = self._ensure_book(book_name)
            if not book:
                raise RuntimeError(self.t("choose_excel"))

            app = book.app
            sheet = book.sheets[sheet_name]
            shape = sheet.shapes[shape_name]

            sheet.api.Activate()
            self.update_idletasks() 
            self._scroll_shape_into_view(app, sheet, shape, pad_rows=5, pad_cols=2)
            self.update_idletasks() 
            shape.api.Select(False)

            if not self._bring_excel_to_front(app):
                try:
                    app.activate(steal_focus=True)
                except Exception:
                    pass

            self.status_label.configure(text=f"{self.t('selected_info')}: {shape_name} [{kw_tag}] — {sheet_name}", text_color="gray")

        except Exception as e:
            self.status_label.configure(text=self.t("error").format(e=e), text_color="red")
            self.refresh_workbooks()

if __name__ == "__main__":
    app = ExcelTextBoxFinder()
    app.mainloop()
