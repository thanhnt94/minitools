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
    v17: Sửa lỗi "Đi đến" (Go To) bằng 2 kỹ thuật:
    - 1. Dùng `shape.api.TopLeftCell.Address` (COM trực tiếp)
         để lấy địa chỉ Cell, ổn định hơn `xlwings`.
    - 2. (Theo ý tưởng) Giả lập phím `{RIGHT}{LEFT}` sau
         khi `Goto` cell để ép Excel refresh màn hình.
    """
    def __init__(self):
        super().__init__()

        # ======= Theme / Window =======
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.title("Excel TextBox Finder") # THAY ĐỔI: v17
        self.geometry("750x750") 
        self.minsize(750, 750)

        # ======= i18n (Giữ nguyên) =======
        self.lang_var = tk.StringVar(value="vi")
        self.i18n = {
            "vi": {
                "title": "Excel TextBox Finder",
                "reload": "Tải lại DS Excel",
                # ... (giữ nguyên tất cả i18n khác) ...
                "mode_partial": "Tương đối",
                "mode_exact": "Tuyệt đối",
                "btn_add": "Thêm vào Danh sách",
                "btn_find_selected": "Tìm mục đã chọn",
                "btn_find_all": "Tìm tất cả",
                "btn_find_first": "Tìm nhanh", 
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
                "help_btn": "Trợ giúp", 
                "help_title": "Hướng dẫn sử dụng",
                "help_content": ( 
                    "\n\n"
                ),
                "caching_start": "Đang xây dựng bộ đệm cho '{book}'... (Lần đầu sẽ chậm)",
                "cache_done": "Xây dựng bộ đệm xong. Tìm thấy {n} shapes.",
                "cache_searching": "Đang tìm trong bộ đệm...",
                "cache_cleared": "Đã xoá bộ đệm.",
                "find_first_searching": "Đang tìm nhanh kết quả đầu tiên...",
                "find_first_done": "Đã tìm thấy 1 kết quả (Tìm nhanh).",
                "find_first_none": "Không tìm thấy (Tìm nhanh)."
            },
            # ... (Các ngôn ngữ 'en' và 'ja' giữ nguyên) ...
            "en": {
                "title": "Excel TextBox Finder",
                "reload": "Reload Excel List",
                "mode_partial": "Contains",
                "mode_exact": "Exact match",
                "btn_add": "Add to List",
                "btn_find_selected": "Find Selected",
                "btn_find_all": "Find All",
                "btn_find_first": "Find First (Quick)", 
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
                "help_btn": "Help", 
                "help_title": "User Guide",
                "help_content": ( 
                    "Welcome to the Excel TextBox Finder!\n\n"
                ),
                "caching_start": "Building cache for '{book}'... (First time is slow)",
                "cache_done": "Cache built. Found {n} shapes.",
                "cache_searching": "Searching in cache...",
                "cache_cleared": "Cache cleared.",
                "find_first_searching": "Quick-searching for first match...",
                "find_first_done": "Found 1 result (Quick search).",
                "find_first_none": "No match found (Quick search)."
            },
            "ja": {
                "title": "Excel テキストボックス検索",
                "reload": "Excel一覧を再読込",
                "mode_partial": "含む（部分一致）",
                "mode_exact": "完全一致",
                "btn_add": "リストに追加",
                "btn_find_selected": "選択した語で検索",
                "btn_find_all": "すべて検索",
                "btn_find_first": "最初の一致 (高速)", 
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
                "help_btn": "ヘルプ",
                "help_title": "使い方",
                "help_content": (
                    "Excel テキストボックス検索へようこそ！\n\n"
                ),
                "caching_start": "'{book}' のキャッシュを構築中... (初回は遅い)",
                "cache_done": "キャッシュ完了。{n}個のシェイプ。",
                "cache_searching": "キャッシュ内を検索中...",
                "cache_cleared": "キャッシュをクリアしました。",
                "find_first_searching": "最初の一致を高速検索中...",
                "find_first_done": "1件見つかりました (高速検索)。",
                "find_first_none": "見つかりません (高速検索)。"
            }
        }
        
        # ======= Dữ liệu =======
        self.search_results = [] 
        self.help_win = None 

        # ======= Dữ liệu CACHE =======
        self.shape_cache = [] # Sẽ lưu (sheet_name, shape_name, text, shape_api_id, cell_addr)
        self.active_cache_book = None 

        # ======= Layout (Giữ nguyên) =======
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) 

        self.top_frame = ctk.CTkFrame(self, corner_radius=12)
        self.top_frame.grid(row=0, column=0, padx=12, pady=(12,8), sticky="ew")
        self.top_frame.grid_columnconfigure(1, weight=5) 
        self.top_frame.grid_columnconfigure(2, weight=1) 
        self.top_frame.grid_columnconfigure(3, weight=1) 

        self.reload_btn = ctk.CTkButton(self.top_frame, text="", command=self.refresh_workbooks, height=36)
        self.reload_btn.grid(row=0, column=0, padx=8, pady=10)

        self.workbook_combo = ctk.CTkComboBox(self.top_frame, values=["[Chưa có file nào]"], 
                                              state="readonly", height=36, 
                                              command=self.on_workbook_change) 
        self.workbook_combo.grid(row=0, column=1, padx=8, pady=10, sticky="ew")

        self.partial_radio = ctk.CTkRadioButton(self.top_frame, text="", variable=tk.StringVar(), value="partial")
        self.exact_radio   = ctk.CTkRadioButton(self.top_frame, text="", variable=self.partial_radio._variable, value="exact")
        self.partial_radio._variable.set("partial")
        self.partial_radio.grid(row=0, column=2, padx=(12,4), pady=10, sticky="w")
        self.exact_radio.grid(row=0, column=3, padx=(4,12), pady=10, sticky="w")

        self.add_frame = ctk.CTkFrame(self, corner_radius=12)
        self.add_frame.grid(row=1, column=0, padx=12, pady=(0,8), sticky="ew")
        self.add_frame.grid_columnconfigure(0, weight=1)

        self.kw_text = ctk.CTkTextbox(self.add_frame, height=120)
        self.kw_text.grid(row=0, column=0, columnspan=3, padx=12, pady=(12,8), sticky="ew")
        
        self.add_btn = ctk.CTkButton(self.add_frame, text="", command=self.add_keywords_bulk, height=38)
        self.add_btn.grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 12), sticky="ew")

        self.bottom_frame = ctk.CTkFrame(self, corner_radius=12)
        self.bottom_frame.grid(row=2, column=0, padx=12, pady=(0,8), sticky="nsew")
        self.bottom_frame.grid_columnconfigure(0, weight=1, uniform="cols")
        self.bottom_frame.grid_columnconfigure(1, weight=1, uniform="cols")
        self.bottom_frame.grid_rowconfigure(1, weight=1)

        self.lbl_kw = ctk.CTkLabel(self.bottom_frame, text="", anchor="w")
        self.lbl_kw.grid(row=0, column=0, padx=12, pady=(12,4), sticky="ew")
        self.lbl_rs = ctk.CTkLabel(self.bottom_frame, text="", anchor="w")
        self.lbl_rs.grid(row=0, column=1, padx=12, pady=(12,4), sticky="ew")

        self.left_frame = ctk.CTkFrame(self.bottom_frame, corner_radius=10)
        self.left_frame.grid(row=1, column=0, padx=(12,6), pady=(0,12), sticky="nsew")
        
        self.left_frame.grid_rowconfigure(0, weight=0) 
        self.left_frame.grid_rowconfigure(1, weight=1) 
        self.left_frame.grid_rowconfigure(2, weight=0) 
        self.left_frame.grid_columnconfigure(0, weight=1)

        self.kw_delete_btns = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.kw_delete_btns.grid(row=0, column=0, padx=10, pady=(10,0), sticky="ew")
        self.kw_delete_btns.grid_columnconfigure((0,1), weight=1)
        
        self.remove_btn = ctk.CTkButton(self.kw_delete_btns, text="", 
                                        fg_color="#9B1C1C", hover_color="#7A1515", 
                                        command=self.remove_selected_keywords, height=34) 
        self.remove_btn.grid(row=0, column=0, padx=4, pady=4, sticky="ew")
        
        self.clear_btn = ctk.CTkButton(self.kw_delete_btns, text="", 
                                       fg_color="#B45309", hover_color="#92400E", 
                                       command=self.clear_keywords, height=34)
        self.clear_btn.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self.kw_listbox = tk.Listbox(self.left_frame, bg="#1f1f1f", fg="white",
                                     borderwidth=0, highlightthickness=0,
                                     selectbackground="#1f6AA5", activestyle="none",
                                     selectmode=tk.EXTENDED)
        self.kw_listbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        self.kw_find_btns = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.kw_find_btns.grid(row=2, column=0, padx=10, pady=(0,10), sticky="ew")
        self.kw_find_btns.grid_columnconfigure((0, 1, 2), weight=1) 

        self.find_sel_btn = ctk.CTkButton(self.kw_find_btns, text="", 
                                          command=lambda: self.find_selected_keywords(focus_list=True), 
                                          height=34,
                                          fg_color="#565B5E", hover_color="#464A4D") 
        self.find_sel_btn.grid(row=0, column=0, padx=4, pady=4, sticky="ew")
        
        self.find_all_btn = ctk.CTkButton(self.kw_find_btns, text="", 
                                          command=self.find_all_keywords, 
                                          height=34,
                                          fg_color="#0A8754", hover_color="#086644") 
        self.find_all_btn.grid(row=0, column=1, padx=4, pady=4, sticky="ew")
        
        self.find_first_btn = ctk.CTkButton(self.kw_find_btns, text="", 
                                          command=self.find_first_keyword, 
                                          height=34,
                                          fg_color="#0E7490", hover_color="#155E75") 
        self.find_first_btn.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

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

        self.footer = ctk.CTkFrame(self, corner_radius=0)
        self.footer.grid(row=3, column=0, padx=0, pady=0, sticky="ew")
        self.footer.grid_columnconfigure(0, weight=1) 

        self.status_label = ctk.CTkLabel(self.footer, text="", text_color="gray")
        self.status_label.grid(row=0, column=0, padx=12, pady=8, sticky="w")

        self.lang_label = ctk.CTkLabel(self.footer, text="")
        self.lang_label.grid(row=0, column=1, padx=(12, 0), pady=8, sticky="e")
        
        self.lang_menu = ctk.CTkOptionMenu(self.footer, values=["Tiếng Việt","English","日本語"],
                                            command=self.on_change_language, height=34)
        self.lang_menu.grid(row=0, column=2, padx=(6, 12), pady=8, sticky="e")

        self.help_btn = ctk.CTkButton(self.footer, text="", command=self.show_help_popup, 
                                      height=34, width=50)
        self.help_btn.grid(row=0, column=3, padx=(0, 12), pady=8, sticky="e")

        self.copyright_label = ctk.CTkLabel(self.footer, text="© KNT15083", text_color="#8b8b8b")
        self.copyright_label.grid(row=0, column=4, padx=(0, 12), pady=8, sticky="e")

        self.apply_language()
        self.refresh_workbooks()
        self.status_label.configure(text=self.t("status_ready"))

    # =================== Help (Giữ nguyên) ===================
    def show_help_popup(self):
        # ... (code giữ nguyên) ...
        if self.help_win is not None and self.help_win.winfo_exists():
            self.help_win.focus()
            return
        self.help_win = ctk.CTkToplevel(self)
        self.help_win.title(self.t("help_title"))
        self.help_win.geometry("600x500")
        self.help_win.transient(self) 
        self.help_win.grab_set()      
        self.help_win.resizable(False, True)
        textbox = ctk.CTkTextbox(self.help_win, corner_radius=0, wrap="word", 
                                 font=("Arial", 13) if sys.platform == "win32" else ("Arial", 14))
        textbox.pack(expand=True, fill="both", padx=10, pady=10)
        help_text = self.t("help_content")
        textbox.insert("1.0", help_text)
        textbox.configure(state="disabled") 
        self.help_win.after(50, self.help_win.focus_force)

    # =================== i18n (Giữ nguyên) ===================
    def t(self, key):
        lang = self.lang_var.get()
        return self.i18n.get(lang, self.i18n["vi"]).get(key, key)

    def on_change_language(self, _value):
        mapping = {"Tiếng Việt": "vi", "English": "en", "日本語": "ja"}
        self.lang_var.set(mapping.get(_value, "vi"))
        self.apply_language()

    def apply_language(self):
        self.title(self.t("title") + " 2.0") # THAY ĐỔI: v17
        self.reload_btn.configure(text=self.t("reload"))
        # ... (giữ nguyên) ...
        self.partial_radio.configure(text=self.t("mode_partial"))
        self.exact_radio.configure(text=self.t("mode_exact"))
        self.add_btn.configure(text=self.t("btn_add"))
        self.find_sel_btn.configure(text=self.t("btn_find_selected"))
        self.find_all_btn.configure(text=self.t("btn_find_all"))
        self.find_first_btn.configure(text=self.t("btn_find_first"))
        self.lbl_kw.configure(text=self.t("keywords"))
        self.lbl_rs.configure(text=self.t("results"))
        self.remove_btn.configure(text=self.t("btn_remove_sel"))
        self.clear_btn.configure(text=self.t("btn_clear"))
        self.goto_btn.configure(text=self.t("btn_goto"))
        self.lang_label.configure(text=self.t("lang") + ":")
        self.help_btn.configure(text=self.t("help_btn"))

    # ================= Excel helpers (Giữ nguyên) =================
    
    def on_workbook_change(self, *args):
        self.clear_shape_cache()

    def refresh_workbooks(self):
        self.clear_shape_cache() 
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
        # ... (code giữ nguyên) ...
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

    # =================== THAY ĐỔI v17 ===================
    def _scroll_shape_into_view(self, app, sheet, shape):
        """
        Di chuyển màn hình Excel đến vị trí shape
        và cố gắng căn giữa.
        """
        try:
            sheet.api.Activate()
            
            # 1. THAY ĐỔI v17: Dùng API trực tiếp
            tl_api = shape.api.TopLeftCell 
            
            # 2. Goto (this selects cell and scrolls)
            app.api.Goto(tl_api, True) # Đây là "click"
            self.update_idletasks() 
            
            # 3. Căn giữa (v16)
            try:
                win = app.api.ActiveWindow
                visible_rows = win.VisibleRange.Rows.Count
                target_row = tl_api.Row # Dùng COM property
                
                scroll_to_row = max(1, target_row - (visible_rows // 2))
                
                if hasattr(win, "ScrollRow") and win.ScrollRow != scroll_to_row:
                    win.ScrollRow = scroll_to_row
            except Exception:
                pass 
            
            return True
        except Exception:
            return False 

    def _match_text(self, text, kw, mode):
        if text is None:
            return False
        return (text == kw) if mode == "exact" else (kw in text)

    # ================= Từ khoá (Giữ nguyên) =================
    def add_keywords_bulk(self):
        # ... (code giữ nguyên) ...
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
        # ... (code giữ nguyên) ...
        sel = self.kw_listbox.curselection()
        if not sel:
            self.status_label.configure(text=self.t("select_kw"), text_color="orange")
            return
        for i in reversed(sel):
            self.kw_listbox.delete(i)
        self.status_label.configure(text=self.t("deleted_sel"), text_color="gray")

    def clear_keywords(self):
        # ... (code giữ nguyên) ...
        self.kw_listbox.delete(0, tk.END)
        self.status_label.configure(text=self.t("deleted_all"), text_color="gray")

    # ================= Cache (THAY ĐỔI v17) =================
    def clear_shape_cache(self):
        if self.active_cache_book:
            self.shape_cache.clear()
            self.active_cache_book = None
            self.status_label.configure(text=self.t("cache_cleared"), text_color="gray")

    def _build_cache_for_book(self, book):
        self.clear_shape_cache()
        self.active_cache_book = book.name
        self.status_label.configure(text=self.t("caching_start").format(book=book.name), text_color="gray")
        self.update_idletasks() 
        
        temp_cache = []
        for sheet in book.sheets:
            for shape in sheet.shapes:
                try:
                    text = getattr(shape, "text", None)
                    if text is not None:
                        shape_id = shape.api.ID # (v15)
                        
                        # THAY ĐỔI v17: Dùng COM API trực tiếp
                        try:
                            cell_addr = shape.api.TopLeftCell.Address
                        except Exception:
                            cell_addr = "$?" # Lỗi (vd: ngoài vùng)
                        
                        temp_cache.append((sheet.name, shape.name, text, shape_id, cell_addr))
                except Exception:
                    continue 
        
        self.shape_cache = temp_cache
        self.status_label.configure(text=self.t("cache_done").format(n=len(self.shape_cache)), text_color="gray")
        return len(self.shape_cache)

    # ================= Tìm kiếm (THAY ĐỔI v17) =================
    
    def _iter_shapes(self, book):
        for sheet in book.sheets:
            for shape in sheet.shapes:
                yield sheet, shape

    def _append_result(self, book_name, sheet_name, shape_name, text, kw_tag, shape_api_id, cell_addr):
        # (v16)
        disp = (text[:70] + '...') if text and len(text) > 70 else (text or "")
        
        # THAY ĐỔI v17: Xóa $ tuyệt đối cho dễ đọc
        cell_addr_clean = cell_addr.replace("$", "")
        
        line = f"[{cell_addr_clean}] {sheet_name} | {shape_name}: '{disp}'"
        line = f"[kw: {kw_tag}] {line}"
        
        self.results_listbox.insert(tk.END, line)
        self.search_results.append((book_name, sheet_name, shape_api_id, kw_tag))

    def _find_core_from_cache(self, keywords, book_name):
        self.results_listbox.delete(0, tk.END)
        self.search_results.clear()
        
        self.status_label.configure(text=self.t("cache_searching"), text_color="gray")
        self.update_idletasks()

        mode = self.partial_radio._variable.get()
        found = 0

        # (v16) Đọc 5-tuple
        for (sheet_name, shape_name, text, shape_api_id, cell_addr) in self.shape_cache:
            for kw in keywords:
                if self._match_text(text, kw, mode):
                    self._append_result(book_name, sheet_name, shape_name, text, kw, shape_api_id, cell_addr)
                    found += 1
                    break 
        
        self.status_label.configure(text=self.t("done_total").format(n=found), text_color="gray")
        return found

    def _get_book_and_build_cache(self):
        # ... (code giữ nguyên) ...
        book_name = self.workbook_combo.get()
        if not book_name or book_name.startswith("["):
            self.status_label.configure(text=self.t("choose_excel"), text_color="red")
            return None
        try:
            book = self._ensure_book(book_name)
            if not book:
                raise RuntimeError(self.t("choose_excel"))
            if self.active_cache_book != book.name:
                self._build_cache_for_book(book)
            return book
        except Exception as e:
            self.status_label.configure(text=self.t("error").format(e=e), text_color="red")
            return None

    def find_selected_keywords(self, focus_list=True):
        # ... (code giữ nguyên) ...
        sel = self.kw_listbox.curselection()
        if not sel:
            self.status_label.configure(text=self.t("select_kw"), text_color="orange")
            if focus_list:
                self.kw_listbox.focus_set()
            return
        book = self._get_book_and_build_cache() 
        if not book:
            return
        kws = [self.kw_listbox.get(i) for i in sel]
        self._find_core_from_cache(kws, book.name)

    def find_all_keywords(self):
        # ... (code giữ nguyên) ...
        count = self.kw_listbox.size()
        if count == 0:
            self.status_label.configure(text=self.t("empty_kw"), text_color="orange")
            self.kw_listbox.focus_set()
            return
        book = self._get_book_and_build_cache() 
        if not book:
            return
        kws = [self.kw_listbox.get(i) for i in range(count)]
        self._find_core_from_cache(kws, book.name)

    def find_first_keyword(self):
        count = self.kw_listbox.size()
        if count == 0:
            self.status_label.configure(text=self.t("empty_kw"), text_color="orange")
            self.kw_listbox.focus_set()
            return
            
        book_name = self.workbook_combo.get()
        if not book_name or book_name.startswith("["):
            self.status_label.configure(text=self.t("choose_excel"), text_color="red")
            return

        try:
            book = self._ensure_book(book_name)
            if not book:
                raise RuntimeError(self.t("choose_excel"))
            
            kws = [self.kw_listbox.get(i) for i in range(count)]
            mode = self.partial_radio._variable.get()

            self.results_listbox.delete(0, tk.END)
            self.search_results.clear()
            self.status_label.configure(text=self.t("find_first_searching"), text_color="gray")
            self.update_idletasks()

            for sheet, shape in self._iter_shapes(book):
                try:
                    text = getattr(shape, "text", None)
                    if text is None:
                        continue
                    for kw in kws:
                        if self._match_text(text, kw, mode):
                            shape_id = shape.api.ID 
                            
                            # THAY ĐỔI v17
                            try:
                                cell_addr = shape.api.TopLeftCell.Address
                            except Exception:
                                cell_addr = "$?"
                            
                            self._append_result(book.name, sheet.name, shape.name, text, kw, shape_id, cell_addr)
                            self.status_label.configure(text=self.t("find_first_done"), text_color="gray")
                            return 
                except Exception:
                    continue
            
            self.status_label.configure(text=self.t("find_first_none"), text_color="gray")

        except Exception as e:
            self.status_label.configure(text=self.t("error").format(e=e), text_color="red")
            return 0

    # =================== THAY ĐỔI v17 ===================
    def go_to_selection(self):
        try:
            sel = self.results_listbox.curselection()
            if not sel:
                return

            idx = sel[0]
            book_name, sheet_name, shape_api_id, kw_tag = self.search_results[idx]

            book = self._ensure_book(book_name)
            if not book:
                raise RuntimeError(self.t("choose_excel"))

            app = book.app
            sheet = book.sheets[sheet_name]

            shape = None
            for s in sheet.shapes:
                try:
                    if s.api.ID == shape_api_id:
                        shape = s
                        break
                except Exception:
                    continue
            
            if shape is None:
                raise RuntimeError(f"Shape ID {shape_api_id} not found. Please Reload list.")

            # 1. Kích hoạt sheet
            sheet.api.Activate()
            self.update_idletasks() 
            
            # 2. Cố gắng cuộn & căn giữa (dùng hàm v17)
            self._scroll_shape_into_view(app, sheet, shape)
            self.update_idletasks() 
            
            # 3. THAY ĐỔI v17: Giả lập phím (Mẹo)
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys("{RIGHT}") # Gửi phím Phải
                time.sleep(0.05)
                shell.SendKeys("{LEFT}") # Gửi phím Trái
                time.sleep(0.05)
            except Exception:
                pass # Bỏ qua nếu WScript lỗi

            self.update_idletasks() 
            
            # 4. Chọn shape (bây giờ nó đã hiển thị)
            shape.api.Select(True) 
            time.sleep(0.05)

            # 5. Đưa Excel ra phía trước
            if not self._bring_excel_to_front(app):
                try:
                    app.activate(steal_focus=True)
                except Exception:
                    pass

            self.status_label.configure(text=f"{self.t('selected_info')}: {shape.name} [{kw_tag}] — {sheet_name}", text_color="gray")

        except Exception as e:
            self.status_label.configure(text=self.t("error").format(e=e), text_color="red")
            self.refresh_workbooks()

if __name__ == "__main__":
    app = ExcelTextBoxFinder()
    app.mainloop()
