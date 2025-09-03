# -*- coding: utf-8 -*-
"""
Resistance & Voltage Checker (Blue/White UI)
- Right panel fixed width with fully working scrolling (mouse wheel + buttons)
- Limits section uses pack (no grid)
- Active column is centered color block (Label)
- R/V values centered; red text when out of spec
- Manual / Auto measurement (simulated)
- Auto-export toggle (save once after all points in Auto mode) + Manual export
- Clear Results (reset only measurements) vs Reset All
- COM ports are numerically sorted (COM2, COM3, COM10 ...)
"""

import os, sys, time, random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------------- Theme ----------------
COLOR_BG        = "#EAF6FF"
COLOR_PANEL     = "#FFFFFF"
COLOR_BORDER    = "#B6D9F2"
COLOR_TEXT      = "#1E2A35"
COLOR_MUTED     = "#6C7A89"

COLOR_GREEN     = "#2ECC71"
COLOR_RED       = "#E74C3C"
COLOR_NEUTRAL   = "#D6EAF8"

COLOR_PASS_BG   = "#D1F2EB"
COLOR_FAIL_BG   = "#FADBD8"
COLOR_IDLE_BG   = "#FFFFFF"
COLOR_PASS_TEXT = "#0E6655"
COLOR_FAIL_TEXT = "#922B21"
COLOR_IDLE_TEXT = COLOR_TEXT

LABEL_OK = "PASS"
LABEL_NG = "NOT PASS"

# -------- Column widths (px) --------
COL_W_POINT = 110
COL_W_LAMP  = 80
COL_W_NUM   = 220
COL_LEFT_PAD = 6

RIGHT_TABLE_WIDTH = COL_W_POINT + COL_W_LAMP + COL_W_NUM*2 + 40  # +40 padding/scrollbar

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Resistance & Voltage Checker")
        self.configure(bg=COLOR_BG)
        self.geometry("1180x720")
        self.minsize(1080, 640)

        # icon (py / PyInstaller)
        try:
            base_dir = sys._MEIPASS if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
            ico = os.path.join(base_dir, "icon", "icon.ico")
            png = os.path.join(base_dir, "icon", "icon.png")
            if sys.platform.startswith("win") and os.path.exists(ico):
                self.iconbitmap(ico)
            elif os.path.exists(png):
                self._icon_img = tk.PhotoImage(file=png)
                self.iconphoto(True, self._icon_img)
        except Exception:
            pass

        # ----- model vars -----
        self.model_name    = tk.StringVar(value="")
        self.num_points    = tk.IntVar(value=20)   # เริ่มต้น 20 จุด
        self.auto_interval = tk.IntVar(value=500)  # ms
        self.mode          = tk.StringVar(value="manual")
        self.save_folder   = tk.StringVar(value="")
        self.auto_export   = tk.BooleanVar(value=False)  # save-after-finish in Auto mode

        # limits
        self.r_min = tk.DoubleVar(value=9.5)
        self.r_max = tk.DoubleVar(value=10.5)
        self.v_min = tk.DoubleVar(value=4.9)
        self.v_max = tk.DoubleVar(value=5.1)

        # serial placeholders
        self.ser = None
        self.com_port = tk.StringVar(value="")
        self.baudrate = tk.StringVar(value="9600")

        # data arrays
        self._init_arrays()

        # auto state
        self._auto_running = False
        self._auto_job = None

        self._setup_styles()
        self._build_ui()

    # ---------- data ----------
    def _init_arrays(self):
        n = int(self.num_points.get())
        self.r_values = [None]*n
        self.v_values = [None]*n
        self.flags    = [False]*n
        self.current_idx = 0

    # ---------- styles ----------
    def _setup_styles(self):
        style = ttk.Style()
        try: style.theme_use("clam")
        except: pass
        style.configure(".", foreground=COLOR_TEXT, font=("Segoe UI", 10))
        style.configure("Card.TFrame", background=COLOR_PANEL)
        style.configure("Card.TLabelframe", background=COLOR_PANEL)
        style.configure("Card.TLabelframe.Label", background=COLOR_PANEL, foreground=COLOR_TEXT, font=("Segoe UI", 10, "bold"))
        style.configure("Heading.TLabel", background=COLOR_PANEL, font=("Segoe UI", 11, "bold"))
        style.configure("Big.TLabel", background=COLOR_PANEL, font=("Segoe UI", 36, "bold"))
        style.configure("Muted.TLabel", background=COLOR_PANEL, foreground=COLOR_MUTED)
        style.configure("Live.TLabel", background=COLOR_PANEL, font=("Segoe UI", 18, "bold"))

    # ---------- UI ----------
    def _build_ui(self):
        nb = ttk.Notebook(self)
        self.tab_main = ttk.Frame(nb, style="Card.TFrame")
        self.tab_set  = ttk.Frame(nb, style="Card.TFrame")
        nb.add(self.tab_main, text="Main")
        nb.add(self.tab_set,  text="Setting")
        nb.pack(fill="both", expand=True, padx=12, pady=12)

        self._build_main()
        self._build_setting()

    # ---------- Main ----------
    def _build_main(self):
        root = self.tab_main
        for c in root.winfo_children(): c.destroy()

        wrapper = ttk.Frame(root, style="Card.TFrame", padding=12)
        wrapper.pack(fill="both", expand=True)

        # ---- top bar ----
        top = ttk.Frame(wrapper, style="Card.TFrame")
        top.pack(fill="x", pady=(0,8))

        ttk.Label(top, text="Model", style="Heading.TLabel").pack(side="left")
        ttk.Entry(top, textvariable=self.model_name, width=26).pack(side="left", padx=(6,16))

        ttk.Label(top, text="Point", style="Heading.TLabel").pack(side="left")
        self.point_combo = ttk.Combobox(top, state="readonly", width=5,
                                        values=[i+1 for i in range(self.num_points.get())])
        self.point_combo.current(self.current_idx)
        self.point_combo.bind("<<ComboboxSelected>>", self._on_combo)
        self.point_combo.pack(side="left", padx=(6,20))

        ttk.Label(top, text="Mode:", style="Heading.TLabel").pack(side="left", padx=(0,6))
        ttk.Radiobutton(top, text="Manual", value="manual", variable=self.mode,
                        command=self._update_mode_buttons).pack(side="left")
        ttk.Radiobutton(top, text="Auto", value="auto", variable=self.mode,
                        command=self._update_mode_buttons).pack(side="left")

        # ---- middle area: left big display + right panel ----
        mid = ttk.Frame(wrapper, style="Card.TFrame")
        mid.pack(fill="both", expand=True)

        # left big result box
        left = ttk.Frame(mid, style="Card.TFrame")
        left.pack(side="left", fill="both", expand=True, padx=(0,10))
        self.big_canvas = tk.Canvas(left, height=360, bg=COLOR_IDLE_BG,
                                    highlightthickness=1, highlightbackground=COLOR_BORDER)
        self.big_canvas.pack(fill="both", expand=True)
        self._draw_big_box()

        # --- right panel (fixed width) ---
        right_holder = ttk.Frame(mid, style="Card.TFrame")
        right_holder.pack(side="left", fill="y", padx=(0,0))

        right = ttk.Frame(right_holder, style="Card.TFrame", width=RIGHT_TABLE_WIDTH)
        right.pack(side="left", fill="y")
        right.pack_propagate(False)

        # ---- Limits (pack, no grid) ----
        limits = ttk.LabelFrame(right, text="Limits", style="Card.TLabelframe")
        limits.pack(fill="x", pady=(0,10), padx=(0,2))

        row1 = ttk.Frame(limits, style="Card.TFrame"); row1.pack(anchor="w", pady=2)
        ttk.Label(row1, text="R", width=2, style="Heading.TLabel").pack(side="left", padx=(0,6))
        ttk.Label(row1, text="Min", width=4, style="Muted.TLabel").pack(side="left")
        self.lbl_rmin = ttk.Label(row1, text=f"{self.r_min.get():.3g}", width=8, anchor="center"); self.lbl_rmin.pack(side="left")
        ttk.Label(row1, text="Max", width=4, style="Muted.TLabel").pack(side="left", padx=(10,0))
        self.lbl_rmax = ttk.Label(row1, text=f"{self.r_max.get():.3g}", width=8, anchor="center"); self.lbl_rmax.pack(side="left")
        ttk.Label(row1, text="Ω", style="Muted.TLabel").pack(side="left", padx=(6,0))

        row2 = ttk.Frame(limits, style="Card.TFrame"); row2.pack(anchor="w", pady=2)
        ttk.Label(row2, text="V", width=2, style="Heading.TLabel").pack(side="left", padx=(0,6))
        ttk.Label(row2, text="Min", width=4, style="Muted.TLabel").pack(side="left")
        self.lbl_vmin = ttk.Label(row2, text=f"{self.v_min.get():.3g}", width=8, anchor="center"); self.lbl_vmin.pack(side="left")
        ttk.Label(row2, text="Max", width=4, style="Muted.TLabel").pack(side="left", padx=(10,0))
        self.lbl_vmax = ttk.Label(row2, text=f"{self.v_max.get():.3g}", width=8, anchor="center"); self.lbl_vmax.pack(side="left")
        ttk.Label(row2, text="V", style="Muted.TLabel").pack(side="left", padx=(6,0))

        # ---- Table (header + scrollable rows) ----
        table = ttk.Frame(right, style="Card.TFrame", padding=(0,2))
        table.pack(fill="both", expand=True)

        hdr = ttk.Frame(table, style="Card.TFrame")
        hdr.pack(fill="x", pady=(0,2))
        self._apply_col_layout(hdr)
        ttk.Label(hdr, text="Point",  style="Heading.TLabel", anchor="w").grid(row=0, column=0, sticky="w", padx=(COL_LEFT_PAD,2))
        ttk.Label(hdr, text="Active", style="Heading.TLabel", anchor="center").grid(row=0, column=1, sticky="ew")
        ttk.Label(hdr, text="R (Ω)",  style="Heading.TLabel", anchor="center").grid(row=0, column=2, sticky="ew")
        ttk.Label(hdr, text="V (V)",  style="Heading.TLabel", anchor="center").grid(row=0, column=3, sticky="ew")

        sc_wrapper = ttk.Frame(table, style="Card.TFrame")
        sc_wrapper.pack(fill="both", expand=True)

        # Canvas: ใช้ fill="both" เพื่อให้กว้างตามกรอบ และสูงตามพื้นที่ที่เหลือ
        self.points_canvas = tk.Canvas(
            sc_wrapper, bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDER
        )
        vbar = ttk.Scrollbar(sc_wrapper, orient="vertical", command=self.points_canvas.yview)
        self.points_canvas.configure(yscrollcommand=vbar.set)
        self.points_canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        # ปุ่มเลื่อน (optional)
        #btns = ttk.Frame(right, style="Card.TFrame")
        #btns.pack(fill="x", pady=(6,0))
        #ttk.Button(btns, text="▲ Up", width=6, command=lambda: self.points_canvas.yview_scroll(-3, "units")).pack(side="left", padx=(0,6))
        #ttk.Button(btns, text="▼ Down", width=6, command=lambda: self.points_canvas.yview_scroll( 3, "units")).pack(side="left")

        # inner frame
        self.points_frame = ttk.Frame(self.points_canvas, style="Card.TFrame")
        self.points_window = self.points_canvas.create_window((0,0), window=self.points_frame, anchor="nw")

        # อัปเดต scrollregion เสมอ (รวมทั้งหลังสร้างเสร็จ)
        self.points_frame.bind("<Configure>", self._update_points_scroll)
        self.points_canvas.bind("<Configure>", self._update_points_scroll)
        self.after_idle(self._update_points_scroll)

        # bind mouse wheel กับทั้ง frame/canvas/wrapper
        self._bind_scrolling(self.points_frame, self.points_canvas)
        self._bind_scrolling(self.points_canvas, self.points_canvas)
        self._bind_scrolling(sc_wrapper, self.points_canvas)

        # rows
        self.row_widgets = []
        for i in range(self.num_points.get()):
            rowf = ttk.Frame(self.points_frame, style="Card.TFrame")
            rowf.grid_columnconfigure(0, minsize=COL_W_POINT, weight=0)
            rowf.grid_columnconfigure(1, minsize=COL_W_LAMP,  weight=0)
            rowf.grid_columnconfigure(2, minsize=COL_W_NUM,   weight=0)
            rowf.grid_columnconfigure(3, minsize=COL_W_NUM,   weight=0)
            rowf.pack(fill="x", pady=4)

            lbl = ttk.Label(rowf, text=f"Point {i+1}")
            lbl.grid(row=0, column=0, sticky="w", padx=(COL_LEFT_PAD,2))
            lbl.bind("<Button-1>", lambda e, idx=i: self._jump_to(idx))

            lamp = tk.Label(rowf, width=8, height=1, bg=COLOR_NEUTRAL, relief="solid", bd=1)
            lamp.grid(row=0, column=1, sticky="ew", padx=4)

            r_var = tk.StringVar(value="")
            r_ent = tk.Entry(rowf, textvariable=r_var, width=20, justify="center",
                             relief="solid", bd=1)
            r_ent.configure(state="readonly", readonlybackground="white")
            r_ent.grid(row=0, column=2, sticky="ew", padx=4, ipady=2)

            v_var = tk.StringVar(value="")
            v_ent = tk.Entry(rowf, textvariable=v_var, width=20, justify="center",
                             relief="solid", bd=1)
            v_ent.configure(state="readonly", readonlybackground="white")
            v_ent.grid(row=0, column=3, sticky="ew", padx=4, ipady=2)

            self.row_widgets.append((lamp, r_var, r_ent, v_var, v_ent))

        # ---- bottom controls ----
        sep = ttk.Separator(wrapper, orient="horizontal")
        sep.pack(fill="x", pady=(10,8))

        ctrl = ttk.Frame(wrapper, style="Card.TFrame")
        ctrl.pack(fill="x")

        self.btn_measure = ttk.Button(ctrl, text="Measure (Manual)", command=self._measure_one)
        self.btn_measure.pack(side="left", padx=(0,8), ipady=2)

        self.lbl_ohm = ttk.Label(ctrl, text="— Ω", style="Live.TLabel"); self.lbl_ohm.pack(side="left", padx=20)
        self.lbl_volt= ttk.Label(ctrl, text="— V",  style="Live.TLabel"); self.lbl_volt.pack(side="left", padx=20)

        self.btn_auto_start = ttk.Button(ctrl, text="Start (Auto)", command=self._auto_start)
        self.btn_auto_stop  = ttk.Button(ctrl, text="Stop", command=self._auto_stop)
        self.btn_auto_start.pack(side="left", padx=(20,6), ipady=2)
        self.btn_auto_stop.pack(side="left", ipady=2)

        self.btn_auto_export = ttk.Button(ctrl, text="Auto Export: OFF", command=self._toggle_auto_export)
        self.btn_auto_export.pack(side="left", padx=(20,0), ipady=2)

        # Save folder row
        save_row = ttk.Frame(wrapper, style="Card.TFrame")
        save_row.pack(fill="x", pady=(10,0))
        ttk.Label(save_row, text="Save Folder", style="Heading.TLabel").pack(side="left")
        self.ent_folder = ttk.Entry(save_row, textvariable=self.save_folder, width=70)
        self.ent_folder.pack(side="left", padx=(8,6), fill="x", expand=True)
        ttk.Button(save_row, text="Browse", command=self._browse_folder).pack(side="left")
        ttk.Label(save_row, text="(optional)", style="Muted.TLabel").pack(side="left", padx=6)

        # bottom bar
        bottom = ttk.Frame(wrapper, style="Card.TFrame")
        bottom.pack(fill="x", pady=(10,0))
        ttk.Button(bottom, text="Clear Results", command=self._reset_results).pack(side="left")
        ttk.Button(bottom, text="Reset All",    command=self._reset).pack(side="left", padx=(8,0))
        ttk.Button(bottom, text="Export (.txt)", command=self._manual_export).pack(side="left", padx=(10,0))

        self._update_mode_buttons()
        self._refresh_rows()
        self._update_big_box()

    # ---------- Setting ----------
    def _build_setting(self):
        frm = ttk.Frame(self.tab_set, style="Card.TFrame", padding=14)
        frm.pack(fill="both", expand=True)

        frm.grid_columnconfigure(0, weight=1)
        frm.grid_columnconfigure(1, weight=1)

        # ========== Measurement ==========
        meas = ttk.Labelframe(frm, text="Measurement", style="Card.TLabelframe", padding=12)
        meas.grid(row=0, column=0, sticky="nsew", padx=(0,8))

        r0 = ttk.Frame(meas, style="Card.TFrame"); r0.pack(anchor="w", pady=6, fill="x")
        ttk.Label(r0, text="Point Count", width=16).pack(side="left")
        ttk.Spinbox(r0, from_=1, to=999, textvariable=self.num_points, width=8, justify="right").pack(side="left", padx=(8,16))
        ttk.Label(r0, text="Auto Interval", width=14).pack(side="left")
        ttk.Spinbox(r0, from_=10, to=100000, increment=10,
                    textvariable=self.auto_interval, width=8, justify="right").pack(side="left", padx=(8,6))
        ttk.Label(r0, text="ms", style="Muted.TLabel").pack(side="left")

        r1 = ttk.Frame(meas, style="Card.TFrame"); r1.pack(anchor="w", pady=4, fill="x")
        ttk.Label(r1, text="R Limits", width=16).pack(side="left")
        ttk.Entry(r1, textvariable=self.r_min, width=10, justify="center").pack(side="left", padx=(8,6))
        ttk.Label(r1, text="–", style="Muted.TLabel").pack(side="left")
        ttk.Entry(r1, textvariable=self.r_max, width=10, justify="center").pack(side="left", padx=(6,6))
        ttk.Label(r1, text="Ω", style="Muted.TLabel").pack(side="left")

        r2 = ttk.Frame(meas, style="Card.TFrame"); r2.pack(anchor="w", pady=4, fill="x")
        ttk.Label(r2, text="V Limits", width=16).pack(side="left")
        ttk.Entry(r2, textvariable=self.v_min, width=10, justify="center").pack(side="left", padx=(8,6))
        ttk.Label(r2, text="–", style="Muted.TLabel").pack(side="left")
        ttk.Entry(r2, textvariable=self.v_max, width=10, justify="center").pack(side="left", padx=(6,6))
        ttk.Label(r2, text="V", style="Muted.TLabel").pack(side="left")

        ttk.Button(meas, text="Apply", command=self._apply_settings).pack(anchor="w", pady=(10,2))
        ttk.Label(meas, text="R_min ≤ R ≤ R_max และ V_min ≤ V ≤ V_max", style="Muted.TLabel").pack(anchor="w")

        # ========== Instrument I/O ==========
        io = ttk.Labelframe(frm, text="Instrument I/O", style="Card.TLabelframe", padding=12)
        io.grid(row=0, column=1, sticky="nsew", padx=(8,0))

        rp = ttk.Frame(io, style="Card.TFrame"); rp.pack(fill="x", pady=6)
        ttk.Label(rp, text="COM Port", width=12).pack(side="left")
        self.combo_port = ttk.Combobox(rp, textvariable=self.com_port, width=18, state="readonly")
        self.combo_port.pack(side="left", padx=(8,8))
        ttk.Button(rp, text="Refresh", command=self._refresh_com_ports).pack(side="left")

        rb = ttk.Frame(io, style="Card.TFrame"); rb.pack(fill="x", pady=6)
        ttk.Label(rb, text="Baudrate", width=12).pack(side="left")
        self.combo_baud = ttk.Combobox(rb, textvariable=self.baudrate, width=18, state="readonly",
                                       values=[4800, 9600, 19200, 38400, 57600, 115200, 230400])
        if not self.combo_baud.get(): self.combo_baud.set("9600")
        self.combo_baud.pack(side="left", padx=(8,8))

        rc = ttk.Frame(io, style="Card.TFrame"); rc.pack(fill="x", pady=6)
        self.btn_connect = ttk.Button(rc, text="Connect", command=self._connect_serial); self.btn_connect.pack(side="left")
        self.btn_disconnect = ttk.Button(rc, text="Disconnect", command=self._disconnect_serial); self.btn_disconnect.pack(side="left", padx=8)
        ttk.Button(rc, text="Test Read", command=lambda: messagebox.showinfo(
            "Test Read",
            f"Simulated read: R={self._read_meter()[0]:.3f} Ω, V={self._read_meter()[1]:.3f} V\n\n"
            "(*จะอ่านจริงเมื่อคุณใส่โปรโตคอลกับเครื่องวัดแล้ว)")).pack(side="left", padx=(8,0))

        self.lbl_conn = ttk.Label(io, text="Status: Disconnected", style="Muted.TLabel")
        self.lbl_conn.pack(anchor="w", pady=(8,0))

        self._refresh_com_ports()
        self._update_serial_buttons()

    # ---------- settings apply ----------
    def _apply_settings(self):
        try:
            new_n  = int(self.num_points.get());  assert new_n > 0
            _ = float(self.r_min.get()); _ = float(self.r_max.get())
            _ = float(self.v_min.get()); _ = float(self.v_max.get())
            iv = int(self.auto_interval.get()); self.auto_interval.set(max(10, iv))
        except Exception as e:
            messagebox.showerror("Invalid", f"Settings error: {e}"); return

        need_rebuild = (new_n != len(self.r_values))
        if need_rebuild:
            self._auto_stop()
            self._init_arrays()
            self._build_main()
        else:
            self._refresh_rows()
            self._update_big_box()

        messagebox.showinfo("Apply", "Settings applied.")

    # --------- Serial helpers ---------
    def _refresh_com_ports(self):
        ports = []
        try:
            from serial.tools import list_ports
            raw = [p.device for p in list_ports.comports()]
            # sort แบบตัวเลขท้าย "COM"
            ports = sorted(
                raw,
                key=lambda x: (x[:3]!="COM", int(x[3:]) if x.startswith("COM") and x[3:].isdigit() else 0, x)
            )
        except Exception as e:
            print("List COM error:", e)
        self.combo_port["values"] = ports
        if self.com_port.get() not in ports:
            self.com_port.set(ports[0] if ports else "")

    def _connect_serial(self):
        try:
            import serial
            port = self.com_port.get().strip()
            if not port:
                messagebox.showwarning("Serial", "Please select a COM Port."); return
            baud = int(self.baudrate.get())
            if self.ser:
                try: self.ser.close()
                except: pass
            self.ser = serial.Serial(port=port, baudrate=baud, timeout=1)
            self.lbl_conn.config(text=f"Status: Connected to {port} @ {baud} bps")
            messagebox.showinfo("Serial", f"Connected to {port} @ {baud} bps")
        except Exception as e:
            self.ser = None
            self.lbl_conn.config(text="Status: Disconnected")
            messagebox.showerror("Serial", f"Connect failed:\n{e}")
        self._update_serial_buttons()

    def _disconnect_serial(self):
        try:
            if self.ser: self.ser.close()
        except: pass
        self.ser = None
        self.lbl_conn.config(text="Status: Disconnected")
        self._update_serial_buttons()
        messagebox.showinfo("Serial", "Disconnected")

    def _update_serial_buttons(self):
        if getattr(self, "btn_connect", None) is None:
            return
        if self.ser:
            self.btn_connect.state(["disabled"])
            self.btn_disconnect.state(["!disabled"])
        else:
            self.btn_connect.state(["!disabled"])
            self.btn_disconnect.state(["disabled"])

    # ---------- helpers ----------
    def _apply_col_layout(self, frame):
        frame.grid_columnconfigure(0, minsize=COL_W_POINT, weight=0)  # Point
        frame.grid_columnconfigure(1, minsize=COL_W_LAMP,  weight=0)  # Active
        frame.grid_columnconfigure(2, minsize=COL_W_NUM,   weight=0)  # R
        frame.grid_columnconfigure(3, minsize=COL_W_NUM,   weight=0)  # V

    def _update_mode_buttons(self):
        is_auto = (self.mode.get() == "auto")
        self.btn_measure.state(["disabled" if is_auto else "!disabled"])
        self.btn_auto_start.state(["!disabled" if is_auto else "disabled"])
        self.btn_auto_stop.state(["!disabled" if is_auto else "disabled"])

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Choose folder to save")
        if path:
            self.save_folder.set(path)

    def _on_combo(self, _e):
        self.current_idx = int(self.point_combo.get()) - 1
        self._update_big_box()
        self._scroll_row_into_view(self.current_idx)

    def _jump_to(self, idx):
        self.current_idx = idx
        self.point_combo.current(idx)
        self._update_big_box()
        self._scroll_row_into_view(idx)

    def _update_points_scroll(self, _e=None):
        # อัปเดตขอบเขตเลื่อนให้ครอบคลุม inner frame
        try:
            self.points_canvas.configure(scrollregion=self.points_canvas.bbox("all"))
            # ปรับความกว้างของ inner window ให้เท่ากับ canvas กว้าง
            self.points_canvas.itemconfigure(self.points_window, width=self.points_canvas.winfo_width())
        except Exception:
            pass

    def _bind_scrolling(self, widget, canvas):
        # รองรับ Windows/Mac
        def on_wheel(e):
            delta = e.delta
            if delta == 0: return
            canvas.yview_scroll(-1 if delta > 0 else 1, "units")
        widget.bind("<Enter>", lambda _: widget.bind_all("<MouseWheel>", on_wheel))
        widget.bind("<Leave>", lambda _: widget.unbind_all("<MouseWheel>"))
        # รองรับ Linux
        widget.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        widget.bind_all("<Button-5>", lambda e: canvas.yview_scroll( 1, "units"))

    def _scroll_row_into_view(self, idx, margin=8):
        if not (0 <= idx < len(self.row_widgets)): return
        self.points_canvas.update_idletasks()
        rows = self.points_frame.winfo_children()
        rowf = rows[idx]
        y = rowf.winfo_y()
        h = rowf.winfo_height()
        top = self.points_canvas.canvasy(0)
        bottom = top + self.points_canvas.winfo_height()
        if y < top + margin:
            new_top = max(0, y - margin)
        elif (y + h) > (bottom - margin):
            new_top = (y + h + margin) - self.points_canvas.winfo_height()
        else:
            return
        total = max(1, self.points_frame.winfo_height())
        self.points_canvas.yview_moveto(max(0, new_top)/total)

    # ---------- rendering ----------
    def _draw_big_box(self):
        self.big_canvas.delete("all")
        r = self.r_values[self.current_idx]
        v = self.v_values[self.current_idx]
        if r is None and v is None:
            bg, fg, text = COLOR_IDLE_BG, COLOR_IDLE_TEXT, ""
        else:
            ok_r = (self.r_min.get() <= (r if r is not None else 1e9) <= self.r_max.get())
            ok_v = (self.v_min.get() <= (v if v is not None else 1e9) <= self.v_max.get())
            is_pass = ok_r and ok_v
            bg = COLOR_PASS_BG if is_pass else COLOR_FAIL_BG
            fg = COLOR_PASS_TEXT if is_pass else COLOR_FAIL_TEXT
            text = LABEL_OK if is_pass else LABEL_NG
        self.big_canvas.configure(bg=bg)
        w = self.big_canvas.winfo_width()  or int(self.big_canvas["width"])
        h = self.big_canvas.winfo_height() or int(self.big_canvas["height"])
        self.big_canvas.create_text(w/2, h/2, text=text, fill=fg, font=("Segoe UI", 60, "bold"))

    def _refresh_rows(self):
        for i,(lamp, r_var, r_ent, v_var, v_ent) in enumerate(self.row_widgets):
            r = self.r_values[i]
            v = self.v_values[i]

            r_var.set("" if r is None else f"{r:.3g}")
            v_var.set("" if v is None else f"{v:.3g}")

            ok_r = (r is not None) and (self.r_min.get() <= r <= self.r_max.get())
            ok_v = (v is not None) and (self.v_min.get() <= v <= self.v_max.get())
            is_pass = ok_r and ok_v
            self.flags[i] = (r is not None) and (v is not None) and is_pass

            lamp.configure(bg=COLOR_GREEN if is_pass else (COLOR_RED if (r is not None or v is not None) else COLOR_NEUTRAL))

            r_ent.configure(state="normal", fg=("red" if (r is not None and not ok_r) else COLOR_TEXT))
            v_ent.configure(state="normal", fg=("red" if (v is not None and not ok_v) else COLOR_TEXT))
            r_ent.configure(state="readonly")
            v_ent.configure(state="readonly")

    def _update_big_box(self):
        r = self.r_values[self.current_idx]
        v = self.v_values[self.current_idx]
        self.lbl_ohm.config(text=("— Ω" if r is None else f"{r:.4g} Ω"))
        self.lbl_volt.config(text=("— V"  if v is None else f"{v:.4g} V"))
        self._draw_big_box()

    # ---------- measurement ----------
    def _measure_one(self, from_auto: bool = False):
        idx = self.current_idx
        r, v = self._read_meter()   # TODO: replace with real meter reading
        self.r_values[idx] = r
        self.v_values[idx] = v
        self._refresh_rows()
        self._update_big_box()

        # เดินหน้าจุดถัดไป หรือจบการวัด
        if self.current_idx < self.num_points.get() - 1:
            self.current_idx += 1
            self.point_combo.current(self.current_idx)
            self._scroll_row_into_view(self.current_idx)
        else:
            # มาถึงจุดสุดท้ายแล้ว
            if from_auto:
                self._auto_running = False
                if self._auto_job is not None:
                    self.after_cancel(self._auto_job)
                    self._auto_job = None
                self._update_mode_buttons()
                # เซฟทีเดียวเมื่อครบทุกจุด ถ้าเปิด Auto Export
                if self.auto_export.get():
                    self._export_snapshot()
                self.after(0, lambda: messagebox.showinfo("Auto", "Auto measurement finished."))
            else:
                messagebox.showinfo("Done", "Measured all points.")

    def _auto_start(self):
        if self._auto_running:
            return
        self._auto_running = True
        self._update_mode_buttons()
        self._tick_auto()

    def _auto_stop(self):
        self._auto_running = False
        if self._auto_job is not None:
            self.after_cancel(self._auto_job)
            self._auto_job = None
        self._update_mode_buttons()

    def _tick_auto(self):
        if not self._auto_running:
            return
        self._measure_one(from_auto=True)
        if self._auto_running:
            interval = max(50, int(self.auto_interval.get()))
            self._auto_job = self.after(interval, self._tick_auto)

    # Simulated meter read (replace later with real IO)
    def _read_meter(self):
        r_center = (self.r_min.get()+self.r_max.get())/2
        v_center = (self.v_min.get()+self.v_max.get())/2
        r_span = (self.r_max.get()-self.r_min.get())/2
        v_span = (self.v_max.get()-self.v_min.get())/2
        if random.random() < 0.8:
            r = random.uniform(r_center-0.6*r_span, r_center+0.6*r_span)
            v = random.uniform(v_center-0.6*v_span, v_center+0.6*v_span)
        else:
            r = r_center + random.choice([-1,1]) * random.uniform(0.7*r_span, 1.6*r_span)
            v = v_center + random.choice([-1,1]) * random.uniform(0.7*v_span, 1.6*v_span)
        return float(r), float(v)

    # ---------- export ----------
    def _toggle_auto_export(self):
        self.auto_export.set(not self.auto_export.get())
        self.btn_auto_export.config(text=f"Auto Export: {'ON' if self.auto_export.get() else 'OFF'}")

    def _result_lines(self):
        lines = []
        lines.append(f"Model\t{self.model_name.get().strip() or '-'}")
        lines.append(f"Time\t{time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("")
        lines.append(f"Limits R\tmin={self.r_min.get()}\tmax={self.r_max.get()}  (Ohm)")
        lines.append(f"Limits V\tmin={self.v_min.get()}\tmax={self.v_max.get()}  (Volt)")
        lines.append("")
        lines.append("Point\tR(Ω)\tV(V)\tResult")
        for i in range(self.num_points.get()):
            r = self.r_values[i]; v = self.v_values[i]
            if r is None or v is None:
                lines.append(f"{i+1}\t\t\tN/A")
            else:
                res = LABEL_OK if self.flags[i] else LABEL_NG
                lines.append(f"{i+1}\t{r:.6f}\t{v:.6f}\t{res}")
        overall = LABEL_OK if all((rv is not None) for rv in self.r_values) and all(self.flags) else LABEL_NG
        lines.append("")
        lines.append(f"Overall\t{overall}")
        return lines

    def _export_snapshot(self):
        folder = self.save_folder.get().strip() or os.getcwd()
        os.makedirs(folder, exist_ok=True)
        ts = time.strftime("%Y%m%d_%H%M%S")
        name = f"{self.model_name.get().strip() or 'model'}_{ts}.txt"
        path = os.path.join(folder, name)
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(self._result_lines()))
            messagebox.showinfo("Export", f"Saved:\n{path}")
        except Exception as e:
            messagebox.showerror("Export", f"Auto export failed:\n{e}")

    def _manual_export(self):
        folder = self.save_folder.get().strip() or os.getcwd()
        os.makedirs(folder, exist_ok=True)
        ts = time.strftime("%Y%m%d_%H%M%S")
        name = f"{self.model_name.get().strip() or 'model'}_{ts}.txt"
        path = filedialog.asksaveasfilename(
            title="Save results",
            initialdir=folder,
            initialfile=name,
            defaultextension=".txt",
            filetypes=[("Text file","*.txt")]
        )
        if not path: return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(self._result_lines()))
            messagebox.showinfo("Export", "Saved successfully.")
        except Exception as e:
            messagebox.showerror("Export", f"Save failed:\n{e}")

    # ---------- reset ----------
    def _reset_results(self):
        """ล้างเฉพาะผลการวัดทั้งหมด คงโหมด/Auto-Export/โฟลเดอร์ไว้"""
        if self._auto_running:
            self._auto_stop()
        n = len(self.r_values)
        self.r_values = [None]*n
        self.v_values = [None]*n
        self.flags    = [False]*n
        self.current_idx = 0
        try: self.point_combo.current(0)
        except: pass
        self._refresh_rows()
        self._update_big_box()
        try: self.points_canvas.yview_moveto(0.0)
        except: pass

    def _reset(self):
        self._auto_stop()
        self._init_arrays()
        self._build_main()

if __name__ == "__main__":
    App().mainloop()
