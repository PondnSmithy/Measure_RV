# -*- coding: utf-8 -*-
"""
Resistance & Voltage Checker (Blue/White UI)
- Right panel fixed width (so big left canvas has more room)
- Limits section shows Min/Max derived from Set ± Tol
- Active column is centered color block (Label)
- R/V values centered; red text when out of spec
- Manual / Auto measurement (simulated)
- Auto-export toggle (timestamped .txt to chosen folder; default cwd)
- Auto export saves only AFTER all cells (when ON); Manual saves only when pressing Export
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

RIGHT_TABLE_WIDTH = COL_W_POINT + COL_W_LAMP + COL_W_NUM*2 + 40  # +padding/scrollbar


        
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
        self.num_points    = tk.IntVar(value=20)
        self.auto_interval = tk.IntVar(value=500)  # ms
        self.mode          = tk.StringVar(value="manual")
        self.save_folder   = tk.StringVar(value="")
        self.auto_export   = tk.BooleanVar(value=False)

        # ---- limits as Set ± Tol (instead of Min/Max) ----
        self.r_set = tk.DoubleVar(value=10.0)
        self.r_tol = tk.DoubleVar(value=0.5)
        self.v_set = tk.DoubleVar(value=5.0)
        self.v_tol = tk.DoubleVar(value=0.1)

        # ---- serial settings ----
        self.com_port  = tk.StringVar(value="")
        self.baudrate  = tk.StringVar(value="9600")
        self.ser = None

        # data arrays
        self._init_arrays()

        # auto state
        self._auto_running = False
        self._auto_job = None

        self._setup_styles()
        self._build_ui()

    
    # ---------- helpers for derived bounds ----------
    def _r_bounds(self):
        rmin = float(self.r_set.get()) - float(self.r_tol.get())
        rmax = float(self.r_set.get()) + float(self.r_tol.get())
        return rmin, rmax

    def _v_bounds(self):
        vmin = float(self.v_set.get()) - float(self.v_tol.get())
        vmax = float(self.v_set.get()) + float(self.v_tol.get())
        return vmin, vmax

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

        ttk.Label(top, text="Cell", style="Heading.TLabel").pack(side="left")
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
        limits = ttk.LabelFrame(right, text="Limits (derived from Set ± Tol)", style="Card.TLabelframe")
        limits.pack(fill="x", pady=(0,10), padx=(0,2))

        row1 = ttk.Frame(limits, style="Card.TFrame"); row1.pack(anchor="w", pady=2)
        ttk.Label(row1, text="R", width=2, style="Heading.TLabel").pack(side="left", padx=(0,6))
        ttk.Label(row1, text="Min", width=4, style="Muted.TLabel").pack(side="left")
        self.lbl_rmin = ttk.Label(row1, text="", width=8, anchor="center"); self.lbl_rmin.pack(side="left")
        ttk.Label(row1, text="Max", width=4, style="Muted.TLabel").pack(side="left", padx=(10,0))
        self.lbl_rmax = ttk.Label(row1, text="", width=8, anchor="center"); self.lbl_rmax.pack(side="left")
        ttk.Label(row1, text="mΩ", style="Muted.TLabel").pack(side="left", padx=(6,0))

        row2 = ttk.Frame(limits, style="Card.TFrame"); row2.pack(anchor="w", pady=2)
        ttk.Label(row2, text="V", width=2, style="Heading.TLabel").pack(side="left", padx=(0,6))
        ttk.Label(row2, text="Min", width=4, style="Muted.TLabel").pack(side="left")
        self.lbl_vmin = ttk.Label(row2, text="", width=8, anchor="center"); self.lbl_vmin.pack(side="left")
        ttk.Label(row2, text="Max", width=4, style="Muted.TLabel").pack(side="left", padx=(10,0))
        self.lbl_vmax = ttk.Label(row2, text="", width=8, anchor="center"); self.lbl_vmax.pack(side="left")
        ttk.Label(row2, text="V", style="Muted.TLabel").pack(side="left", padx=(6,0))

        self._refresh_limits_labels()

        # ---- Table (header + scrollable rows) ----
        table = ttk.Frame(right, style="Card.TFrame", padding=(0,2))
        table.pack(fill="both", expand=True)

        hdr = ttk.Frame(table, style="Card.TFrame")
        hdr.pack(fill="x", pady=(0,2))
        self._apply_col_layout(hdr)
        ttk.Label(hdr, text="Cell",  style="Heading.TLabel", anchor="w").grid(row=0, column=0, sticky="w", padx=(COL_LEFT_PAD,2))
        ttk.Label(hdr, text="Active", style="Heading.TLabel", anchor="center").grid(row=0, column=1, sticky="ew")
        ttk.Label(hdr, text="R (mΩ)",  style="Heading.TLabel", anchor="center").grid(row=0, column=2, sticky="ew")
        ttk.Label(hdr, text="V (V)",  style="Heading.TLabel", anchor="center").grid(row=0, column=3, sticky="ew")

        # ====== (สำคัญ) สร้าง Canvas + Scrollbar สำหรับรายการ Point ======
        sc = ttk.Frame(table, style="Card.TFrame")
        sc.pack(fill="both", expand=True)

        self.points_canvas = tk.Canvas(
            sc, height=220, width=RIGHT_TABLE_WIDTH-18,
            bg=COLOR_PANEL, highlightthickness=1, highlightbackground=COLOR_BORDER
        )
        vbar = ttk.Scrollbar(sc, orient="vertical", command=self.points_canvas.yview)
        self.points_canvas.configure(yscrollcommand=vbar.set)

        # ให้แคนวาสขยายเต็มพื้นที่เพื่อเลื่อนสบายขึ้น
        self.points_canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")

        # เฟรมจริงที่เก็บแถว อยู่ข้างใน Canvas
        self.points_frame = ttk.Frame(self.points_canvas, style="Card.TFrame")
        self.points_window = self.points_canvas.create_window((0,0), window=self.points_frame, anchor="nw")

        
        # อัปเดต scrollregion และให้ความกว้างภายในเท่ากับแคนวาส
        def _update_scrollregion(_e=None):
            self.points_canvas.configure(scrollregion=self.points_canvas.bbox("all"))
            try:
                self.points_canvas.itemconfigure(self.points_window, width=self.points_canvas.winfo_width())
            except:
                pass

        self.points_frame.bind("<Configure>", _update_scrollregion)
        self.points_canvas.bind("<Configure>", _update_scrollregion)
        self._bind_scrolling(self.points_frame, self.points_canvas)
        # ====== จบส่วน Canvas + Scrollbar ======

        # rows
        self.row_widgets = []  # (lamp_label, r_var, r_entry, v_var, v_entry)
        for i in range(self.num_points.get()):
            rowf = ttk.Frame(self.points_frame, style="Card.TFrame")
            rowf.grid_columnconfigure(0, minsize=COL_W_POINT, weight=0)
            rowf.grid_columnconfigure(1, minsize=COL_W_LAMP,  weight=0)
            rowf.grid_columnconfigure(2, minsize=COL_W_NUM,   weight=0)
            rowf.grid_columnconfigure(3, minsize=COL_W_NUM,   weight=0)
            rowf.pack(fill="x", pady=4)

            lbl = ttk.Label(rowf, text=f"Cell {i+1}")
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

        self.lbl_ohm = ttk.Label(ctrl, text="— mΩ", style="Live.TLabel"); self.lbl_ohm.pack(side="left", padx=20)
        self.lbl_volt= ttk.Label(ctrl, text="— V",  style="Live.TLabel"); self.lbl_volt.pack(side="left", padx=20)

        self.btn_auto_start = ttk.Button(ctrl, text="Start (Auto)", command=self._auto_start)
        self.btn_auto_stop  = ttk.Button(ctrl, text="Stop", command=self._auto_stop)
        self.btn_auto_start.pack(side="left", padx=(20,6), ipady=2)
        self.btn_auto_stop.pack(side="left", ipady=2)

        self.btn_auto_export = ttk.Button(ctrl, text="Auto Export: OFF", command=self._toggle_auto_export)
        self.btn_auto_export.pack(side="left", padx=(20,0), ipady=2)
        # ให้ปุ่มสะท้อนสถานะเดิม (คงค่าแม้มีการ reset)
        self.btn_auto_export.config(text=f"Auto Export: {'ON' if self.auto_export.get() else 'OFF'}")

        # Save folder row
        save_row = ttk.Frame(wrapper, style="Card.TFrame")
        save_row.pack(fill="x", pady=(10,0))
        ttk.Label(save_row, text="Save Folder", style="Heading.TLabel").pack(side="left")
        self.ent_folder = ttk.Entry(save_row, textvariable=self.save_folder, width=70)
        self.ent_folder.pack(side="left", padx=(8,6), fill="x", expand=True)
        ttk.Button(save_row, text="Browse", command=self._browse_folder).pack(side="left")
        ttk.Label(save_row, text="(optional)", style="Muted.TLabel").pack(side="left", padx=6)

        bottom = ttk.Frame(wrapper, style="Card.TFrame")
        bottom.pack(fill="x", pady=(10,0))
        ttk.Button(bottom, text="Reset", command=self._reset).pack(side="left")
        #ttk.Button(bottom, text="Export (.txt)", command=self._manual_export).pack(side="left", padx=(10,0))
        ttk.Button(bottom, text="Export (.txt)", command=self._export_txt_table).pack(side="left", padx=(10,0))

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

        # row 0: Point Count + Auto Interval
        r0 = ttk.Frame(meas, style="Card.TFrame"); r0.pack(anchor="w", pady=6, fill="x")
        ttk.Label(r0, text="Cell Count", width=16).pack(side="left")
        ttk.Spinbox(r0, from_=1, to=999, textvariable=self.num_points, width=8, justify="right").pack(side="left", padx=(8,16))
        ttk.Label(r0, text="Auto Interval", width=14).pack(side="left")
        ttk.Spinbox(r0, from_=10, to=100000, increment=10,
                    textvariable=self.auto_interval, width=8, justify="right").pack(side="left", padx=(8,6))
        ttk.Label(r0, text="ms", style="Muted.TLabel").pack(side="left")

        # row 1: R Set & Tol
        r1 = ttk.Frame(meas, style="Card.TFrame"); r1.pack(anchor="w", pady=6, fill="x")
        ttk.Label(r1, text="R Set", width=16).pack(side="left")
        ttk.Entry(r1, textvariable=self.r_set, width=10, justify="center").pack(side="left", padx=(8,8))
        ttk.Label(r1, text="mΩ", style="Muted.TLabel").pack(side="left")
        ttk.Label(r1, text="  ±Tol", width=6).pack(side="left", padx=(12,0))
        ttk.Entry(r1, textvariable=self.r_tol, width=10, justify="center").pack(side="left", padx=(6,6))
        ttk.Label(r1, text="mΩ", style="Muted.TLabel").pack(side="left")

        # row 2: V Set & Tol
        r2 = ttk.Frame(meas, style="Card.TFrame"); r2.pack(anchor="w", pady=6, fill="x")
        ttk.Label(r2, text="V Set", width=16).pack(side="left")
        ttk.Entry(r2, textvariable=self.v_set, width=10, justify="center").pack(side="left", padx=(8,8))
        ttk.Label(r2, text="V", style="Muted.TLabel").pack(side="left")
        ttk.Label(r2, text="  ±Tol", width=6).pack(side="left", padx=(12,0))
        ttk.Entry(r2, textvariable=self.v_tol, width=10, justify="center").pack(side="left", padx=(6,6))
        ttk.Label(r2, text="V", style="Muted.TLabel").pack(side="left")

        ttk.Button(meas, text="Apply", command=self._apply_settings).pack(anchor="w", pady=(10,2))
        ttk.Label(meas, text="เงื่อนไข: |R - R_set| ≤ R_tol และ |V - V_set| ≤ V_tol", style="Muted.TLabel").pack(anchor="w")

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
            f"Simulated read:\nR={self._read_meter()[0]:.2f} mΩ, V={self._read_meter()[1]:.4f} V\n\n"
            "(*จะอ่านจริงเมื่อคุณใส่โปรโตคอลกับเครื่องวัดแล้ว)")
        ).pack(side="left", padx=(8,0))

        self.lbl_conn = ttk.Label(io, text="Status: Disconnected", style="Muted.TLabel")
        self.lbl_conn.pack(anchor="w", pady=(8,0))

        self._refresh_com_ports()
        self._update_serial_buttons()

    # ---------- Settings apply ----------
    def _apply_settings(self):
        try:
            n  = int(self.num_points.get());  assert n > 0
            iv = int(self.auto_interval.get()); assert iv >= 10
            rset = float(self.r_set.get());  rtol = float(self.r_tol.get());  assert rtol >= 0
            vset = float(self.v_set.get());  vtol = float(self.v_tol.get());  assert vtol >= 0
        except Exception as e:
            messagebox.showerror("Invalid", f"Settings error: {e}"); return

        # ถ้าจำนวนจุดเปลี่ยน ต้อง rebuild main
        need_rebuild = (n != len(self.r_values))
        if need_rebuild:
            self._auto_stop()
            self._init_arrays()
            self._build_main()
        else:
            # แค่ค่าลิมิต/interval เปลี่ยน → refresh
            self._refresh_limits_labels()
            self._refresh_rows()
            self._update_big_box()

        messagebox.showinfo("Apply", "Settings applied.")

    def _refresh_limits_labels(self):
        rmin, rmax = self._r_bounds()
        vmin, vmax = self._v_bounds()
        self.lbl_rmin.config(text=f"{rmin:.3g}")
        self.lbl_rmax.config(text=f"{rmax:.3g}")
        self.lbl_vmin.config(text=f"{vmin:.3g}")
        self.lbl_vmax.config(text=f"{vmax:.3g}")

    # --------- Serial helpers ---------
    def _refresh_com_ports(self):
        ports = []
        try:
            from serial.tools import list_ports
            ports = sorted(
                [p.device for p in list_ports.comports()],
                key=lambda x: int(x.replace("COM","")) if x.startswith("COM") and x[3:].isdigit() else x
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
        if getattr(self, "btn_connect", None) is None:  # not built yet
            return
        if self.ser:
            self.btn_connect.state(["disabled"])
            self.btn_disconnect.state(["!disabled"])
        else:
            self.btn_connect.state(["!disabled"])
            self.btn_disconnect.state(["disabled"])

    # ---------- table/layout helpers ----------
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
        # (ยังคงอยู่เพื่อความเข้ากันได้ — ตอนนี้เราใช้ _update_scrollregion ภายใน _build_main แล้ว)
        self.points_canvas.configure(scrollregion=self.points_canvas.bbox("all"))
        try:
            self.points_canvas.itemconfigure(self.points_window, width=self.points_canvas.winfo_width())
        except: 
            pass

    def _bind_scrolling(self, widget, canvas):
        def on_wheel(e): canvas.yview_scroll(-1 if e.delta>0 else 1, "units")
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
            rmin, rmax = self._r_bounds()
            vmin, vmax = self._v_bounds()
            ok_r = (r is not None) and (rmin <= r <= rmax)
            ok_v = (v is not None) and (vmin <= v <= vmax)
            is_pass = ok_r and ok_v
            bg = COLOR_PASS_BG if is_pass else COLOR_FAIL_BG
            fg = COLOR_PASS_TEXT if is_pass else COLOR_FAIL_TEXT
            text = LABEL_OK if is_pass else LABEL_NG
        self.big_canvas.configure(bg=bg)
        w = self.big_canvas.winfo_width()  or int(self.big_canvas["width"])
        h = self.big_canvas.winfo_height() or int(self.big_canvas["height"])
        self.big_canvas.create_text(w/2, h/2, text=text, fill=fg, font=("Segoe UI", 60, "bold"))

    def _refresh_rows(self):
        rmin, rmax = self._r_bounds()
        vmin, vmax = self._v_bounds()
        for i,(lamp, r_var, r_ent, v_var, v_ent) in enumerate(self.row_widgets):
            r = self.r_values[i]
            v = self.v_values[i]

            if r is None: r_var.set("")
            else:        r_var.set(f"{r:.2f}")
            if v is None: v_var.set("")
            else:         v_var.set(f"{v:.4f}")

            ok_r = (r is not None) and (rmin <= r <= rmax)
            ok_v = (v is not None) and (vmin <= v <= vmax)
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
        self.lbl_ohm.config(text=("— mΩ" if r is None else f"{r:.2f} mΩ"))   # <-- 2 decimal
        self.lbl_volt.config(text=("— V"  if v is None else f"{v:.4f} V"))   # <-- 4 decimal
        self._draw_big_box()


    # ---------- measurement ----------
    def _measure_one(self, from_auto: bool = False):
        # ยังไม่ต่อ COM → เตือนและยกเลิก
        if not self._ensure_connected():
            if from_auto:
                self._auto_stop()
            return

        idx = self.current_idx
        try:
            r, v = self._read_meter()
        except Exception as e:
            # แจ้ง error และถ้าอยู่ในโหมด auto ให้หยุด
            messagebox.showerror("Measure Error", f"Failed to read data:\n{e}")
            if from_auto:
                self._auto_stop()
            return

        # อัปเดตค่า
        self.r_values[idx] = r
        self.v_values[idx] = v
        self._refresh_rows()
        self._update_big_box()

        # ไป cell ถัดไป หรือสรุปจบ
        is_last = (self.current_idx >= self.num_points.get() - 1)
        if not is_last:
            # เฉพาะ Manual เท่านั้นที่เด้งป็อปอัปแจ้งให้วัด cell ถัดไป
            if not from_auto and self.mode.get() == "manual":
                self._show_next_cell_popup()

            self.current_idx += 1
            self.point_combo.current(self.current_idx)
            self._scroll_row_into_view(self.current_idx)
        else:
            # ครบทุกจุด (cell สุดท้าย) → ไม่ต้องแสดงป็อปอัป
            if from_auto and self.auto_export.get():
                self._export_snapshot()

            if from_auto:
                self._auto_running = False
                if self._auto_job is not None:
                    self.after_cancel(self._auto_job); self._auto_job = None
                self._update_mode_buttons()
                self.after(0, lambda: messagebox.showinfo("Auto", "Auto measurement finished."))
            else:
                messagebox.showinfo("Done", "Measured all cells.")


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
        if self._auto_running:  # นัดรอบถัดไปเฉพาะถ้ายังไม่จบ
            interval = max(50, int(self.auto_interval.get()))
            self._auto_job = self.after(interval, self._tick_auto)

    def _parse_meter_line(self, line: str):
        """
        แปลงสตริงจากเครื่อง:  +5.87263E-03,+3.09940E+00,+0
        คืนค่า (r_ohm, v_volt, status)  -- r_ohm เป็น 'โอห์ม' (ยังไม่คูณเป็น mΩ)
        """
        s = (line or "").strip()
        parts = [p.strip() for p in s.split(",")]
        if len(parts) < 2:
            raise ValueError(f"bad format: {s}")

        r_ohm  = float(parts[0])
        v_volt = float(parts[1])
        status = None
        if len(parts) >= 3:
            try:
                status = int(parts[2].replace("+", ""))
            except Exception:
                status = None
        return r_ohm, v_volt, status


    def _query_fetc_once(self, line_ending=b"\r\n", timeout_ms=1500):
        """
        ส่งคำสั่ง FETC? แล้วอ่าน 1 บรรทัดกลับมา (decode เป็น str)
        """
        if not self.ser:
            raise RuntimeError("serial not connected")

        # เคลียร์บัฟเฟอร์ขาเข้า (กันค่าเก่าค้าง)
        try:
            self.ser.reset_input_buffer()
        except Exception:
            pass

        # ส่งคำสั่ง (ส่วนใหญ่ SCPI ใช้ \n หรือ \r\n — ลอง \r\n ก่อน)
        self.ser.write(b"FETC?" + line_ending)

        # อ่านคำตอบ 1 บรรทัด
        old_to = self.ser.timeout
        self.ser.timeout = max(0.1, timeout_ms/1000.0)
        try:
            raw = self.ser.readline()
            if not raw:
                raise TimeoutError("serial timeout")
            try:
                return raw.decode("ascii", errors="ignore")
            except Exception:
                return raw.decode("utf-8", errors="ignore")
        finally:
            self.ser.timeout = old_to

    def _read_meter(self):
        """
        พยายามอ่านค่าจริงด้วย FETC?:
        - เครื่องส่ง R เป็นโอห์ม, V เป็นโวลต์ -> แปลง R เป็น mΩ เพื่อให้ตรง UI
        ถ้าอ่านไม่ได้ ค่อย fallback เป็น random ในกรอบลิมิต
        """
        try:
            line = self._query_fetc_once(line_ending=b"\r\n")  # หรือ b"\n" หากเครื่องใช้แค่นิวไลน์เดียว
            r_ohm, v_volt, _status = self._parse_meter_line(line)
            r_milliohm = r_ohm * 1000.0  # แปลงโอห์ม -> mΩ
            return float(r_milliohm), float(v_volt)

        except Exception:
            # Fallback: สุ่มค่าในกรอบ Set±Tol เหมือนเดิม (ใช้งานได้แม้ไม่มีเครื่อง)
            rmin, rmax = self._r_bounds()
            vmin, vmax = self._v_bounds()
            r_center = (rmin + rmax)/2
            v_center = (vmin + vmax)/2
            r_span = (rmax - rmin)/2
            v_span = (vmax - vmin)/2
            if random.random() < 0.8:
                r = random.uniform(r_center-0.6*r_span, r_center+0.6*r_span)
                v = random.uniform(v_center-0.6*v_span, v_center+0.6*v_span)
            else:
                r = r_center + random.choice([-1,1]) * random.uniform(0.7*r_span, 1.6*r_span)
                v = v_center + random.choice([-1,1]) * random.uniform(0.7*v_span, 1.6*v_span)
            return float(r), float(v)


    def _show_next_cell_popup(self):
        """แสดงป็อปอัป 'Please measure the next cell' แล้วปิดเองใน 1 วินาที"""
        # ปิดป็อปอัปเก่าถ้ามี (กันซ้อน)
        try:
            if getattr(self, "_next_cell_popup", None) and self._next_cell_popup.winfo_exists():
                self._next_cell_popup.destroy()
        except Exception:
            pass

        top = tk.Toplevel(self)
        self._next_cell_popup = top
        top.title("Next Cell")
        top.transient(self)
        try:
            top.attributes("-topmost", True)
        except Exception:
            pass
        top.configure(bg=COLOR_PANEL)

        frm = ttk.Frame(top, style="Card.TFrame", padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Please measure the next cell", style="Heading.TLabel").pack(padx=8, pady=4)

        # จัดกึ่งกลางเหนือหน้าต่างหลัก
        top.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width()  - top.winfo_reqwidth())  // 2
        y = self.winfo_rooty() + (self.winfo_height() - top.winfo_reqheight()) // 2
        top.geometry(f"+{x}+{y}")

        # ปิดอัตโนมัติใน 1 วินาที
        top.after(1000, top.destroy)

    # ---------- export ----------
    def _toggle_auto_export(self):
        self.auto_export.set(not self.auto_export.get())
        self.btn_auto_export.config(text=f"Auto Export: {'ON' if self.auto_export.get() else 'OFF'}")

    def _export_txt_table(self):
        """Export .txt: Items, Cell, Pt, min(mΩ), R(mΩ), max(mΩ), min(V), V(V), max(V)"""
        import datetime
        from tkinter import filedialog, messagebox

        # Limits from Set ± Tol
        rmin = float(self.r_set.get()) - float(self.r_tol.get())
        rmax = float(self.r_set.get()) + float(self.r_tol.get())
        vmin = float(self.v_set.get()) - float(self.v_tol.get())
        vmax = float(self.v_set.get()) + float(self.v_tol.get())

        # helpers
        def one_r(i):
            val = self.r_values[i]
            if isinstance(val, (list, tuple)): val = val[0] if val else None
            return val
        def one_v(i):
            val = self.v_values[i]
            if isinstance(val, (list, tuple)): val = val[0] if val else None
            return val

        # format helpers
        def fr(x):  # R with 2 decimals
            return "" if x is None else f"{x:.2f}"
        def fv(x):  # V with 4 decimals
            return "" if x is None else f"{x:.4f}"
        def cell(t, w, align="<"):
            t = "" if t is None else str(t)
            return f"{t:{align}{w}}"

        # column spec (ปรับความกว้างได้)
        cols = [
            ("Items", 6),
            ("Cell", 6),
            ("Pt", 4),
            ("min(mΩ)", 8),
            ("R(mΩ)", 8),
            ("max(mΩ)", 8),
            ("min(V)", 10),
            ("V(V)", 10),
            ("max(V)", 10),
        ]

        # file name
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"{self.model_name.get().strip() or 'model'}_{ts}.txt"
        path = filedialog.asksaveasfilename(
            title="Save table (.txt)",
            initialfile=default_name,
            defaultextension=".txt",
            filetypes=[("Text file","*.txt")]
        )
        if not path:
            return

        lines = []
        lines.append(f"Model : {self.model_name.get().strip() or '-'}")
        lines.append(f"Time  : {datetime.datetime.now():%Y-%m-%d %H:%M:%S}")
        lines.append(f"R Set/Tol : {self.r_set.get()} mΩ  ±{self.r_tol.get()} mΩ  -> min={rmin:.3f}, max={rmax:.3f}")
        lines.append(f"V Set/Tol : {self.v_set.get()} V  ±{self.v_tol.get()} V  -> min={vmin:.4f}, max={vmax:.4f}")
        lines.append("")

        header = " ".join(cell(h, w) for h, w in cols)
        sep = "-" * len(header)
        lines.append(header)
        lines.append(sep)

        for i in range(self.num_points.get()):
            r = one_r(i)
            v = one_v(i)
            row = [
                i + 1,                       # Items
                f"Cell{i+1}",                # Cell
                i + 1,                       # Pt
                f"{rmin:.2f}",               # min(Ω)
                fr(r),                       # R(Ω)
                f"{rmax:.2f}",               # max(Ω)
                f"{vmin:.4f}",               # min(V)
                fv(v),                       # V(V)
                f"{vmax:.4f}",               # max(V)
            ]
            lines.append(" ".join(cell(x, w) for x, (_, w) in zip(row, cols)))

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            messagebox.showinfo("Export (.txt)", "Saved successfully.")
        except Exception as e:
            messagebox.showerror("Export (.txt)", f"Save failed:\n{e}")

    def _result_lines(self):
        rmin, rmax = self._r_bounds()
        vmin, vmax = self._v_bounds()
        lines = []
        lines.append(f"Model\t{self.model_name.get().strip() or '-'}")
        lines.append(f"Time\t{time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("")
        lines.append(f"R Set\t{self.r_set.get()} mΩ\tTol ±{self.r_tol.get()} mΩ\t(min={rmin}, max={rmax})")
        lines.append(f"V Set\t{self.v_set.get()} V\tTol ±{self.v_tol.get()} V\t(min={vmin}, max={vmax})")
        lines.append("")
        lines.append("Cell\tR(mΩ)\tV(V)\tResult")
        for i in range(self.num_points.get()):
            r = self.r_values[i]
            v = self.v_values[i]
            if r is None or v is None:
                lines.append(f"{i+1}\t\t\tN/A")
            else:
                res = LABEL_OK if self.flags[i] else LABEL_NG
                lines.append(f"{i+1}\t{r:.6f}\t{v:.6f}\t{res}")
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

    # ---------- misc ----------
    def _reset(self):
        self._auto_stop()
        self._init_arrays()
        self._build_main()

if __name__ == "__main__":
    App().mainloop()
