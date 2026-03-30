#!/usr/bin/env python3
"""
Power Stabiliser GUI
- Controls the Arduino-based power stabiliser (enable/disable, set point capture, calibration)
- Connects to a PM16 power meter via orca_drivers
- Plots live power readings and attenuator calibration curve
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import re
from collections import deque

import serial
import serial.tools.list_ports

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

try:
    from orca_drivers.drivers.power_meter.pm16 import PM16
    ORCA_DRIVERS_AVAILABLE = True
except ImportError:
    ORCA_DRIVERS_AVAILABLE = False


# ─── Backend classes ──────────────────────────────────────────────────────────

class StabiliserSerial:
    """
    Manages serial communication with the Arduino power stabiliser.

    Normal output:   Photodiode: 1234 | Reference: 1300 | Attenuator: 2048
    Calibration:     send 'cal\\n'  → Arduino streams  DAC,ADC  CSV lines,
                     ending with a line containing '# Calibration complete'.
    Set point:       send 'set\\n'  → Arduino captures current reading as reference.
    """

    _STATUS_PATTERN = re.compile(
        r"Photodiode:\s*(\d+)\s*\|\s*Reference:\s*(\d+)\s*\|\s*Attenuator:\s*(\d+)"
    )
    _CAL_PATTERN = re.compile(r"^(\d+),(\d+)$")
    _CAL_DONE_MARKER = "# Calibration complete"

    def __init__(self):
        self.ser: serial.Serial | None = None
        self.connected = False
        self._lock = threading.Lock()
        self._data: dict = {"photodiode": None, "reference": None, "attenuator": None}
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

        # Calibration state
        self._calibrating = False
        self._cal_data: list[tuple[int, int]] = []
        self._cal_done = threading.Event()

        # Raw terminal buffer (last 2000 lines)
        self._raw_lines: deque[str] = deque(maxlen=2000)

    # ── Connection ────────────────────────────────────────────────────────────

    def connect(self, port: str, baud: int = 9600) -> tuple[bool, str]:
        try:
            self.ser = serial.Serial(
                port=port,
                baudrate=baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=2,
            )
            time.sleep(0.1)
            self.connected = True
            self._stop.clear()
            self._thread = threading.Thread(target=self._read_loop, daemon=True)
            self._thread.start()
            return True, "Connected"
        except Exception as exc:
            self.connected = False
            return False, str(exc)

    def disconnect(self):
        self._stop.set()
        self.connected = False
        if self.ser and self.ser.is_open:
            self.ser.close()
        self.ser = None
        with self._lock:
            self._data = {"photodiode": None, "reference": None, "attenuator": None}

    # ── Commands ──────────────────────────────────────────────────────────────

    def send_get(self):
        """Request a single status snapshot from the Arduino."""
        if self.connected and self.ser and self.ser.is_open:
            self.ser.write(b"get\n")

    def send_set(self):
        """Capture current photodiode reading as the reference set point."""
        if self.connected and self.ser and self.ser.is_open:
            self.ser.write(b"set\n")

    def send_setref(self, value: int):
        """Set referenceValue to an explicit photodiode count."""
        if self.connected and self.ser and self.ser.is_open:
            self.ser.write(f"setref {int(value)}\n".encode())

    def send_cal(self):
        """Start a calibration sweep. Poll get_cal_result() for completion."""
        if not (self.connected and self.ser and self.ser.is_open):
            return
        with self._lock:
            self._cal_data = []
            self._calibrating = True
        self._cal_done.clear()
        self.ser.write(b"cal\n")

    # ── Data accessors ────────────────────────────────────────────────────────

    def get_data(self) -> dict:
        with self._lock:
            return dict(self._data)

    def get_cal_result(self) -> tuple[bool, list[tuple[int, int]]]:
        """Return (done, [(dac, adc), ...]).  done=True once sweep is complete."""
        done = self._cal_done.is_set()
        with self._lock:
            data = list(self._cal_data)
        return done, data

    def drain_raw_lines(self) -> list[str]:
        """Return and clear all buffered raw lines since the last call."""
        with self._lock:
            lines = list(self._raw_lines)
            self._raw_lines.clear()
        return lines

    # ── Serial read loop ──────────────────────────────────────────────────────

    def _read_loop(self):
        while not self._stop.is_set():
            try:
                if self.ser and self.ser.in_waiting > 0:
                    line = self.ser.readline().decode("ascii", errors="replace").strip()
                    with self._lock:
                        self._raw_lines.append(line)
                        in_cal = self._calibrating

                    if in_cal:
                        m = self._CAL_PATTERN.match(line)
                        if m:
                            with self._lock:
                                self._cal_data.append((int(m.group(1)), int(m.group(2))))
                        elif self._CAL_DONE_MARKER in line:
                            with self._lock:
                                self._calibrating = False
                            self._cal_done.set()
                    else:
                        m = self._STATUS_PATTERN.search(line)
                        if m:
                            with self._lock:
                                self._data = {
                                    "photodiode": int(m.group(1)),
                                    "reference":  int(m.group(2)),
                                    "attenuator": int(m.group(3)),
                                }
            except Exception:
                pass
            time.sleep(0.05)


class PM16Controller:
    """Wraps the orca_drivers PM16 driver for use from the GUI thread."""

    def __init__(self):
        self.pm: PM16 | None = None
        self.connected = False
        self.enabled = False
        self._power = 0.0
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def connect(self, address: str) -> tuple[bool, str]:
        if not ORCA_DRIVERS_AVAILABLE:
            return False, "orca_drivers package not found on PYTHONPATH"
        try:
            self.pm = PM16(configs={"address": address})
            result = self.pm.connect()
            if any(b in result for b in (b"Connected", b"Already connected")):
                self.connected = True
                return True, "Connected"
            return False, str(result)
        except Exception as exc:
            return False, str(exc)

    def disconnect(self):
        self._stop_polling()
        if self.pm and self.connected:
            try:
                self.pm.disconnect()
            except Exception:
                pass
        self.connected = False
        self.enabled = False
        self.pm = None

    def enable(self) -> tuple[bool, str]:
        if not self.connected or self.pm is None:
            return False, "Not connected"
        try:
            self.pm.enable()
            self.enabled = True
            self._stop.clear()
            self._thread = threading.Thread(target=self._poll_loop, daemon=True)
            self._thread.start()
            return True, "Enabled"
        except Exception as exc:
            return False, str(exc)

    def disable(self) -> tuple[bool, str]:
        self._stop_polling()
        if self.pm and self.connected:
            try:
                self.pm.disable()
            except Exception:
                pass
        self.enabled = False
        return True, "Disabled"

    def get_power(self) -> float:
        with self._lock:
            return self._power

    def _stop_polling(self):
        self._stop.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)

    def _poll_loop(self):
        while not self._stop.is_set():
            if self.pm:
                try:
                    p = self.pm.power
                    with self._lock:
                        self._power = p
                except Exception:
                    pass
            time.sleep(0.2)


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _histogram(values: list[float], n_bins: int) -> tuple[list[int], list[float]]:
    """Minimal histogram returning (counts, edges) without numpy."""
    lo, hi = min(values), max(values)
    if lo == hi:
        return [len(values)], [lo, hi + 1e-9]
    width = (hi - lo) / n_bins
    edges = [lo + i * width for i in range(n_bins + 1)]
    counts = [0] * n_bins
    for v in values:
        idx = min(int((v - lo) / width), n_bins - 1)
        counts[idx] += 1
    return counts, edges


# ─── GUI ─────────────────────────────────────────────────────────────────────

MAX_PLOT_POINTS = 500
UPDATE_INTERVAL_MS = 200


class StabiliserGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Power Stabiliser Control")
        self.root.minsize(950, 520)

        self.stabiliser = StabiliserSerial()
        self.pm = PM16Controller()

        self._power_times: deque[float] = deque(maxlen=MAX_PLOT_POINTS)
        self._power_values: deque[float] = deque(maxlen=MAX_PLOT_POINTS)
        self._t0: float | None = None
        self._calibrating_ui = False   # tracks whether cal is in progress in the GUI

        self._build_ui()
        self._schedule_update()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)

        controls = ttk.Frame(outer)
        controls.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))

        self._build_stabiliser_panel(controls)
        self._build_pm_panel(controls)
        self._build_plots(outer)

        self._refresh_ports()

    def _build_stabiliser_panel(self, parent):
        frame = ttk.LabelFrame(parent, text=" Stabiliser (Arduino) ", padding=8)
        frame.pack(fill=tk.X, pady=(0, 10))

        # Port row
        row = ttk.Frame(frame)
        row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(row, text="Port:").pack(side=tk.LEFT)
        self._port_var = tk.StringVar()
        self._port_combo = ttk.Combobox(row, textvariable=self._port_var, width=10)
        self._port_combo.pack(side=tk.LEFT, padx=4)
        ttk.Button(row, text="↺", width=3, command=self._refresh_ports).pack(side=tk.LEFT)

        # Connect / status
        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=(0, 4))
        self._stab_conn_btn = ttk.Button(row2, text="Connect", width=12,
                                         command=self._toggle_stab_connect)
        self._stab_conn_btn.pack(side=tk.LEFT)
        self._stab_conn_lbl = tk.Label(row2, text="● Disconnected",
                                       fg="red", font=("TkDefaultFont", 9))
        self._stab_conn_lbl.pack(side=tk.LEFT, padx=6)

        # Get (single snapshot)
        self._stab_get_btn = ttk.Button(frame, text="Get",
                                        command=self._send_get,
                                        state=tk.DISABLED)
        self._stab_get_btn.pack(fill=tk.X, pady=(0, 4))

        # Enable / disable
        self._stab_en_btn = ttk.Button(frame, text="Enable Stabiliser",
                                       command=self._toggle_stab_enable,
                                       state=tk.DISABLED)
        self._stab_en_btn.pack(fill=tk.X, pady=(0, 4))
        self._stab_enabled = False

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=4)

        # Readings
        grid = ttk.Frame(frame)
        grid.pack(fill=tk.X)
        labels = ["Photodiode:", "Reference (setpoint):", "Attenuator:", "Error:"]
        self._stab_vars = {}
        for i, (lbl, key) in enumerate(zip(labels,
                                           ["photodiode", "reference", "attenuator", "error"])):
            ttk.Label(grid, text=lbl, anchor=tk.W).grid(row=i, column=0, sticky=tk.W, pady=1)
            var = tk.StringVar(value="—")
            ttk.Label(grid, textvariable=var, width=10, anchor=tk.E).grid(
                row=i, column=1, sticky=tk.E, padx=(6, 0))
            self._stab_vars[key] = var

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # Set point — capture current or enter explicit value
        ttk.Label(frame, text="Set point:").pack(anchor=tk.W)

        sp_row = ttk.Frame(frame)
        sp_row.pack(fill=tk.X, pady=(0, 2))
        self._stab_set_btn = ttk.Button(sp_row, text="Capture current →",
                                        command=self._capture_setpoint,
                                        state=tk.DISABLED)
        self._stab_set_btn.pack(side=tk.LEFT)

        sp_val_row = ttk.Frame(frame)
        sp_val_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(sp_val_row, text="Value:").pack(side=tk.LEFT)
        self._setref_var = tk.IntVar(value=0)
        ttk.Spinbox(sp_val_row, from_=0, to=16383, increment=1, width=7,
                    textvariable=self._setref_var).pack(side=tk.LEFT, padx=4)
        self._stab_setref_btn = ttk.Button(sp_val_row, text="Set →",
                                           command=self._send_setref,
                                           state=tk.DISABLED)
        self._stab_setref_btn.pack(side=tk.LEFT)

        self._cal_btn = ttk.Button(frame, text="Calibrate",
                                   command=self._run_calibration,
                                   state=tk.DISABLED)
        self._cal_btn.pack(fill=tk.X)

        self._cal_status_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self._cal_status_var,
                  foreground="gray", font=("TkDefaultFont", 8)).pack(anchor=tk.W)

    def _build_pm_panel(self, parent):
        frame = ttk.LabelFrame(parent, text=" Power Meter (PM16) ", padding=8)
        frame.pack(fill=tk.X)

        if not ORCA_DRIVERS_AVAILABLE:
            ttk.Label(frame, text="orca_drivers not found on PYTHONPATH",
                      foreground="gray").pack()
            return

        visa_lbl_row = ttk.Frame(frame)
        visa_lbl_row.pack(fill=tk.X)
        ttk.Label(visa_lbl_row, text="VISA device:").pack(side=tk.LEFT)
        ttk.Button(visa_lbl_row, text="↺", width=3,
                   command=self._refresh_visa_resources).pack(side=tk.RIGHT)
        self._visa_var = tk.StringVar()
        self._visa_combo = ttk.Combobox(frame, textvariable=self._visa_var, width=38)
        self._visa_combo.pack(fill=tk.X, pady=(2, 6))
        self._refresh_visa_resources()

        row = ttk.Frame(frame)
        row.pack(fill=tk.X, pady=(0, 4))
        self._pm_conn_btn = ttk.Button(row, text="Connect", width=12,
                                       command=self._toggle_pm_connect)
        self._pm_conn_btn.pack(side=tk.LEFT)
        self._pm_conn_lbl = tk.Label(row, text="● Disconnected",
                                     fg="red", font=("TkDefaultFont", 9))
        self._pm_conn_lbl.pack(side=tk.LEFT, padx=6)

        self._pm_en_btn = ttk.Button(frame, text="Enable Measurements",
                                     command=self._toggle_pm_enable,
                                     state=tk.DISABLED)
        self._pm_en_btn.pack(fill=tk.X, pady=(0, 6))
        self._pm_enabled = False

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=4)

        pw_row = ttk.Frame(frame)
        pw_row.pack(fill=tk.X)
        ttk.Label(pw_row, text="Power:").pack(side=tk.LEFT)
        self._power_var = tk.StringVar(value="— mW")
        ttk.Label(pw_row, textvariable=self._power_var,
                  font=("TkDefaultFont", 11, "bold")).pack(side=tk.LEFT, padx=6)

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=4)

        # Statistics over last N seconds
        stats_grid = ttk.Frame(frame)
        stats_grid.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(stats_grid, text="Stats window:").grid(row=0, column=0, sticky=tk.W)
        self._stats_window_var = tk.DoubleVar(value=10.0)
        ttk.Spinbox(stats_grid, from_=1, to=3600, increment=1, width=6,
                    textvariable=self._stats_window_var).grid(row=0, column=1, padx=4)
        ttk.Label(stats_grid, text="s").grid(row=0, column=2, sticky=tk.W)

        ttk.Label(stats_grid, text="Mean:").grid(row=1, column=0, sticky=tk.W, pady=1)
        self._mean_var = tk.StringVar(value="— mW")
        ttk.Label(stats_grid, textvariable=self._mean_var).grid(row=1, column=1,
                                                                 columnspan=2, sticky=tk.W)
        ttk.Label(stats_grid, text="Std dev:").grid(row=2, column=0, sticky=tk.W, pady=1)
        self._std_var = tk.StringVar(value="— mW")
        ttk.Label(stats_grid, textvariable=self._std_var).grid(row=2, column=1,
                                                                columnspan=2, sticky=tk.W)

        ttk.Button(frame, text="Clear plot", command=self._clear_power_plot).pack(
            fill=tk.X, pady=(4, 0))

    def _build_plots(self, parent):
        """Two-tab notebook: Power vs Time  |  Calibration."""
        self._notebook = ttk.Notebook(parent)
        self._notebook.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ── Tab 1: Power ──────────────────────────────────────────────────────
        power_tab = ttk.Frame(self._notebook)
        self._notebook.add(power_tab, text="  Power vs Time  ")

        self._fig_power = Figure(figsize=(7, 4), dpi=100)
        gs = self._fig_power.add_gridspec(1, 2, width_ratios=[4, 1], wspace=0.05)
        self._ax_power = self._fig_power.add_subplot(gs[0])
        self._ax_hist  = self._fig_power.add_subplot(gs[1], sharey=self._ax_power)

        self._ax_power.set_xlabel("Time (s)")
        self._ax_power.set_ylabel("Power (mW)")
        self._ax_power.set_title("PM16 Live Power")
        (self._power_line,) = self._ax_power.plot([], [], color="#00aacc", linewidth=1)

        self._ax_hist.set_xlabel("Counts\n(norm.)")
        self._ax_hist.tick_params(labelleft=False)
        self._ax_hist.set_title("Hist.")
        self._fig_power.tight_layout()

        canvas_power = FigureCanvasTkAgg(self._fig_power, master=power_tab)
        canvas_power.draw()
        canvas_power.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self._canvas_power = canvas_power

        # ── Tab 3: Arduino terminal ───────────────────────────────────────────
        term_tab = ttk.Frame(self._notebook)
        self._notebook.add(term_tab, text="  Arduino Terminal  ")

        self._term_text = tk.Text(
            term_tab, state=tk.DISABLED, wrap=tk.NONE,
            bg="#1e1e1e", fg="#d4d4d4",
            font=("Courier New", 9),
            relief=tk.FLAT,
        )
        term_scroll_y = ttk.Scrollbar(term_tab, orient=tk.VERTICAL,
                                      command=self._term_text.yview)
        term_scroll_x = ttk.Scrollbar(term_tab, orient=tk.HORIZONTAL,
                                      command=self._term_text.xview)
        self._term_text.configure(yscrollcommand=term_scroll_y.set,
                                  xscrollcommand=term_scroll_x.set)
        term_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        term_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self._term_text.pack(fill=tk.BOTH, expand=True)

        ttk.Button(term_tab, text="Clear",
                   command=self._clear_terminal).place(relx=1.0, rely=0.0,
                                                       anchor="ne", x=-20, y=2)

        # ── Tab 2: Calibration ────────────────────────────────────────────────
        cal_tab = ttk.Frame(self._notebook)
        self._notebook.add(cal_tab, text="  Calibration  ")

        self._fig_cal = Figure(figsize=(6, 4), dpi=100)
        self._ax_cal = self._fig_cal.add_subplot(111)
        self._ax_cal.set_xlabel("Attenuator (DAC)")
        self._ax_cal.set_ylabel("Photodiode (ADC)")
        self._ax_cal.set_title("Attenuator Calibration")
        self._ax_cal.text(0.5, 0.5, "Run calibration to populate this plot",
                          ha="center", va="center", transform=self._ax_cal.transAxes,
                          color="gray")
        self._fig_cal.tight_layout()

        canvas_cal = FigureCanvasTkAgg(self._fig_cal, master=cal_tab)
        canvas_cal.draw()
        canvas_cal.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self._canvas_cal = canvas_cal

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self._port_combo["values"] = ports
        if ports and not self._port_var.get():
            self._port_var.set(ports[0])

    def _refresh_visa_resources(self):
        try:
            import pyvisa
            rm = pyvisa.ResourceManager()
            resources = list(rm.list_resources())
            rm.close()
        except Exception:
            resources = []
        self._visa_combo["values"] = resources
        if resources and not self._visa_var.get():
            self._visa_var.set(resources[0])

    def _set_stab_connected(self, connected: bool):
        if connected:
            self._stab_conn_btn.config(text="Disconnect")
            self._stab_conn_lbl.config(text="● Connected", fg="#008800")
            self._stab_get_btn.config(state=tk.NORMAL)
            self._stab_en_btn.config(state=tk.NORMAL)
            self._stab_setref_btn.config(state=tk.NORMAL)
        else:
            self._stab_conn_btn.config(text="Connect")
            self._stab_conn_lbl.config(text="● Disconnected", fg="red")
            self._stab_get_btn.config(state=tk.DISABLED)
            self._stab_en_btn.config(text="Enable Stabiliser", state=tk.DISABLED)
            self._stab_set_btn.config(state=tk.DISABLED)
            self._stab_setref_btn.config(state=tk.DISABLED)
            self._cal_btn.config(state=tk.DISABLED)
            self._cal_status_var.set("")
            self._stab_enabled = False
            for var in self._stab_vars.values():
                var.set("—")

    def _set_pm_connected(self, connected: bool):
        if connected:
            self._pm_conn_btn.config(text="Disconnect")
            self._pm_conn_lbl.config(text="● Connected", fg="#008800")
            self._pm_en_btn.config(state=tk.NORMAL)
        else:
            self._pm_conn_btn.config(text="Connect")
            self._pm_conn_lbl.config(text="● Disconnected", fg="red")
            self._pm_en_btn.config(text="Enable Measurements", state=tk.DISABLED)
            self._pm_enabled = False
            self._power_var.set("— mW")

    # ── Stabiliser actions ────────────────────────────────────────────────────

    def _toggle_stab_connect(self):
        if not self.stabiliser.connected:
            port = self._port_var.get()
            if not port:
                messagebox.showerror("No port", "Select a COM port first.")
                return
            ok, msg = self.stabiliser.connect(port)
            if ok:
                self._set_stab_connected(True)
            else:
                messagebox.showerror("Connection failed", msg)
        else:
            if self._stab_enabled:
                self._toggle_stab_enable()
            self.stabiliser.disconnect()
            self._set_stab_connected(False)

    def _toggle_stab_enable(self):
        self._stab_enabled = not self._stab_enabled
        if self._stab_enabled:
            self._stab_en_btn.config(text="Disable Stabiliser")
            self._stab_set_btn.config(state=tk.NORMAL)
            self._cal_btn.config(state=tk.NORMAL)
        else:
            self._stab_en_btn.config(text="Enable Stabiliser")
            self._stab_set_btn.config(state=tk.DISABLED)
            self._cal_btn.config(state=tk.DISABLED)

    def _send_get(self):
        self.stabiliser.send_get()

    def _send_setref(self):
        try:
            value = int(self._setref_var.get())
        except (tk.TclError, ValueError):
            messagebox.showerror("Invalid value", "Enter an integer photodiode count.")
            return
        self.stabiliser.send_setref(value)

    def _capture_setpoint(self):
        if not self.stabiliser.connected:
            messagebox.showwarning("Not connected", "Connect to stabiliser first.")
            return
        self.stabiliser.send_set()

    def _run_calibration(self):
        if not self.stabiliser.connected:
            messagebox.showwarning("Not connected", "Connect to stabiliser first.")
            return
        if self._calibrating_ui:
            return
        self._calibrating_ui = True
        self._cal_btn.config(state=tk.DISABLED, text="Calibrating…")
        self._stab_get_btn.config(state=tk.DISABLED)
        self._cal_status_var.set("Sweep in progress…")
        self.stabiliser.send_cal()

    # ── Power meter actions ───────────────────────────────────────────────────

    def _toggle_pm_connect(self):
        if not self.pm.connected:
            addr = self._visa_var.get().strip()
            if not addr:
                messagebox.showerror("No address", "Enter a VISA address.")
                return
            ok, msg = self.pm.connect(addr)
            if ok:
                self._set_pm_connected(True)
            else:
                messagebox.showerror("Connection failed", f"PM16: {msg}")
        else:
            if self._pm_enabled:
                self._toggle_pm_enable()
            self.pm.disconnect()
            self._set_pm_connected(False)

    def _toggle_pm_enable(self):
        if not self._pm_enabled:
            ok, msg = self.pm.enable()
            if ok:
                self._pm_enabled = True
                self._pm_en_btn.config(text="Disable Measurements")
                self._t0 = time.time()
                self._power_times.clear()
                self._power_values.clear()
            else:
                messagebox.showerror("Enable failed", f"PM16: {msg}")
        else:
            self.pm.disable()
            self._pm_enabled = False
            self._pm_en_btn.config(text="Enable Measurements")

    def _clear_terminal(self):
        self._term_text.config(state=tk.NORMAL)
        self._term_text.delete("1.0", tk.END)
        self._term_text.config(state=tk.DISABLED)

    def _clear_power_plot(self):
        self._power_times.clear()
        self._power_values.clear()
        if self._t0 is not None:
            self._t0 = time.time()
        self._power_line.set_data([], [])
        if hasattr(self, "_stats_hlines"):
            for artist in self._stats_hlines:
                artist.remove()
            self._stats_hlines = []
        self._ax_power.legend().remove() if self._ax_power.get_legend() else None
        self._mean_var.set("— mW")
        self._std_var.set("— mW")
        self._ax_power.relim()
        self._ax_hist.cla()
        self._ax_hist.tick_params(labelleft=False)
        self._ax_hist.set_xlabel("Counts\n(norm.)")
        self._ax_hist.set_title("Hist.")
        self._canvas_power.draw_idle()

    # ── Calibration plot ──────────────────────────────────────────────────────

    def _draw_calibration_plot(self, data: list[tuple[int, int]]):
        if not data:
            return

        dac_vals = [d for d, _ in data]
        adc_vals = [a for _, a in data]
        max_adc = max(adc_vals)
        half_max = max_adc / 2.0

        # Find the DAC value where photodiode first drops to half-max
        half_dac = None
        for dac, adc in data:
            if adc <= half_max:
                half_dac = dac
                break

        ax = self._ax_cal
        ax.clear()
        ax.plot(dac_vals, adc_vals, "b-o", linewidth=1.5, markersize=4, label="PD response")
        ax.axhline(half_max, color="orange", linestyle="--", linewidth=1,
                   label=f"Half-max ({half_max:.0f} ADC)")
        if half_dac is not None:
            ax.axvline(half_dac, color="green", linestyle="--", linewidth=1,
                       label=f"50 % point (DAC={half_dac})")

        ax.set_xlabel("Attenuator (DAC)")
        ax.set_ylabel("Photodiode (ADC)")
        ax.set_title("Attenuator Calibration")
        ax.legend(fontsize=8)
        self._fig_cal.tight_layout()
        self._canvas_cal.draw_idle()

    # ── Periodic update ───────────────────────────────────────────────────────

    def _schedule_update(self):
        self._update()
        self.root.after(UPDATE_INTERVAL_MS, self._schedule_update)

    def _update(self):
        # Stabiliser readings — auto-poll 'get' since the Arduino no longer
        # streams continuously; the response is parsed by the existing read loop.
        if self.stabiliser.connected and self._stab_enabled and not self._calibrating_ui:
            self.stabiliser.send_get()
            d = self.stabiliser.get_data()
            pd  = d["photodiode"]
            ref = d["reference"]
            att = d["attenuator"]
            if pd is not None:
                self._stab_vars["photodiode"].set(str(pd))
                self._stab_vars["reference"].set(str(ref))
                self._stab_vars["attenuator"].set(str(att))
                self._stab_vars["error"].set(str(ref - pd) if ref is not None else "—")

        # Calibration completion check
        if self._calibrating_ui:
            done, data = self.stabiliser.get_cal_result()
            if done:
                self._calibrating_ui = False
                self._cal_btn.config(state=tk.NORMAL, text="Calibrate")
                self._stab_get_btn.config(state=tk.NORMAL)
                self._cal_status_var.set(f"Done — {len(data)} points")
                self._draw_calibration_plot(data)
                self._notebook.select(2)   # switch to Calibration tab

        # Arduino terminal
        if self.stabiliser.connected:
            new_lines = self.stabiliser.drain_raw_lines()
            if new_lines:
                self._term_text.config(state=tk.NORMAL)
                for line in new_lines:
                    self._term_text.insert(tk.END, line + "\n")
                self._term_text.see(tk.END)
                self._term_text.config(state=tk.DISABLED)

        # Power meter
        if self._pm_enabled:
            p_w = self.pm.get_power()
            p_mw = p_w * 1e3
            self._power_var.set(f"{p_mw:.4f} mW")
            t = time.time() - self._t0
            self._power_times.append(t)
            self._power_values.append(p_mw)
            self._update_power_plot()

    def _update_power_plot(self):
        if not self._power_times:
            return

        times = list(self._power_times)
        values = list(self._power_values)
        self._power_line.set_data(times, values)
        self._ax_power.relim()
        self._ax_power.autoscale_view()

        # ── Statistics over the last N seconds ───────────────────────────────
        try:
            window = float(self._stats_window_var.get())
        except (tk.TclError, ValueError):
            window = 10.0
        t_now = times[-1]
        window_vals = [v for t, v in zip(times, values) if t >= t_now - window]

        if window_vals:
            mean = sum(window_vals) / len(window_vals)
            variance = sum((v - mean) ** 2 for v in window_vals) / len(window_vals)
            std = variance ** 0.5
            self._mean_var.set(f"{mean:.4f} mW")
            self._std_var.set(f"{std:.4f} mW")

            # ── Mean ± std on time-series ─────────────────────────────────────
            if hasattr(self, "_stats_hlines"):
                for artist in self._stats_hlines:
                    artist.remove()
            mean_line = self._ax_power.axhline(
                mean, color="orange", linestyle="--", linewidth=1,
                label=f"Mean ({mean:.4f} mW)")
            band = self._ax_power.axhspan(
                mean - std, mean + std, alpha=0.12, color="orange",
                label=f"±1σ ({std:.4f} mW)")
            self._stats_hlines = [mean_line, band]
            self._ax_power.legend(fontsize=7, loc="upper left")

            # ── Histogram (horizontal, shares y-axis) ─────────────────────────
            self._ax_hist.cla()
            self._ax_hist.tick_params(labelleft=False)
            self._ax_hist.set_xlabel("Counts\n(norm.)")
            self._ax_hist.set_title("Hist.")

            n_bins = min(50, max(10, len(window_vals) // 5))
            counts, edges = _histogram(window_vals, n_bins)
            norm = max(counts) if max(counts) > 0 else 1.0
            counts_norm = [c / norm for c in counts]
            bin_centers = [(edges[i] + edges[i + 1]) / 2 for i in range(len(counts))]
            bin_width = edges[1] - edges[0]

            self._ax_hist.barh(bin_centers, counts_norm, height=bin_width * 0.9,
                               color="#00aacc", alpha=0.6)
            self._ax_hist.axhline(mean, color="orange", linestyle="--", linewidth=1)
            self._ax_hist.axhspan(mean - std, mean + std, alpha=0.12, color="orange")
            self._ax_hist.set_xlim(left=0)

        self._canvas_power.draw_idle()

    def cleanup(self):
        self.stabiliser.disconnect()
        self.pm.disconnect()


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    app = StabiliserGUI(root)
    root.protocol("WM_DELETE_WINDOW", lambda: _cleanup(app, root))
    root.mainloop()


def _cleanup(app: StabiliserGUI, root: tk.Tk):
    app.cleanup()
    root.destroy()


if __name__ == "__main__":
    main()
