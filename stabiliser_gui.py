#!/usr/bin/env python3
"""
Power Stabiliser GUI
- Controls the Arduino-based power stabiliser (enable/disable, set point capture)
- Connects to a PM16 power meter via orca_drivers
- Plots live power readings
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

    The Arduino sends continuous lines of the form:
        Photodiode: 1234 | Reference: 1300 | Attenuator: 2048
    Sending "set\\n" tells the Arduino to capture the current photodiode
    reading as the new reference set point.
    """

    _PATTERN = re.compile(
        r"Photodiode:\s*(\d+)\s*\|\s*Reference:\s*(\d+)\s*\|\s*Attenuator:\s*(\d+)"
    )

    def __init__(self):
        self.ser: serial.Serial | None = None
        self.connected = False
        self._lock = threading.Lock()
        self._data: dict = {"photodiode": None, "reference": None, "attenuator": None}
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

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

    def send_set(self):
        """Send the 'set' command to capture the current reading as set point."""
        if self.connected and self.ser and self.ser.is_open:
            with self._lock:
                self.ser.write(b"set\n")

    def get_data(self) -> dict:
        with self._lock:
            return dict(self._data)

    def _read_loop(self):
        while not self._stop.is_set():
            try:
                if self.ser and self.ser.in_waiting > 0:
                    line = self.ser.readline().decode("ascii", errors="replace").strip()
                    m = self._PATTERN.search(line)
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
    """Wraps the orca_drivers PM16 driver for use from a GUI thread."""

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


# ─── GUI ─────────────────────────────────────────────────────────────────────

MAX_PLOT_POINTS = 500
UPDATE_INTERVAL_MS = 200


class StabiliserGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Power Stabiliser Control")
        self.root.minsize(900, 480)

        self.stabiliser = StabiliserSerial()
        self.pm = PM16Controller()

        self._power_times: deque[float] = deque(maxlen=MAX_PLOT_POINTS)
        self._power_values: deque[float] = deque(maxlen=MAX_PLOT_POINTS)
        self._t0: float | None = None

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
        self._build_plot(outer)

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

        # Enable / disable
        self._stab_en_btn = ttk.Button(frame, text="Enable Stabiliser",
                                       command=self._toggle_stab_enable,
                                       state=tk.DISABLED)
        self._stab_en_btn.pack(fill=tk.X, pady=(0, 6))
        self._stab_enabled = False

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=4)

        # Readings grid
        grid = ttk.Frame(frame)
        grid.pack(fill=tk.X)

        labels = ["Photodiode:", "Reference (setpoint):", "Attenuator:", "Error:"]
        self._stab_vars = {}
        keys = ["photodiode", "reference", "attenuator", "error"]
        for i, (lbl, key) in enumerate(zip(labels, keys)):
            ttk.Label(grid, text=lbl, anchor=tk.W).grid(row=i, column=0, sticky=tk.W, pady=1)
            var = tk.StringVar(value="—")
            ttk.Label(grid, textvariable=var, width=10, anchor=tk.E).grid(
                row=i, column=1, sticky=tk.E, padx=(6, 0))
            self._stab_vars[key] = var

        ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        # Set point capture
        sp_row = ttk.Frame(frame)
        sp_row.pack(fill=tk.X)
        ttk.Label(sp_row, text="Set point:").pack(side=tk.LEFT)
        self._stab_set_btn = ttk.Button(sp_row, text="Capture current →",
                                        command=self._capture_setpoint,
                                        state=tk.DISABLED)
        self._stab_set_btn.pack(side=tk.LEFT, padx=6)

    def _build_pm_panel(self, parent):
        frame = ttk.LabelFrame(parent, text=" Power Meter (PM16) ", padding=8)
        frame.pack(fill=tk.X)

        if not ORCA_DRIVERS_AVAILABLE:
            ttk.Label(frame, text="orca_drivers not found on PYTHONPATH",
                      foreground="gray").pack()
            return

        ttk.Label(frame, text="VISA address:").pack(anchor=tk.W)
        self._visa_var = tk.StringVar(value="USB0::0x1313::0x807B::240426527::INSTR")
        ttk.Entry(frame, textvariable=self._visa_var, width=38).pack(fill=tk.X, pady=(2, 6))

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

        # Clear plot button
        ttk.Button(frame, text="Clear plot", command=self._clear_plot).pack(
            fill=tk.X, pady=(6, 0))

    def _build_plot(self, parent):
        plot_frame = ttk.LabelFrame(parent, text=" Power vs Time ", padding=4)
        plot_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._fig = Figure(figsize=(6, 4), dpi=100)
        self._ax = self._fig.add_subplot(111)
        self._ax.set_xlabel("Time (s)")
        self._ax.set_ylabel("Power (mW)")
        self._ax.set_title("PM16 Live Power")
        (self._line,) = self._ax.plot([], [], "b-", linewidth=1)
        self._fig.tight_layout()

        self._canvas = FigureCanvasTkAgg(self._fig, master=plot_frame)
        self._canvas.draw()
        self._canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self._port_combo["values"] = ports
        if ports and not self._port_var.get():
            self._port_var.set(ports[0])

    def _set_stab_connected(self, connected: bool):
        if connected:
            self._stab_conn_btn.config(text="Disconnect")
            self._stab_conn_lbl.config(text="● Connected", fg="#008800")
            self._stab_en_btn.config(state=tk.NORMAL)
        else:
            self._stab_conn_btn.config(text="Connect")
            self._stab_conn_lbl.config(text="● Disconnected", fg="red")
            self._stab_en_btn.config(text="Enable Stabiliser", state=tk.DISABLED)
            self._stab_set_btn.config(state=tk.DISABLED)
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
        else:
            self._stab_en_btn.config(text="Enable Stabiliser")
            self._stab_set_btn.config(state=tk.DISABLED)

    def _capture_setpoint(self):
        """Send 'set' to Arduino — captures current photodiode reading as reference."""
        if not self.stabiliser.connected:
            messagebox.showwarning("Not connected", "Connect to stabiliser first.")
            return
        self.stabiliser.send_set()

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

    def _clear_plot(self):
        self._power_times.clear()
        self._power_values.clear()
        if self._t0 is not None:
            self._t0 = time.time()
        self._line.set_data([], [])
        self._ax.relim()
        self._canvas.draw_idle()

    # ── Periodic update ───────────────────────────────────────────────────────

    def _schedule_update(self):
        self._update()
        self.root.after(UPDATE_INTERVAL_MS, self._schedule_update)

    def _update(self):
        # Stabiliser readings
        if self.stabiliser.connected and self._stab_enabled:
            d = self.stabiliser.get_data()
            pd  = d["photodiode"]
            ref = d["reference"]
            att = d["attenuator"]
            if pd is not None:
                self._stab_vars["photodiode"].set(str(pd))
                self._stab_vars["reference"].set(str(ref))
                self._stab_vars["attenuator"].set(str(att))
                self._stab_vars["error"].set(str(ref - pd) if ref is not None else "—")

        # Power meter
        if self._pm_enabled:
            p_w = self.pm.get_power()
            p_mw = p_w * 1e3
            self._power_var.set(f"{p_mw:.4f} mW")
            t = time.time() - self._t0
            self._power_times.append(t)
            self._power_values.append(p_mw)
            self._update_plot()

    def _update_plot(self):
        if not self._power_times:
            return
        xs = list(self._power_times)
        ys = list(self._power_values)
        self._line.set_data(xs, ys)
        self._ax.relim()
        self._ax.autoscale_view()
        self._canvas.draw_idle()

    def cleanup(self):
        self.stabiliser.disconnect()
        self.pm.disconnect()


# ─── Entry point ─────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    app = StabiliserGUI(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (_cleanup(app, root)))
    root.mainloop()


def _cleanup(app: StabiliserGUI, root: tk.Tk):
    app.cleanup()
    root.destroy()


if __name__ == "__main__":
    main()
