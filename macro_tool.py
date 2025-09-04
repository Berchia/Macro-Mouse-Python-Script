import threading
import time
import sys
import json
import random
import os
from pathlib import Path
from datetime import datetime
from tkinter import Tk, Toplevel, StringVar, IntVar, DoubleVar, BooleanVar, ttk, messagebox, filedialog

try:
    from pynput import keyboard, mouse
    from pynput.keyboard import Key, Controller as KeyController, KeyCode
    from pynput.mouse import Button, Controller as MouseController
except ImportError:
    print("The 'pynput' package is required. Install with: pip install pynput")
    sys.exit(1)

APP_NAME = "MacroTool"
SETTINGS_FILE = "settings.json"

# -----------------------------
# Helpers
# -----------------------------
def clamp(v, lo, hi):
    return max(lo, min(hi, v))

def key_to_str(k) -> str:
    if isinstance(k, KeyCode):
        return (k.char or "").lower()
    if isinstance(k, Key):
        return str(k).split(".")[-1].lower()
    return str(k).lower()

SPECIAL_KEY_MAP = {
    "enter": Key.enter, "return": Key.enter,
    "space": Key.space, "tab": Key.tab,
    "esc": Key.esc, "escape": Key.esc,
    "backspace": Key.backspace, "delete": Key.delete,
    "home": Key.home, "end": Key.end,
    "page_up": Key.page_up, "pageup": Key.page_up,
    "page_down": Key.page_down, "pagedown": Key.page_down,
    "up": Key.up, "down": Key.down, "left": Key.left, "right": Key.right,
    "shift": Key.shift, "ctrl": Key.ctrl, "alt": Key.alt, "cmd": Key.cmd, "win": Key.cmd,
    "caps_lock": Key.caps_lock,
    "f1": Key.f1, "f2": Key.f2, "f3": Key.f3, "f4": Key.f4, "f5": Key.f5, "f6": Key.f6,
    "f7": Key.f7, "f8": Key.f8, "f9": Key.f9, "f10": Key.f10, "f11": Key.f11, "f12": Key.f12
}

def str_to_key(s: str):
    s = (s or "").strip().lower()
    if s in SPECIAL_KEY_MAP:
        return SPECIAL_KEY_MAP[s]
    if len(s) == 1:
        return KeyCode.from_char(s)
    if s.startswith("f") and s[1:].isdigit():
        idx = int(s[1:])
        try:
            return getattr(Key, f"f{idx}")
        except AttributeError:
            pass
    return None

def get_config_dir() -> Path:
    if os.name == "nt":
        base = os.environ.get("APPDATA")
        if base:
            return Path(base) / APP_NAME
        return Path.home() / f"AppData/Roaming/{APP_NAME}"
    if sys.platform == "darwin":
        return Path.home() / "Library/Application Support" / APP_NAME
    return Path.home() / ".config" / APP_NAME

def get_settings_path() -> Path:
    cfg_dir = get_config_dir()
    cfg_dir.mkdir(parents=True, exist_ok=True)
    return cfg_dir / SETTINGS_FILE

# -----------------------------
# Main App
# -----------------------------
class MacroTool:
    def __init__(self, root: Tk):
        self.root = root

        #The taskbar/window title reflects Active/Idle
        self._active = False
        self._apply_active_title()

        self.root.attributes("-topmost", True)

        self.kctl = KeyController()
        self.mctl = MouseController()

        #State
        self.mode = StringVar(value="key")           # 'key' or 'mouse'
        self.spam_key = StringVar(value="r")
        self.status = StringVar(value="Status: IDLE")

        #Mouse click settings
        self.click_button = StringVar(value="left")  #left/right
        self.click_type = StringVar(value="single")  #single/double
        self.target_mode = StringVar(value="cursor") #cursor/fixed
        self.fixed_x = IntVar(value=0)
        self.fixed_y = IntVar(value=0)

        #Nudge settings
        self.nudge_mode = StringVar(value="off")     #'off' | 'on'
        self.nudge_x = IntVar(value=0)
        self.nudge_y = IntVar(value=0)
        self.nudge_random = BooleanVar(value=False)  #"Humanized Mouse Clicks"

        #Interval (four boxes, all active at once)
        self.int_hours = IntVar(value=0)
        self.int_minutes = IntVar(value=0)
        self.int_seconds = IntVar(value=0)
        self.int_millis = IntVar(value=50)

        #Global hotkeys (single keys)
        self.action_hotkey = StringVar(value="f7")
        self.record_hotkey = StringVar(value="f8")
        self.play_hotkey = StringVar(value="f9")

        #Run flags
        self.running_event = threading.Event()
        self.playback_stop = threading.Event()
        self.recording = False

        #Recording data
        self.record_events = []
        self.record_start_time = 0.0

        #Repeat count for playback (0 = infinite)
        self.repeat_count = IntVar(value=1)

        #Autosave debounce id
        self._save_after_id = None

        #For “capture next key” when clicking the Key box
        self.capturing_spam_key = False

        #UI
        self.build_ui()

        #Load persisted settings
        self.load_settings()

        #Hotkeys
        self.start_global_listeners()

        #Auto-save bindings
        self.attach_autosave_traces()
        self.root.bind("<Configure>", self._on_configure)

    #Active/Idle Title
    def _apply_active_title(self):
        self.root.title("Active" if self._active else "Idle")

    def _set_active(self, active: bool):
        self._active = active
        self._apply_active_title()

    #UI
    def build_ui(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        #Mode selector
        mode_box = ttk.LabelFrame(frm, text="Mode")
        mode_box.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        ttk.Radiobutton(mode_box, text="Key Masher", variable=self.mode, value="key").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(mode_box, text="Auto Clicker", variable=self.mode, value="mouse").grid(row=0, column=1, padx=5, pady=5)

        #Key masher settings
        key_box = ttk.LabelFrame(frm, text="Key Settings")
        key_box.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        ttk.Label(key_box, text="Key to spam:").grid(row=0, column=0, sticky="w", padx=5, pady=5)

        self.spam_key_entry = ttk.Entry(key_box, textvariable=self.spam_key, width=12)
        self.spam_key_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        #When the entry gets focus, capture next key pressed globally
        self.spam_key_entry.bind("<FocusIn>", lambda e: self._begin_capture_spam_key())
        #If user types manually, still autosave
        self.spam_key_entry.bind("<KeyRelease>", lambda e: self._schedule_save())

        ttk.Label(key_box, text="(click box, then press a key — e.g., Space, Enter, F7)").grid(row=0, column=2, sticky="w", padx=5)

        #Mouse clicker settings
        mouse_box = ttk.LabelFrame(frm, text="Mouse Settings")
        mouse_box.grid(row=2, column=0, sticky="ew", padx=5, pady=5)
        ttk.Label(mouse_box, text="Button:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Combobox(mouse_box, values=["left", "right"], textvariable=self.click_button, width=8, state="readonly").grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(mouse_box, text="Click:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        ttk.Combobox(mouse_box, values=["single", "double"], textvariable=self.click_type, width=8, state="readonly").grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(mouse_box, text="Target:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Radiobutton(mouse_box, text="Current Cursor", variable=self.target_mode, value="cursor").grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(mouse_box, text="Fixed Position", variable=self.target_mode, value="fixed").grid(row=1, column=2, padx=5, pady=5, sticky="w")

        ttk.Label(mouse_box, text="X:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        ttk.Entry(mouse_box, textvariable=self.fixed_x, width=8).grid(row=2, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(mouse_box, text="Y:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        ttk.Entry(mouse_box, textvariable=self.fixed_y, width=8).grid(row=2, column=3, sticky="w", padx=5, pady=5)
        ttk.Button(mouse_box, text="Select Position", command=self.select_position).grid(row=2, column=4, padx=10, pady=5)

        #Nudge UI (Off/On + Humanized)
        nudge_box = ttk.LabelFrame(frm, text="Nudge Before Click")
        nudge_box.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        ttk.Radiobutton(nudge_box, text="Off", variable=self.nudge_mode, value="off", command=self._update_nudge_state)\
            .grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(nudge_box, text="On", variable=self.nudge_mode, value="on", command=self._update_nudge_state)\
            .grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(nudge_box, text="ΔX:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.nudge_x_entry = ttk.Entry(nudge_box, textvariable=self.nudge_x, width=8)
        self.nudge_x_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(nudge_box, text="ΔY:").grid(row=1, column=2, sticky="e", padx=5, pady=5)
        self.nudge_y_entry = ttk.Entry(nudge_box, textvariable=self.nudge_y, width=8)
        self.nudge_y_entry.grid(row=1, column=3, sticky="w", padx=5, pady=5)

        self.humanized_chk = ttk.Checkbutton(nudge_box, text="Humanized Mouse Clicks", variable=self.nudge_random)
        self.humanized_chk.grid(row=1, column=4, padx=10, pady=5)

        #Interval (4 boxes)
        interval_box = ttk.LabelFrame(frm, text="Interval (sum of all fields)")
        interval_box.grid(row=4, column=0, sticky="ew", padx=5, pady=5)
        ttk.Label(interval_box, text="Hours:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        ttk.Spinbox(interval_box, from_=0, to=999999, textvariable=self.int_hours, width=8).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(interval_box, text="Minutes:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        ttk.Spinbox(interval_box, from_=0, to=999999, textvariable=self.int_minutes, width=8).grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(interval_box, text="Seconds:").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        ttk.Spinbox(interval_box, from_=0, to=999999, textvariable=self.int_seconds, width=8).grid(row=0, column=5, padx=5, pady=5)
        ttk.Label(interval_box, text="Milliseconds:").grid(row=0, column=6, padx=5, pady=5, sticky="e")
        ttk.Spinbox(interval_box, from_=0, to=999999, textvariable=self.int_millis, width=10).grid(row=0, column=7, padx=5, pady=5)

        #Controls
        ctl_box = ttk.LabelFrame(frm, text="Controls")
        ctl_box.grid(row=5, column=0, sticky="ew", padx=5, pady=5)
        ttk.Button(ctl_box, text="Start", command=self.start_action).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(ctl_box, text="Stop", command=self.stop_action).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(ctl_box, text="Hotkey Settings", command=self.open_hotkey_settings).grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(ctl_box, textvariable=self.status).grid(row=0, column=3, padx=10, pady=5, sticky="w")

        #Recorder
        rec_box = ttk.LabelFrame(frm, text="Recorder")
        rec_box.grid(row=6, column=0, sticky="ew", padx=5, pady=5)
        ttk.Button(rec_box, text="Start Recording (F8)", command=self.toggle_recording).grid(row=0, column=0, padx=5, pady=5)
        ttk.Button(rec_box, text="Play (F9)", command=self.play_recording).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(rec_box, text="Clear", command=self.clear_recording).grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(rec_box, text="Repeat (0 = infinite):").grid(row=0, column=3, padx=5, pady=5)
        ttk.Spinbox(rec_box, from_=0, to=999999, textvariable=self.repeat_count, width=8).grid(row=0, column=4, padx=5, pady=5)

        #Save / Load macros
        ttk.Button(rec_box, text="Save Macro", command=self.save_macro).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(rec_box, text="Load Macro", command=self.load_macro).grid(row=1, column=1, padx=5, pady=5)

        #Footer + "Not Working?"
        footer = ttk.Frame(frm)
        footer.grid(row=7, column=0, sticky="ew", padx=5, pady=(0,5))
        ttk.Label(footer, text="Defaults:F7=Start/Stop | F8=Record | F9=Play  •  Changeable Via Hotkey Settings • Made by Berchia").grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="Not Working?", command=self.show_not_working_help).grid(row=0, column=1, padx=10)

        #Initialize nudge control state
        self._update_nudge_state()

    #Global Hotkeys & Key Capture
    def start_global_listeners(self):
        def on_press(k):
            try:
                name = key_to_str(k)
                if not name:
                    return

                #If we're capturing the next key for "Key to spam", take it & stop
                if self.capturing_spam_key:
                    # Normalize names (e.g., 'return' -> 'enter')
                    if name == "return":
                        name = "enter"
                    self.spam_key.set(name)
                    self.capturing_spam_key = False
                    # move focus away so another accidental key doesn't overwrite
                    self.root.focus()
                    self._schedule_save()
                    return

                #Otherwise, handle hotkeys
                if name == self.action_hotkey.get().lower():
                    self.toggle_action_quick()
                elif name == self.record_hotkey.get().lower():
                    self.toggle_recording()
                elif name == self.play_hotkey.get().lower():
                    self.play_recording()
            except Exception:
                pass

        self.k_listener = keyboard.Listener(on_press=on_press)
        self.k_listener.daemon = True
        self.k_listener.start()

    def _begin_capture_spam_key(self):
        self.capturing_spam_key = True

    #Select Position
    def select_position(self):
        self.status.set("Status: Click anywhere to select position…")
        picked = []

        def on_click(x, y, button, pressed):
            if pressed:
                picked.append((x, y))
                return False  # stop listener

        listener = mouse.Listener(on_click=on_click)
        listener.start()
        listener.join(timeout=10.0)
        if picked:
            x, y = picked[0]
            self.fixed_x.set(int(x))
            self.fixed_y.set(int(y))
            #Auto-switch to Fixed
            self.target_mode.set("fixed")
            self.status.set(f"Status: Fixed position set to ({x}, {y})")
            self._schedule_save()
        else:
            self.status.set("Status: IDLE")

    #Start/Stop Action
    def start_action(self):
        if self.running_event.is_set():
            return
        self.running_event.set()
        t = threading.Thread(target=self._run_action_loop, daemon=True)
        t.start()
        self.status.set(f"Status: RUNNING ({'Key' if self.mode.get()=='key' else 'Mouse'})")
        self._set_active(True)

    def stop_action(self):
        self.running_event.clear()
        self.status.set("Status: IDLE")
        self._set_active(False)

    def toggle_action_quick(self):
        if self.running_event.is_set():
            self.stop_action()
        else:
            self.start_action()

    #Nudge helpers 
    def _update_nudge_state(self):
        enabled = (self.nudge_mode.get() == "on")
        state = "normal" if enabled else "disabled"
        self.nudge_x_entry.configure(state=state)
        self.nudge_y_entry.configure(state=state)
        self.humanized_chk.configure(state=state)
        self._schedule_save()

    def _apply_nudge(self, base_x: int, base_y: int):
        if self.nudge_mode.get() != "on":
            return base_x, base_y
        dx = self.nudge_x.get()
        dy = self.nudge_y.get()
        if self.nudge_random.get():
            dx = random.randint(-abs(dx), abs(dx))
            dy = random.randint(-abs(dy), abs(dy))
        return base_x + dx, base_y + dy

    #Interval helper
    def _interval_seconds(self) -> float:
        #Sum all 4 boxes; any can be 0
        h = max(0, int(self.int_hours.get()))
        m = max(0, int(self.int_minutes.get()))
        s = max(0, int(self.int_seconds.get()))
        ms = max(0, int(self.int_millis.get()))
        total = (h * 3600.0) + (m * 60.0) + s + (ms / 1000.0)
        #Safety minimum sleep
        return max(0.0005, total)

    #Main Action Loop
    def _run_action_loop(self):
        while self.running_event.is_set():
            sleep_s = self._interval_seconds()

            if self.mode.get() == "key":
                key_str = self.spam_key.get().strip().lower()
                k = str_to_key(key_str)
                if k is None:
                    try:
                        if key_str:
                            self.kctl.type(key_str[:1])
                    except Exception:
                        pass
                else:
                    try:
                        self.kctl.press(k)
                        self.kctl.release(k)
                    except Exception:
                        pass

            else:
                btn = Button.left if self.click_button.get() == "left" else Button.right
                count = 2 if self.click_type.get() == "double" else 1

                if self.target_mode.get() == "fixed":
                    base_x, base_y = int(self.fixed_x.get()), int(self.fixed_y.get())
                else:
                    base_x, base_y = self.mctl.position

                try:
                    self.mctl.position = (base_x, base_y)
                    time.sleep(0.002)
                    nx, ny = self._apply_nudge(base_x, base_y)
                    self.mctl.position = (nx, ny)
                    time.sleep(0.001)
                except Exception:
                    pass

                try:
                    self.mctl.click(btn, count)
                except Exception:
                    pass

            time.sleep(sleep_s)

    #Recording
    def toggle_recording(self):
        if self.recording:
            self.stop_recording()
        else:
            self.start_recording()

    def start_recording(self):
        if self.recording:
            return
        self.record_events.clear()
        self.playback_stop.clear()
        self.recording = True
        self.record_start_time = time.perf_counter()
        self.status.set("Status: RECORDING… (F8 to stop)")
        ignore_keys = {self.action_hotkey.get().lower(),
                       self.record_hotkey.get().lower(),
                       self.play_hotkey.get().lower()}

        #Keyboard
        def on_press(k):
            if not self.recording: return
            ks = key_to_str(k)
            if ks in ignore_keys: return
            t = time.perf_counter() - self.record_start_time
            self.record_events.append({"t": t, "type": "key_down", "key": ks})

        def on_release(k):
            if not self.recording: return
            ks = key_to_str(k)
            if ks in ignore_keys: return
            t = time.perf_counter() - self.record_start_time
            self.record_events.append({"t": t, "type": "key_up", "key": ks})

        #Mouse
        def on_move(x, y):
            if not self.recording: return
            t = time.perf_counter() - self.record_start_time
            self.record_events.append({"t": t, "type": "move", "x": int(x), "y": int(y)})

        def on_click(x, y, btn, pressed):
            if not self.recording: return
            t = time.perf_counter() - self.record_start_time
            b = "left" if btn == Button.left else "right"
            self.record_events.append({"t": t, "type": "click", "x": int(x), "y": int(y), "button": b, "pressed": bool(pressed)})

        def on_scroll(x, y, dx, dy):
            if not self.recording: return
            t = time.perf_counter() - self.record_start_time
            self.record_events.append({"t": t, "type": "scroll", "x": int(x), "y": int(y), "dx": int(dx), "dy": int(dy)})

        self.rec_k_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.rec_m_listener = mouse.Listener(on_move=on_move, on_click=on_click, on_scroll=on_scroll)
        self.rec_k_listener.start()
        self.rec_m_listener.start()

    def stop_recording(self):
        if not self.recording:
            return
        self.recording = False
        try:
            self.rec_k_listener.stop()
            self.rec_m_listener.stop()
        except Exception:
            pass
        self.status.set(f"Status: Recorded {len(self.record_events)} events")

    def clear_recording(self):
        self.record_events.clear()
        self.status.set("Status: Recording cleared")

    #Save / Load Macros
    def save_macro(self):
        if not self.record_events:
            messagebox.showinfo("Nothing to save", "No recorded events to save.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Macro",
            defaultextension=".json",
            filetypes=[("Macro JSON", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        payload = {
            "version": 1,
            "created": datetime.utcnow().isoformat() + "Z",
            "meta": {"repeat_suggestion": self.repeat_count.get()},
            "events": self.record_events
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            self.status.set(f"Status: Saved macro → {path}")
        except Exception as e:
            messagebox.showerror("Save Failed", f"Could not save macro:\n{e}")

    def load_macro(self):
        path = filedialog.askopenfilename(
            title="Load Macro",
            filetypes=[("Macro JSON", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            events = payload.get("events")
            if not isinstance(events, list) or not events:
                raise ValueError("No events found in file.")
            self.record_events = events
            meta = payload.get("meta") or {}
            rep = meta.get("repeat_suggestion")
            if isinstance(rep, int) and rep > 0:
                self.repeat_count.set(rep)
            self.status.set(f"Status: Loaded macro ({len(self.record_events)} events)")
        except Exception as e:
            messagebox.showerror("Load Failed", f"Could not load macro:\n{e}")

    #Playback
    def play_recording(self):
        if not self.record_events:
            messagebox.showinfo("Nothing to play", "No recorded events. Click 'Start Recording' first or Load a macro.")
            return
        if self.running_event.is_set():
            self.stop_action()

        self.playback_stop.clear()
        t = threading.Thread(target=self._playback_worker, daemon=True)
        t.start()

    def _playback_worker(self):
        self.status.set("Status: PLAYBACK…")
        self._set_active(True)
        try:
            times = [e.get("t", 0.0) for e in self.record_events]
            t0 = times[0] if times else 0.0

            repeats = self.repeat_count.get()
            infinite = (repeats == 0)
            current = 0

            while infinite or current < repeats:
                if self.playback_stop.is_set():
                    break
                prev_t = 0.0
                for e in self.record_events:
                    if self.playback_stop.is_set():
                        break
                    et = float(e.get("t", 0.0))
                    dt = max(0.0, (et - t0) - prev_t)
                    time.sleep(dt)
                    prev_t = et - t0

                    typ = e.get("type")
                    if typ == "key_down":
                        k = str_to_key(e.get("key", ""))
                        if k is None and e.get("key"):
                            try:
                                self.kctl.press(KeyCode.from_char(e["key"][0]))
                            except Exception:
                                pass
                        elif k is not None:
                            self.kctl.press(k)
                    elif typ == "key_up":
                        k = str_to_key(e.get("key", ""))
                        if k is None and e.get("key"):
                            try:
                                self.kctl.release(KeyCode.from_char(e["key"][0]))
                            except Exception:
                                pass
                        elif k is not None:
                            self.kctl.release(k)
                    elif typ == "move":
                        self.mctl.position = (int(e.get("x", 0)), int(e.get("y", 0)))
                    elif typ == "click":
                        if e.get("pressed", False):
                            btn = Button.left if e.get("button") == "left" else Button.right
                            self.mctl.position = (int(e.get("x", 0)), int(e.get("y", 0)))
                            time.sleep(0.001)
                            self.mctl.click(btn, 1)
                    elif typ == "scroll":
                        self.mctl.position = (int(e.get("x", 0)), int(e.get("y", 0)))
                        self.mctl.scroll(int(e.get("dx", 0)), int(e.get("dy", 0)))
                current += 1
        finally:
            self.status.set("Status: IDLE")
            self._set_active(False)

    #Hotkey Settings
    def open_hotkey_settings(self):
        dlg = Toplevel(self.root)
        dlg.title("Hotkey Settings")
        dlg.grab_set()
        dlg.attributes("-topmost", True)
        pad = {"padx": 8, "pady": 6}

        ttk.Label(dlg, text="Set single-key hotkeys (e.g., f7, f8, f9, a, enter, space)").grid(row=0, column=0, columnspan=3, **pad)

        rows = [
            ("Action Start/Stop:", self.action_hotkey),
            ("Record Start/Stop:", self.record_hotkey),
            ("Playback:", self.play_hotkey),
        ]
        for i, (label, var) in enumerate(rows, start=1):
            ttk.Label(dlg, text=label).grid(row=i, column=0, sticky="e", **pad)
            ent = ttk.Entry(dlg, textvariable=var, width=12)
            ent.grid(row=i, column=1, **pad)
            ttk.Button(dlg, text="Bind…", command=lambda v=var: self.capture_hotkey(v)).grid(row=i, column=2, **pad)

        ttk.Button(dlg, text="Close", command=dlg.destroy).grid(row=len(rows)+1, column=0, columnspan=3, **pad)

    def capture_hotkey(self, target_var: StringVar):
        messagebox.showinfo("Capture Hotkey", "Press the key you want to assign…")
        captured = []
        def on_press(k):
            captured.append(k)
            return False

        listener = keyboard.Listener(on_press=on_press)
        listener.start()
        listener.join(timeout=5.0)
        if captured:
            name = key_to_str(captured[0])
            if name:
                target_var.set(name.lower())
                self._schedule_save()

    #Help Popup
    def show_not_working_help(self):
        messagebox.showinfo(
            "Troubleshooting",
            "Admin rights: Global keyboard/mouse hooks sometimes need elevated permissions on Windows.\n"
            "If hotkeys or recording don’t work, right-click the EXE → Run as administrator."
        )

    #Settings Persistence
    def to_settings_dict(self) -> dict:
        return {
            "geometry": self.root.geometry(),
            "mode": self.mode.get(),
            "spam_key": self.spam_key.get(),
            "click_button": self.click_button.get(),
            "click_type": self.click_type.get(),
            "target_mode": self.target_mode.get(),
            "fixed_x": int(self.fixed_x.get()),
            "fixed_y": int(self.fixed_y.get()),
            "nudge_mode": self.nudge_mode.get(),
            "nudge_x": int(self.nudge_x.get()),
            "nudge_y": int(self.nudge_y.get()),
            "nudge_random": bool(self.nudge_random.get()),
            "action_hotkey": self.action_hotkey.get(),
            "record_hotkey": self.record_hotkey.get(),
            "play_hotkey": self.play_hotkey.get(),
            "repeat_count": int(self.repeat_count.get()),
            # interval 4-box
            "int_hours": int(self.int_hours.get()),
            "int_minutes": int(self.int_minutes.get()),
            "int_seconds": int(self.int_seconds.get()),
            "int_millis": int(self.int_millis.get()),
        }

    def apply_settings(self, d: dict):
        g = lambda k, default=None: d.get(k, default)
        geom = g("geometry")
        if isinstance(geom, str) and "x" in geom:
            try: self.root.geometry(geom)
            except Exception: pass

        self.mode.set(g("mode", self.mode.get()))
        self.spam_key.set(g("spam_key", self.spam_key.get()))
        self.click_button.set(g("click_button", self.click_button.get()))
        self.click_type.set(g("click_type", self.click_type.get()))
        self.target_mode.set(g("target_mode", self.target_mode.get()))
        try:
            self.fixed_x.set(int(g("fixed_x", self.fixed_x.get())))
            self.fixed_y.set(int(g("fixed_y", self.fixed_y.get())))
        except Exception:
            pass

        self.nudge_mode.set(g("nudge_mode", self.nudge_mode.get()))
        try:
            self.nudge_x.set(int(g("nudge_x", self.nudge_x.get())))
            self.nudge_y.set(int(g("nudge_y", self.nudge_y.get())))
        except Exception:
            pass
        self.nudge_random.set(bool(g("nudge_random", self.nudge_random.get())))

        self.action_hotkey.set(g("action_hotkey", self.action_hotkey.get()))
        self.record_hotkey.set(g("record_hotkey", self.record_hotkey.get()))
        self.play_hotkey.set(g("play_hotkey", self.play_hotkey.get()))
        try:
            self.repeat_count.set(int(g("repeat_count", self.repeat_count.get())))
        except Exception:
            pass

        #Interval 4-box
        try:
            self.int_hours.set(int(g("int_hours", self.int_hours.get())))
            self.int_minutes.set(int(g("int_minutes", self.int_minutes.get())))
            self.int_seconds.set(int(g("int_seconds", self.int_seconds.get())))
            self.int_millis.set(int(g("int_millis", self.int_millis.get())))
        except Exception:
            pass

        self._update_nudge_state()

    def save_settings(self):
        data = self.to_settings_dict()
        path = get_settings_path()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[WARN] Failed to save settings to {path}: {e}")

    def load_settings(self):
        path = get_settings_path()
        if not path.exists():
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                self.apply_settings(data)
        except Exception as e:
            print(f"[WARN] Failed to load settings from {path}: {e}")

    #Autosave wiring (traces + geometry debounce)
    def attach_autosave_traces(self):
        vars_to_trace = [
            self.mode, self.spam_key,
            self.click_button, self.click_type, self.target_mode,
            self.nudge_mode, self.nudge_x, self.nudge_y, self.nudge_random,
            self.action_hotkey, self.record_hotkey, self.play_hotkey,
            self.repeat_count, self.fixed_x, self.fixed_y,
            self.int_hours, self.int_minutes, self.int_seconds, self.int_millis
        ]
        for v in vars_to_trace:
            try:
                v.trace_add("write", lambda *args: self._schedule_save())
            except Exception:
                pass

    def _on_configure(self, _event):
        self._schedule_save()

    def _schedule_save(self, delay_ms: int = 250):
        if self._save_after_id is not None:
            try:
                self.root.after_cancel(self._save_after_id)
            except Exception:
                pass
        self._save_after_id = self.root.after(delay_ms, self.save_settings)

    #Cleanup
    def on_close(self):
        self.running_event.clear()
        self.playback_stop.set()
        try:
            self.k_listener.stop()
        except Exception:
            pass
        self.save_settings()
        self.root.destroy()

# -----------------------------
# Run
# -----------------------------
def main():
    root = Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    app = MacroTool(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()

if __name__ == "__main__":
    main()
