# gui.py
import tkinter as tk
from tkinter import ttk
import asyncio, threading
from kindleReader import main_loop
import config
from utils import get_voice_list, load_voices

class KindleTTSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kindle TTS Reader")
        self.root.geometry("900x500")
        self.root.minsize(800, 450)

        # center window
        self.root.update_idletasks()
        width = 900
        height = 500
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

        self.loop = asyncio.new_event_loop()
        self.stop_event = threading.Event()
        self.task = None

        # Main layout: left (settings), right (voice list)
        main = ttk.Frame(root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(main)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        right = ttk.Frame(main)
        right.pack(side=tk.RIGHT, fill=tk.Y)

        # --- LEFT PANEL: TTS SETTINGS + CONTROL ---
        ttk.Label(left, text="TTS Settings", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 5))

        ttk.Label(left, text="Voice:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.voice_var = tk.StringVar(value=config.TTS_VOICE)
        ttk.Entry(left, textvariable=self.voice_var, width=35).grid(row=1, column=1, sticky="w")

        ttk.Label(left, text="Rate:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.rate_var = tk.StringVar(value=config.TTS_RATE)
        ttk.Entry(left, textvariable=self.rate_var, width=35).grid(row=2, column=1, sticky="w")

        ttk.Label(left, text="Volume:").grid(row=3, column=0, sticky=tk.W, pady=3)
        self.volume_var = tk.StringVar(value=config.TTS_VOLUME)
        ttk.Entry(left, textvariable=self.volume_var, width=35).grid(row=3, column=1, sticky="w")

        self.tcp_var = tk.BooleanVar(value=config.TTS_USE_TCP)
        ttk.Checkbutton(left, text="Use TCP mode", variable=self.tcp_var).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

        ttk.Separator(left).grid(row=5, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Label(left, text="Controls", font=("Segoe UI", 12, "bold")).grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 5))

        control_frame = ttk.Frame(left)
        control_frame.grid(row=7, column=0, columnspan=2, pady=(0, 5), sticky="w")

        self.start_btn = ttk.Button(control_frame, text="▶ Start Reading", command=self.start)
        self.start_btn.grid(row=0, column=0, padx=3)

        self.stop_btn = ttk.Button(control_frame, text="■ Stop", command=self.stop, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=3)

        self.add_voice_btn = ttk.Button(control_frame, text="➕ Load Voices", command=self.add_voices)
        self.add_voice_btn.grid(row=0, column=2, padx=3)

        ttk.Button(control_frame, text="Exit", command=self.exit_program).grid(row=0, column=3, padx=3)

        ttk.Separator(left).grid(row=8, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Label(left, text="Status", font=("Segoe UI", 12, "bold")).grid(row=9, column=0, columnspan=2, sticky="w", pady=(0, 5))
        self.status_var = tk.StringVar(value="Stopped")
        ttk.Label(left, textvariable=self.status_var, foreground="gray").grid(row=10, column=0, columnspan=2, sticky="w")

        # --- RIGHT PANEL: VOICE LIST ---
        ttk.Label(right, text="Available Voices", font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(right, textvariable=self.search_var)
        self.search_entry.pack(fill=tk.X, pady=(4, 6))
        self.search_var.trace_add('write', lambda *_: self._filter_voices())

        self.voice_listbox = tk.Listbox(right, height=20, width=40)
        self.voice_listbox.pack(fill=tk.BOTH, expand=True)
        self.voice_listbox.bind('<Double-1>', lambda e: self._select_voice_from_list())

        self.select_voice_btn = ttk.Button(right, text="Select Voice", command=self._select_voice_from_list)
        self.select_voice_btn.pack(fill=tk.X, pady=(6, 0))

        # load persisted voices
        self._voices_store = load_voices()
        self._all_voice_items = []
        self._populate_voice_list()

    def start(self):
        config.TTS_VOICE = self.voice_var.get()
        config.TTS_RATE = self.rate_var.get()
        config.TTS_VOLUME = self.volume_var.get()
        config.TTS_USE_TCP = self.tcp_var.get()

        self.stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        def runner():
            asyncio.set_event_loop(self.loop)
            self.task = self.loop.create_task(main_loop(self.stop_event))
            self.loop.run_until_complete(self.task)

        threading.Thread(target=runner, daemon=True).start()
        self.status_var.set("Running — open Kindle window")

    def _fetch_voices_thread(self):
        try:
            self.root.after(0, lambda: self.status_var.set("Fetching voices..."))
            store = asyncio.run(get_voice_list())

            def finish(store=store):
                self._voices_store = store
                self._populate_voice_list()
                count = len(getattr(store, 'voices', []))
                self.status_var.set(f"Voices added: {count}")

            self.root.after(0, finish)
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Voice fetch failed: {e}"))

    def add_voices(self):
        threading.Thread(target=self._fetch_voices_thread, daemon=True).start()

    def _populate_voice_list(self):
        self.voice_listbox.delete(0, tk.END)
        self._all_voice_items = []
        if not self._voices_store:
            return
        for v in getattr(self._voices_store, 'voices', []):
            label = f"{v.name} — {getattr(v, 'locale', '')}"
            self._all_voice_items.append((label, v))
            self.voice_listbox.insert(tk.END, label)

    def _filter_voices(self):
        q = self.search_var.get().lower().strip()
        self.voice_listbox.delete(0, tk.END)
        for label, v in self._all_voice_items:
            if not q or q in label.lower():
                self.voice_listbox.insert(tk.END, label)

    def _select_voice_from_list(self):
        sel = self.voice_listbox.curselection()
        if not sel:
            return
        label = self.voice_listbox.get(sel[0])
        for lab, v in self._all_voice_items:
            if lab == label:
                self.voice_var.set(v.name)
                config.TTS_VOICE = v.name
                self.status_var.set(f"Selected: {v.name}")
                return

    def stop(self):
        self.stop_event.set()
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("Stopped")

    def exit_program(self):
        self.stop()
        self.root.destroy()
