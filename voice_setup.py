import threading
from tkinter import messagebox

import customtkinter as ctk

from audio_devices import selected_input_device_name
from voice_clone import record_reference_voice


class VoiceEnrollmentDialog(ctk.CTkToplevel):
    def __init__(self, master, on_enrolled, before_record=None, on_close=None):
        super().__init__(master)
        self.on_enrolled = on_enrolled
        self.before_record = before_record
        self.on_close = on_close
        self.recording = False

        self.title("Set up your voice")
        self.geometry("510x380")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.close)

        ctk.CTkLabel(
            self,
            text="Make Alpha speak in your voice",
            font=ctk.CTkFont(size=19, weight="bold"),
        ).pack(anchor="w", padx=22, pady=(20, 10))
        ctk.CTkLabel(
            self,
            text=(
                "When you press Record, Alpha captures 12 seconds from the "
                f"selected microphone ({selected_input_device_name()}). "
                "Choose MIC SOURCE first if this is not the microphone you "
                "will speak into. Read the sentence below in a clear, natural voice:"
            ),
            wraplength=460,
            justify="left",
        ).pack(anchor="w", padx=22, pady=5)
        ctk.CTkLabel(
            self,
            text=(
                "“Hello, this is my voice. I am speaking clearly for Alpha. "
                "Please listen to me. I can ask Alpha a question, and Alpha "
                "will reply in my voice.”"
            ),
            wraplength=450,
            justify="left",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", padx=22, pady=10)
        ctk.CTkLabel(
            self,
            text=(
                "The recording is saved locally on this PC and used only to "
                "synthesize Alpha’s replies. It is not uploaded. XTTS model "
                "weights download locally when needed. Use only your own voice."
            ),
            wraplength=460,
            justify="left",
        ).pack(anchor="w", padx=22, pady=8)
        self.status = ctk.CTkLabel(self, text="Ready to record.")
        self.status.pack(anchor="w", padx=22, pady=(4, 8))
        self.record_button = ctk.CTkButton(
            self,
            text="Record my voice",
            command=self.start_recording,
        )
        self.record_button.pack(anchor="e", padx=22, pady=(4, 18))

    def start_recording(self):
        if self.before_record:
            try:
                self.before_record()
            except Exception as error:
                messagebox.showerror("Microphone unavailable", str(error), parent=self)
                return

        self.recording = True
        self.record_button.configure(state="disabled", text="Recording…")
        self.status.configure(text="Recording now — read the sentence above.")
        threading.Thread(target=self._record, daemon=True).start()

    def _record(self):
        try:
            sample_path = record_reference_voice()
        except Exception as error:
            self.after(0, lambda detail=str(error): self._recording_failed(detail))
        else:
            self.after(0, lambda path=sample_path: self._recording_finished(path))

    def _recording_failed(self, error):
        self.recording = False
        self.record_button.configure(state="normal", text="Try recording again")
        self.status.configure(text="Voice recording did not complete.")
        messagebox.showerror("Could not record your voice", error, parent=self)

    def _recording_finished(self, sample_path):
        try:
            self.on_enrolled(sample_path)
        except Exception as error:
            self.recording = False
            self.record_button.configure(state="normal", text="Try again")
            self.status.configure(text="Could not save your voice.")
            messagebox.showerror("Could not save your voice", str(error), parent=self)
            return
        self.recording = False
        self.close()

    def close(self):
        if self.recording:
            return
        self.destroy()
        if self.on_close:
            self.on_close()
