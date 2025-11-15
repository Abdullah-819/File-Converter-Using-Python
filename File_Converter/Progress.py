# progress.py
import tkinter as tk

class CircularProgressBar(tk.Canvas):
    def __init__(self, parent, size=120, width=12, fg="#4CAF50", bg="#ddd", text_color="#333"):
        super().__init__(parent, width=size, height=size, bg=parent["bg"], highlightthickness=0)
        self.size = size
        self.width = width
        self.fg = fg
        self.bg = bg
        self.text_color = text_color

        self.create_oval(width//2, width//2, size-width//2, size-width//2, outline=self.bg, width=self.width)
        self.arc = self.create_arc(width//2, width//2, size-width//2, size-width//2, start=-90, extent=0,
                                   style="arc", outline=self.fg, width=self.width)
        self.text = self.create_text(size//2, size//2, text="0%", font=("Arial",18,"bold"), fill=self.text_color)

    def update_progress(self, percent):
        percent = max(0, min(100, percent))
        self.itemconfig(self.arc, extent=percent*3.6)
        self.itemconfig(self.text, text=f"{int(percent)}%")
        self.update_idletasks()
