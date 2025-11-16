import tkinter as tk

class CircularProgressBar(tk.Canvas):
    def __init__(self, parent, size=120, width=12, fg="#4CAF50", bg="#ccc"):
        super().__init__(parent, width=size, height=size, bg=parent["bg"], highlightthickness=0)
        
        self.size = size
        self.width = width

        self.create_oval(
            width//2, width//2, 
            size-width//2, size-width//2, 
            outline=bg, width=width
        )

        self.arc = self.create_arc(
            width//2, width//2, 
            size-width//2, size-width//2,
            start=-90, extent=0,
            style="arc", outline=fg, width=width
        )

        self.text = self.create_text(
            size//2, size//2,
            text="0%", font=("Segoe UI", 18, "bold"), fill="#333"
        )

    def update_progress(self, value):
        value = max(0, min(100, value))
        self.itemconfig(self.arc, extent=value * 3.6)
        self.itemconfig(self.text, text=f"{int(value)}%")
        self.update_idletasks()
