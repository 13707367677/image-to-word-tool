import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'C:\Users\Administrator\lobsterai\project\scripts')

import tkinter as tk
from PIL import Image, ImageTk

root = tk.Tk()
root.geometry("400x300")
root.update_idletasks()

# Simulate the preview canvas
canvas = tk.Canvas(root, bg="#e8e0d8", width=380, height=280)
canvas.pack(fill="both", expand=True)
canvas.update_idletasks()

print("Canvas dims:", canvas.winfo_width(), "x", canvas.winfo_height())

cw = canvas.winfo_width()
ch = canvas.winfo_height()
print("After update:", cw, ch)

# Simulate preview
if cw <= 1 or ch <= 1:
    print("Canvas too small, would use delay")
else:
    print("Canvas OK, would render image")

# Test ImageTk
img = Image.new("RGB", (200, 150), "red")
photo = ImageTk.PhotoImage(img)
canvas.create_image(cw//2, ch//2, anchor="center", image=photo)

root.mainloop()
