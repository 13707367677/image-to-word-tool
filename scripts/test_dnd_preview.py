"""最小化测试：拖拽 + 预览"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'C:\Users\Administrator\lobsterai\project\scripts')

import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from PIL import Image, ImageTk

# Test 1: DnD 窗口是否正常创建
print("=== Test 1: TkinterDnD window ===")
root = TkinterDnD.Tk()
root.title("DnD Test")
root.geometry("300x200")
print("TkinterDnD.Tk() OK")

# Test 2: 注册拖拽目标
label = tk.Label(root, text="拖拽文件到这", bg="#e8f4fd", font=("微软雅黑", 14))
label.pack(fill="both", expand=True, padx=20, pady=20)

def on_drop(event):
    print("DROP EVENT DATA:", repr(event.data))
    files = []
    cleaned = (event.data or "").strip().strip('{}')
    for p in cleaned.split():
        p = p.strip('"')
        if os.path.isfile(p) and p.lower().endswith(('.png', '.jpg', '.jpeg')):
            files.append(p)
    print("Parsed files:", files)
    if files:
        label.config(text="收到: " + os.path.basename(files[0]), bg="#d0f0c0")

label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', on_drop)
print("DnD registered on label OK")

# Test 3: 预览
print("\n=== Test 3: Preview ===")
preview_canvas = tk.Canvas(root, bg="#e8e0d8", width=280, height=200)
preview_canvas.pack(pady=10)

def load_preview():
    # 使用实际存在的图片
    test_img = r'C:\Users\Administrator\lobsterai\project\after_login.png'
    if not os.path.exists(test_img):
        print("Test image not found")
        return
    print("Loading:", test_img)
    pil_img = Image.open(test_img)
    print("Image size:", pil_img.size)
    pil_img.thumbnail((260, 180), Image.LANCZOS)
    print("Thumbnail size:", pil_img.size)
    photo = ImageTk.PhotoImage(pil_img)
    print("PhotoImage created:", photo.width(), "x", photo.height())
    preview_canvas.delete("all")
    preview_canvas.create_image(140, 100, anchor="center", image=photo)
    preview_canvas.image = photo  # 保持引用
    label.config(text="预览加载成功: " + str(pil_img.size))

btn = tk.Button(root, text="加载预览", command=load_preview)
btn.pack()

root.update_idletasks()
print("Canvas dims:", preview_canvas.winfo_width(), "x", preview_canvas.winfo_height())

print("\nAll tests passed - window should be visible")
root.mainloop()
