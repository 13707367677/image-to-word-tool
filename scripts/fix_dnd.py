import re

with open(r'C:\Users\Administrator\lobsterai\project\scripts\image_to_word.py', 'r', encoding='utf-8') as f:
    content = f.read()

print('Before: HAS_DND=', 'HAS_DND' in content, 'TkinterDnD=', 'TkinterDnD' in content)

# Step 1: Remove the HAS_DND / TkinterDnD import block
# Find and remove the try/except block for tkinterdnd2
pattern = r"try:\s+\n\s+from tkinterdnd2 import [^\n]+\n\s+HAS_DND = True\s+except ImportError:\s+HAS_DND = False\s+\n"
content = re.sub(pattern, '', content)

# Step 2: Change _setup_dnd() call to _setup_drop_hint()
content = content.replace('self._setup_dnd()', 'self._setup_drop_hint()')

# Step 3: Replace _setup_dnd method with _setup_drop_hint
old_method = '''    def _setup_dnd(self):
        if not HAS_DND:
            self.drop_frame.bind('<Button-1>', lambda e: self._add_images())
            self.drop_label.config(text="点击选择图片文件", cursor="hand2")
            return

        # DnD 注册在 listbox 上（最直观的放置区域）
        for widget in [self.listbox, self.drop_frame]:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', self._on_drop_dnd)
            except Exception:
                pass
        self.drop_label.config(text="拖拽图片到列表或此处即可添加")

    def _setup_dnd_fallback(self):
        self.drop_frame.bind('<Button-1>', lambda e: self._add_images())
        self.drop_label.config(text="点击选择图片文件", cursor="hand2")

    def _on_drop_dnd(self, event):
        """tkinterdnd2 拖拽回调"""
        import re
        files = []
        cleaned = (event.data or "").strip()
        # Windows DnD may wrap paths in {}
        if cleaned.startswith('{') and cleaned.endswith('}'):
            cleaned = cleaned[1:-1]
        for p in re.split(r'\s+', cleaned):
            p = p.strip('"').strip("'")
            if os.path.isfile(p) and self._is_image(p):
                files.append(p)
        if files:
            self._add_files(files)

    def _is_image(self, path):
        ext = os.path.splitext(path)[1].lower()
        return ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')

    # ── 功能 ─────────────────────────────────────────────────
    def _init_preview(self):'''

new_method = '''    # ── 拖拽区域提示 ──────────────────────────────────────────
    def _setup_drop_hint(self):
        self.drop_frame.bind('<Button-1>', lambda e: self._add_images())
        self.drop_label.config(text="点击选择图片文件", cursor="hand2")

    # ── 功能 ─────────────────────────────────────────────────
    def _init_preview(self):'''

if old_method in content:
    content = content.replace(old_method, new_method)
    print('Step 3: replaced _setup_dnd method OK')
else:
    print('Step 3: old method NOT found, trying alternate...')
    # Try without the Chinese comments
    old_method2 = '''    def _setup_dnd(self):
        if not HAS_DND:
            self.drop_frame.bind('<Button-1>', lambda e: self._add_images())
            self.drop_label.config(text="点击选择图片文件", cursor="hand2")
            return

        # DnD 注册在 listbox 上
        for widget in [self.listbox, self.drop_frame]:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', self._on_drop_dnd)
            except Exception:
                pass
        self.drop_label.config(text="拖拽图片到列表或此处即可添加")

    def _setup_dnd_fallback(self):
        self.drop_frame.bind('<Button-1>', lambda e: self._add_images())
        self.drop_label.config(text="点击选择图片文件", cursor="hand2")

    def _on_drop_dnd(self, event):
        """tkinterdnd2 拖拽回调"""
        import re
        files = []
        cleaned = (event.data or "").strip()
        if cleaned.startswith('{') and cleaned.endswith('}'):
            cleaned = cleaned[1:-1]
        for p in re.split(r'\s+', cleaned):
            p = p.strip('"').strip("'")
            if os.path.isfile(p) and self._is_image(p):
                files.append(p)
        if files:
            self._add_files(files)

    def _is_image(self, path):
        ext = os.path.splitext(path)[1].lower()
        return ext in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')

    # ── 功能 ─────────────────────────────────────────────────
    def _init_preview(self):'''
    if old_method2 in content:
        content = content.replace(old_method2, new_method)
        print('Step 3: replaced with alternate text')
    else:
        print('Step 3: still not found')

# Step 4: Replace bottom TkinterDnD with plain Tk
old_bottom = '''    # 使用 TkinterDnD.Tk 以支持拖拽文件
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()'''
new_bottom = '''    # 普通 Tk 窗口
    root = tk.Tk()'''

if old_bottom in content:
    content = content.replace(old_bottom, new_bottom)
    print('Step 4: replaced bottom OK')
else:
    print('Step 4: bottom NOT found')

with open(r'C:\Users\Administrator\lobsterai\project\scripts\image_to_word.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('After: HAS_DND=', 'HAS_DND' in content, 'TkinterDnD=', 'TkinterDnD' in content)
print('Done')
