import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'C:\Users\Administrator\lobsterai\project\scripts')

import tkinter as tk
root = tk.Tk()
root.withdraw()

import tkinter.filedialog as fd
fd.asksaveasfilename = lambda **kw: os.path.join(os.environ['TEMP'], 'gui_test2.docx')

import tkinter.messagebox as mb
_orig_showinfo = mb.showinfo
_orig_showerror = mb.showerror
def patch_info(title, msg):
    print('Would showinfo:', repr(title[:30]), repr(msg[:50]))
    _orig_showinfo(title, msg)
def patch_error(title, msg):
    print('Would showerror:', repr(title[:30]), repr(msg[:50]))
mb.showinfo = patch_info
mb.showerror = patch_error

from image_to_word import ImageToWordApp

app = ImageToWordApp(root)
app.image_paths = [r'C:\Users\Administrator\lobsterai\project\after_login.png']
app.listbox.insert(0, 'after_login.png')

print('Testing generate...')
try:
    app._generate()
    print('Generate returned OK')
except Exception as e:
    print('ERROR:', type(e).__name__, '-', str(e))
    import traceback
    traceback.print_exc()

root.destroy()
