import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'C:\Users\Administrator\lobsterai\project\scripts')

import tkinter as tk
root = tk.Tk()
root.withdraw()

# IMPORTANT: patch BEFORE importing image_to_word
import tkinter.filedialog as fd
_orig_asksave = fd.asksaveasfilename
fd.asksaveasfilename = lambda **kw: os.path.join(os.environ['TEMP'], 'gui_test3.docx')

import tkinter.messagebox as mb
_orig_showinfo = mb.showinfo
_orig_showerror = mb.showerror
def safe_info(title, msg):
    print('[showinfo]', repr(title[:20]), repr(msg[:50]))
def safe_error(title, msg):
    print('[showerror]', repr(title[:20]), repr(msg[:50]))
mb.showinfo = safe_info
mb.showerror = safe_error

# Now import - the module will capture the already-patched references
import image_to_word
# Re-patch in the module's own namespace
image_to_word.filedialog.asksaveasfilename = fd.asksaveasfilename
image_to_word.messagebox.showinfo = safe_info
image_to_word.messagebox.showerror = safe_error

app = image_to_word.ImageToWordApp(root)
app.image_paths = [r'C:\Users\Administrator\lobsterai\project\after_login.png']
app.listbox.insert(0, 'after_login.png')

print('Calling _generate...')
try:
    app._generate()
    print('Generate returned OK')
except Exception as e:
    print('ERROR:', type(e).__name__, '-', str(e))
    import traceback
    traceback.print_exc()

print('Calling root.destroy()')
root.destroy()
print('Done')
