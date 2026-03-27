import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, r'C:\Users\Administrator\lobsterai\project\scripts')

# Patch tkinter for headless
import tkinter as tk
root = tk.Tk()
root.withdraw()

# Now import and test the app
from image_to_word import ImageToWordApp, set_cell_width, add_cell_image, add_cell_text, add_cell_text

# Patch messagebox and filedialog to avoid blocking
import tkinter.messagebox, tkinter.filedialog
_orig_ask = tkinter.filedialog.asksaveasfilename
tkinter.filedialog.asksaveasfilename = lambda **kw: os.path.join(os.environ['TEMP'], 'test_output.docx')
tkinter.messagebox.showinfo = lambda **kw: None
tkinter.messagebox.showwarning = lambda **kw: None
tkinter.messagebox.askyesno = lambda **kw: True

# Create app and add test images
app = ImageToWordApp(root)
app.image_paths = [
    r'C:\Users\Administrator\lobsterai\project\after_login.png',
    r'C:\Users\Administrator\lobsterai\project\api_call_page.png',
    r'C:\Users\Administrator\lobsterai\project\final_page.png',
]
app.listbox.insert(0, 'after_login.png')
app.listbox.insert(1, 'api_call_page.png')
app.listbox.insert(2, 'final_page.png')

print('Testing _generate...')
app._generate()
print('Done')
root.destroy()
