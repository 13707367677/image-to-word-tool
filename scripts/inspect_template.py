import zipfile, sys, re
sys.stdout.reconfigure(encoding='utf-8')
path = r'C:\Users\ADMINI~1\AppData\Local\Temp\lobsterai\attachments\照片模板-1774504235078-0c4sa6.docx'
with zipfile.ZipFile(path) as z:
    doc_xml = z.read('word/document.xml').decode('utf-8')

all_extents = re.findall(r'cx="(\d+)"\s+cy="(\d+)"', doc_xml)
unique = set(all_extents)
print('Unique extents:')
for cx, cy in unique:
    w_cm = int(cx) / 914400 * 2.54
    h_cm = int(cy) / 914400 * 2.54
    print('  cx=' + cx + ' cy=' + cy + ' => ' + ('%.2f' % w_cm) + 'cm x ' + ('%.2f' % h_cm) + 'cm')

print()
# Check image sizes
from docx import Document
from PIL import Image
import io

doc = Document(path)
t = doc.tables[0]

# Look at row structure more carefully
# Get image in first cell
with zipfile.ZipFile(path) as z:
    rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
    img_refs = re.findall(r'Id="(rId\d+)"[^>]*?Target="media/(image\d+\.jpeg)"', rels_xml)
    print('Image refs sample:')
    for rid, fname in img_refs[:5]:
        print('  ' + rid + ' -> ' + fname)

# Get actual image dimensions from first few images
print()
print('Actual image dimensions:')
with zipfile.ZipFile(path) as z:
    for fname in ['image1.jpeg', 'image2.jpeg', 'image3.jpeg']:
        data = z.read('word/media/' + fname)
        img = Image.open(io.BytesIO(data))
        print('  ' + fname + ': ' + str(img.size) + ' ' + str(img.format))
