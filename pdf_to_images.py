import fitz
import os

pdf_path = r"C:\Users\jsber\OneDrive\Documents\Claude_task\indian_restaurant_slides.pdf"
output_dir = r"C:\Users\jsber\OneDrive\Documents\Claude_task"

# Remove old slide images
for f in os.listdir(output_dir):
    if f.startswith("slide-") and f.endswith(".jpg"):
        os.remove(os.path.join(output_dir, f))

doc = fitz.open(pdf_path)
for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=150)
    out_path = os.path.join(output_dir, f"slide-{i+1}.jpg")
    pix.save(out_path)
    print(out_path)

doc.close()
print(f"Exported {len(doc)} slides")
