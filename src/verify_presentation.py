"""
Verify presentation layout, text boundaries, and speaker notes.
"""
from pptx import Presentation

prs = Presentation('docs/P054_Presentation.pptx')
print(f"Total slides: {len(prs.slides)}")
print(f"Dimensions: {prs.slide_width.inches} x {prs.slide_height.inches}")

for idx, slide in enumerate(prs.slides, 1):
    print(f"\n=================== SLIDE {idx} ===================")
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
        print(f"Speaker Notes ({len(notes)} chars): {notes[:90]}...")
    else:
        print("WARNING: No notes slide!")
        
    for s_idx, shape in enumerate(slide.shapes, 1):
        w = shape.width.inches
        h = shape.height.inches
        l = shape.left.inches
        t = shape.top.inches
        
        # Check boundaries
        if l < 0 or t < 0 or (l + w) > 10.01 or (t + h) > 5.63:
            print(f"  WARNING: Shape {s_idx} out of slide boundaries: ({l:.2f}, {t:.2f}, {w:.2f}, {h:.2f})")
            
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                txt = p.text.strip()
                if txt:
                    # Estimate line count: approximate 14 chars per inch at 12pt
                    est_chars_per_line = max(10, int(w * 14))
                    est_lines = max(1, len(txt) // est_chars_per_line + (1 if len(txt) % est_chars_per_line else 0))
                    est_height = est_lines * 0.22 # inches
                    overflow = est_height > (h + 0.1)
                    flag = " [OVERFLOW RISK]" if overflow else ""
                    print(f"  Box ({l:.2f}, {t:.2f}, w={w:.2f}, h={h:.2f}){flag}: {txt[:60]}... ({len(txt)} chars)")
        elif shape.has_table:
            print(f"  Table ({l:.2f}, {t:.2f}, w={w:.2f}, h={h:.2f}): {len(shape.table.rows)} rows x {len(shape.table.columns)} cols")
        elif hasattr(shape, 'image'):
            print(f"  Image ({l:.2f}, {t:.2f}, w={w:.2f}, h={h:.2f})")
