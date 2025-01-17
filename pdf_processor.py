import fitz  # PyMuPDF

def highlight_year_and_bleed_marks(pdf_path, output_path):
    year_keywords = [
        "2025",  # English
        "२०२५",  # Hindi/Marathi
        "২০২৫",  # Bengali
        "2025",  # Malayalam (uses English numerals)
        "೨೦೨೫",  # Kannada
        "੨੦੨੫",  # Punjabi
        "２０２５",  # Gujarati
    ]

    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        if (page_num + 1) % 2 == 0:
            page = doc[page_num]
            print(f"Processing Page {page_num + 1}...")
            
            text_dict = page.get_text("dict")
            drawings = page.get_drawings()
            
            year_positions = []
            bleed_marks = []

            for block in text_dict["blocks"]:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].strip()
                        # text = span["text"].strip().replace(" ", "")  # Remove all spaces
                        bbox = span["bbox"]

                        if any(keyword in text for keyword in year_keywords):
                            # print(f"Highlighting Year: {text} at {bbox}")
                            page.draw_rect(bbox, color=(1, 0, 0), width=2)
                            year_positions.append(bbox)

            for drawing in drawings:
                rect = drawing["rect"]
                if abs(rect[1] - rect[3]) < 2:
                    # print(f"Highlighting Bleed Mark at {rect}")
                    page.draw_rect(rect, color=(0, 1, 0), width=2)
                    bleed_marks.append(rect)

            for year_bbox in year_positions:
                top_y = year_bbox[1]
                closest_bleed_mark = None
                closest_distance = float('inf')

                for rect in bleed_marks:
                    bottom_y = rect[3]
                    if bottom_y < top_y: 
                        distance_points = top_y - bottom_y

                        if distance_points < closest_distance:
                            closest_distance = distance_points
                            closest_bleed_mark = rect

                if closest_bleed_mark is not None:
                    distance_mm = closest_distance * (25.4 / 72)
                    print(f"Distance: {distance_mm:.2f} mm")
                    mid_x = (year_bbox[0] + year_bbox[2]) / 2
                    bottom_y_closest = closest_bleed_mark[3]
                    page.draw_line((mid_x, top_y), (mid_x, bottom_y_closest), color=(0, 0, 1), width=2)

    doc.save(output_path)
    print(f"Highlights applied and saved to '{output_path}'.")
