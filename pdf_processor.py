import fitz  # PyMuPDF
from excel_reader import store_incorrect_bleeds  # Import the function to store incorrect bleeds



def highlight_year_and_bleed_marks(pdf_path, output_path, url):
    year_keywords = [
        "2025",  # English
        "२०२५",  # Hindi/Marathi
        "২০২৫",  # Bengali
        "2025",  # Malayalam (uses English numerals)
        "೨೦೨೫",  # Kannada
        "੨੦੨੫",  # Punjabi
        "２０２５",  # Gujarati
        "2026",  # English
        "२०२६",  # Hindi/Marathi
        "২০২৬",  # Bengali
        "೨೦೨೬",  # Telugu (in Telugu script
    ]

    incorrect_bleeds = []  # List to store incorrect bleeds
    incorrect_rtp_folder = "wrong_rtp"
    incorrect_rtp_file = "incorrect_bleeds.csv"  # Changed to .csv for consistency

    try:
        doc = fitz.open(pdf_path)

        for page_num in range(len(doc)):
            if (page_num + 1) % 2 == 0:  # Process only even-numbered pages
                page = doc[page_num]
                print(f"Processing Page {page_num + 1}...")
                
                text_dict = page.get_text("dict")  # Extract text as a dictionary structure
                drawings = page.get_drawings()      # Get any existing drawings on the page
                
                year_positions = []  # Store bounding boxes of highlighted years
                bleed_marks = []     # Store bounding boxes of bleed marks

                # Iterate through text blocks, lines, and spans to find keywords
                for block in text_dict["blocks"]:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text = span["text"].strip()  # Get text and remove leading/trailing spaces
                            bbox = span["bbox"]           # Get bounding box of the text

                            if any(keyword in text for keyword in year_keywords):
                                page.draw_rect(bbox,color=(1, 0, 0), width=1)  # Highlight in red with no border width
                                year_positions.append(bbox)  # Store adjusted bounding box of highlighted year

                for drawing in drawings:
                    rect = drawing["rect"]
                    if abs(rect[1] - rect[3]) < 2:  # Check if it's a horizontal line (bleed mark)
                        page.draw_rect(rect, color=(0, 1, 0), width=1)  # Highlight bleed marks in green
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

                        if distance_mm < 13:  
                            print(f"Distance less than: {distance_mm:.2f} mm")
                            
                            incorrect_bleeds.append({  
                                "url": url,
                                "distance_mm": round(distance_mm, 2)
                            })
                            store_incorrect_bleeds(incorrect_rtp_folder, incorrect_rtp_file, incorrect_bleeds)  
                            return

    except Exception as e:
        print(f"An error occurred while processing the PDF: {e}")

    finally:
        doc.save(output_path)  # Save the modified PDF with highlights applied
        print(f"Highlights applied and saved to '{output_path}'.")
