from pptx import Presentation
from pptx.util import Inches
from docx import Document
import os

def extract_headings(doc_path):
    doc = Document(doc_path)
    headings = []
    current_section = None
    
    for para in doc.paragraphs:
        if para.style.name.startswith("Heading 1"):
            current_section = {"title": para.text, "slides": []}
            headings.append(current_section)
        elif para.style.name.startswith("Heading 2") and current_section:
            slide_content = {"title": para.text, "text": ""}
            current_section["slides"].append(slide_content)
        elif para.style.name.startswith("Heading 3") or para.style.name.startswith("Heading 4"):
            if current_section and current_section["slides"]:
                current_section["slides"][-1]["text"] += f"\n{para.text}"
    
    return headings

def create_presentation(headings, output_ppt):
    prs = Presentation()
    
    # Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Automatische PowerPoint"
    subtitle.text = "Gegenereerd uit Word-document"
    
    # Summary Slide
    summary_slide = prs.slides.add_slide(prs.slide_layouts[1])
    summary_slide.shapes.title.text = "Overzicht"
    
    # Process Sections (without add_section)
    for section in headings:
        # Section Title Slide
        section_slide = prs.slides.add_slide(prs.slide_layouts[1])
        section_slide.shapes.title.text = section["title"]
        
        for slide_content in section["slides"]:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = slide_content["title"]
            slide.placeholders[1].text = slide_content["text"]
    
    # Thank You Slide
    thank_you_slide = prs.slides.add_slide(prs.slide_layouts[1])
    thank_you_slide.shapes.title.text = "Bedankt voor uw aandacht!"
    thank_you_slide.placeholders[1].text = "Vragen? Neem contact op."
    
    # Save presentation
    prs.save(output_ppt)
    print(f"PowerPoint opgeslagen als {output_ppt}")

# Example usage
doc_path = "input.docx"  # Vervang dit door jouw Word-bestand
output_ppt = "output.pptx"
headings = extract_headings(doc_path)
create_presentation(headings, output_ppt)
