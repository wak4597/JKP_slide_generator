import streamlit as st
import io
import re
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN

# ==========================================
# 1. CORE PARSING LOGIC (WORD TO DICTIONARY)
# ==========================================
def classify_decision(ulasan_text):
    """Classifies the long text in the ISU/ULASAN column into JKP standard decisions."""
    text = ulasan_text.upper()
    if "TIDAK DAPAT DIPERTIMBANGKAN" in text or "TIDAK DAPAT DIPROSES" in text:
        return "TIDAK DAPAT DIPERTIMBANGKAN"
    elif "TARIK BALIK" in text:
        return "TARIK BALIK"
    elif "PELUPUSAN" in text:
        return "PELUPUSAN"
    elif "TERMA ATAU SYARAT" in text:
        return "DIPERAKUI DENGAN TERMA ATAU SYARAT"
    elif "DIPERAKUI SETELAH" in text:
        return "DIPERAKUI SETELAH MAKLUM BALAS"
    elif "DIPERAKUI" in text or "TIADA HALANGAN" in text or "BOLEH DISOKONG" in text:
        return "DIPERAKUI"
    else:
        return "LAIN-LAIN"

def parse_docx(file_buffer):
    """Parses the uploaded Word document, interleaving paragraphs and tables to maintain context."""
    doc = Document(file_buffer)
    parsed_data = []
    
    current_borang = "LAIN-LAIN PERMOHONAN"
    current_category = "HOSPITAL SWASTA" # Default
    
    # Iterate through the underlying XML to preserve the order of paragraphs and tables
    for child in doc.element.body:
        # Check if the element is a Paragraph
        if isinstance(child, CT_P):
            p = Paragraph(child, doc)
            text = p.text.strip().upper()
            if text and ("BORANG" in text or "PERMOHONAN" in text or "HAL-HAL LAIN" in text):
                # Clean up the section name for the slide title
                current_borang = p.text.strip()
                
        # Check if the element is a Table
        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                
                if len(cells) < 2:
                    continue
                
                col1_text = cells[1].upper()
                
                # Detect Category Row (Usually merged cells or empty trailing columns)
                if "SWASTA" in col1_text or "KEMUDAHAN" in col1_text or "BROWNFIELD" in col1_text:
                    current_category = cells[1].strip()
                    continue
                
                # Skip header rows
                if "SENARAI PERMOHONAN" in col1_text or not cells[1]:
                    continue
                
                # Parse actual data rows
                if len(cells) >= 3:
                    facility_details = cells[1]
                    ulasan = cells[2]
                    
                    # Facility name is typically the first line
                    facility_name = facility_details.split('\n')[0].strip()
                    
                    # Remove common prefixes like numbers or "Pemohon:"
                    facility_name = re.sub(r'^\d+\.\s*', '', facility_name)
                    facility_name = re.sub(r'^Pemohon:\s*', '', facility_name)
                    
                    status = classify_decision(ulasan)
                    
                    if facility_name:
                        parsed_data.append({
                            "borang": current_borang,
                            "category": current_category,
                            "facility": facility_name,
                            "status": status
                        })
                        
    return parsed_data

# ==========================================
# 2. SLIDE GENERATION LOGIC
# ==========================================
def generate_pptx(parsed_data, template_buffer):
    """Generates a PPTX file using the parsed data and the user's template."""
    
    # Load the template if provided, otherwise use default python-pptx blank
    if template_buffer:
        prs = Presentation(template_buffer)
    else:
        prs = Presentation()
        
    # Standard PPTX layouts: 0=Title, 1=Title and Content, 5=Title Only, 6=Blank
    try:
        title_layout = prs.slide_layouts[0]
        content_layout = prs.slide_layouts[1] 
    except IndexError:
        # Fallback if the template has weird layouts
        title_layout = prs.slide_layouts[0]
        content_layout = prs.slide_layouts[0]

    # Organize data into a hierarchy: Borang -> Category -> Status -> List of Facilities
    grouped_data = {}
    for item in parsed_data:
        b = item['borang']
        c = item['category']
        s = item['status']
        f = item['facility']
        
        if b not in grouped_data: grouped_data[b] = {}
        if c not in grouped_data[b]: grouped_data[b][c] = {}
        if s not in grouped_data[b][c]: grouped_data[b][c][s] = []
        
        grouped_data[b][c][s].append(f)

    # 1. Add Auto-Generated Intro Slide
    slide = prs.slides.add_slide(title_layout)
    if slide.shapes.title:
        slide.shapes.title.text = "Laporan Slaid Automatik JKP"
    if len(slide.placeholders) > 1:
        slide.placeholders[1].text = f"Jumlah Permohonan Diproses: {len(parsed_data)}"

    # 2. Iterate and generate content slides
    for borang, categories in grouped_data.items():
        
        # Determine total applications in this Borang
        total_borang_apps = sum([len(facs) for cats in categories.values() for facs in cats.values()])
        
        for category, statuses in categories.items():
            slide = prs.slides.add_slide(content_layout)
            
            # Set slide Title (Borang + Category)
            if slide.shapes.title:
                slide.shapes.title.text = f"{borang}\n{category} ({total_borang_apps})"
                
            # If the template layout has a body text placeholder, use it
            if len(slide.placeholders) > 1:
                body_shape = slide.placeholders[1]
                tf = body_shape.text_frame
                tf.clear()
                
                for status, facilities in statuses.items():
                    # Add Status Header
                    p = tf.add_paragraph()
                    p.text = f"{status} ({len(facilities)}):"
                    p.font.bold = True
                    p.font.size = Pt(18)
                    # NOTE: Removed the font color override here so it doesn't cause RGBColor errors
                    
                    # Add Facility Items
                    for idx, facility in enumerate(facilities, 1):
                        p_fac = tf.add_paragraph()
                        p_fac.text = f"  {idx}. {facility}"
                        p_fac.level = 1
                        p_fac.font.size = Pt(14)
            else:
                # Fallback: Manually add a text box if layout is missing standard placeholders
                txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
                tf = txBox.text_frame
                tf.text = f"{category} - Data Parsed successfully, but template layout lacked a content placeholder. Please adjust Master Slide."

    # Save to binary buffer for Streamlit download
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. STREAMLIT WEB INTERFACE
# ==========================================
st.set_page_config(page_title="JKP Slide Generator", page_icon="📑", layout="centered")

st.title("📑 JKP Slide Generator (Automated)")
st.write("Sila muat naik dokumen Word (Lampiran 2) dan Templat PowerPoint mesyuarat untuk menjana slaid secara automatik.")

# Drag and drop file uploaders
col1, col2 = st.columns(2)
with col1:
    docx_file = st.file_uploader("1. Muat Naik Dokumen Word (.docx)", type=["docx"])
with col2:
    pptx_template = st.file_uploader("2. Muat Naik Templat Slaid (.pptx) (Pilihan)", type=["pptx"])

# Processing state
if st.button("Jana Slaid Sekarang", type="primary", use_container_width=True):
    if not docx_file:
        st.error("Sila muat naik dokumen Word (.docx) terlebih dahulu!")
    else:
        with st.spinner("Mengekstrak data dari dokumen..."):
            try:
                # 1. Parse the Word file
                parsed_data = parse_docx(docx_file)
                
                if len(parsed_data) == 0:
                    st.warning("Tiada permohonan ditemui dalam dokumen ini. Sila pastikan format jadual adalah betul.")
                else:
                    st.success(f"Berjaya mengekstrak {len(parsed_data)} permohonan!")
                    
                    # 2. Generate the PowerPoint
                    with st.spinner("Menjana Slaid PowerPoint..."):
                        pptx_buffer = generate_pptx(parsed_data, pptx_template)
                        
                        st.balloons()
                        st.download_button(
                            label="📥 Muat Turun Slaid (.pptx)",
                            data=pptx_buffer,
                            file_name=f"SLIDE_AUTOMATIK_{docx_file.name.replace('.docx', '.pptx')}",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                        
                        # Provide a preview of the extracted data
                        with st.expander("Lihat Pratonton Data yang Diekstrak"):
                            st.dataframe(parsed_data)
                            
            except Exception as e:
                st.error(f"Ralat berlaku semasa pemprosesan: {e}")
