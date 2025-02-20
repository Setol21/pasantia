import shutil, os, pptx
import pandas as pd
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


def read_and_get() -> None:
    my_data = pd.read_excel("Planilla de Pagos MX.xlsx",sheet_name="Detalle OCs",skiprows=1)
    
    if not os.path.exists("ppts"):
        for x in my_data["Retail"].unique():
            for y in my_data["Descripción"].unique():
                os.makedirs(f"ppts\\{x}\\{y}")

    for i in range(len(my_data["Descripción"])):
        create_ppt(my_data["Descripción"][i],my_data["Retail"][i],int(my_data["Comprometido Retail"][i]),my_data["OC"][i])

def create_ppt(type_name: str, retail_name_specific: str, amount: int ,OC_number_specific: str)-> None:
    amount_formatted = f"{amount:,}"

    # Creating the presentation
    prs = pptx.Presentation()

    # Creating the title
    title_slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    title.text = "EVIDENCIAS"

    # Creating the RETAIL box centered
    top = height = Inches(2) 
    width = Inches(5)
    left = (prs.slide_width - width)/2
    retail_box = slide.shapes.add_textbox(left, top, width, height)

    # Textframe inside the box
    tf_retail_name = retail_box.text_frame

    # Adding the paragraph
    retail_name = tf_retail_name.paragraphs[0]
    retail_name.text = retail_name_specific.upper()
    retail_name.font.bold = True
    retail_name.alignment = PP_ALIGN.CENTER

    run = retail_name.runs[0]
    run.font.size = Pt(72)
    run.font.color.rgb = RGBColor(128,0,64)


    # Add the OC number
    oc_number = tf_retail_name.add_paragraph()
    oc_number.text = OC_number_specific
    oc_number.alignment = PP_ALIGN.CENTER
    oc_number.font.size = Pt(48)

    # Add the Period
    period = tf_retail_name.add_paragraph()
    period.text = "Oct 1st to Dec 31st 2024"
    period.alignment = PP_ALIGN.CENTER
    period.font.size = Pt(28)

    prs.save(f"ppts\\{retail_name_specific}\\{type_name}\\{type_name.upper()} {retail_name_specific.upper()} (${amount_formatted.replace(',','.')}) - {OC_number_specific}.pptx")
    return

read_and_get()