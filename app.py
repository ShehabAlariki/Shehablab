import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import arabic_reshaper
from bidi.algorithm import get_display
import io
import zipfile

def format_arabic(text):
    if pd.isna(text) or text == "":
        return ""
    text = str(text)
    reshaped_text = arabic_reshaper.reshape(text)
    return get_display(reshaped_text)

st.title("نظام توليد استمارات الترحيل 📄")

# رفع ملف الإكسل
uploaded_file = st.file_uploader("قم برفع ملف الإكسل هنا", type=["xlsx"])

if uploaded_file is not None:
    if st.button("توليد الاستمارات"):
        try:
            df = pd.read_excel(uploaded_file, skiprows=2)
            df = df.dropna(how='all')
            
            template_pdf = 'استمارة ترحيل بضاعة .pdf'
            font_path = "arial.ttf"
            
            COORDS = {
                'رقم_الاستمارة': (120, 140), 'التاريخ': (470, 185), 'رقم_اللوحة': (180, 185),
                'اسم_السائق': (470, 205), 'رقم_الجوال': (470, 225), 'الميناء': (470, 245),
                'اسم_الصنف': (450, 350), 'الكمية_رقم': (290, 350), 'الكمية_وحدة': (210, 350),
                'نوع_البضاعة_تعهد': (350, 715), 'الكمية_تعهد': (250, 715),
                'اسم_السائق_تعهد': (250, 770), 'رقم_اللوحة_تعهد': (120, 770)
            }

            # إنشاء ملف Zip في الذاكرة لتجميع الاستمارات
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in df.iterrows():
                    doc = fitz.open(template_pdf)
                    page = doc[0]
                    
                    form_num = str(row['رقم الاستمارة']).replace('.0', '')
                    date_val = str(row['التاريخ']).split()[0]
                    driver_name = row['اسم السائق']
                    qty_full = str(row['الكمية']).split()
                    qty_num = qty_full[0] if len(qty_full) > 0 else ""
                    qty_unit = qty_full[1] if len(qty_full) > 1 else ""
                    
                    try:
                        page.insert_font(fontname="arab", fontfile=font_path)
                        font_name = "arab"
                    except:
                        font_name = "helv"
                    
                    def add_text(text, key):
                        if text == "": return
                        page.insert_text(COORDS[key], format_arabic(text), fontname=font_name, fontsize=11, color=(1, 0, 0))

                    add_text(form_num, 'رقم_الاستمارة')
                    add_text(date_val, 'التاريخ')
                    add_text(driver_name, 'اسم_السائق')
                    add_text(str(row['رقم الجوال']).replace('.0', ''), 'رقم_الجوال')
                    add_text(row['الميناء'], 'الميناء')
                    add_text(row['رقم السيارة'], 'رقم_اللوحة')
                    add_text(row['نوع البضاعة'], 'اسم_الصنف')
                    add_text(qty_num, 'الكمية_رقم')
                    add_text(qty_unit, 'الكمية_وحدة')
                    add_text(row['نوع البضاعة'], 'نوع_البضاعة_تعهد')
                    add_text(f"{qty_num} {qty_unit}", 'الكمية_تعهد')
                    add_text(driver_name, 'اسم_السائق_تعهد')
                    add_text(row['رقم السيارة'], 'رقم_اللوحة_تعهد')
                    
                    # حفظ الـ PDF في الذاكرة ثم إضافته لملف الـ Zip
                    pdf_bytes = doc.write()
                    doc.close()
                    zip_file.writestr(f"استمارة_{form_num}_{driver_name}.pdf", pdf_bytes)

            st.success("تم توليد جميع الاستمارات بنجاح! 🎉")
            
            # توفير زر لتحميل الملف المضغوط
            st.download_button(
                label="تحميل الاستمارات (ملف Zip)",
                data=zip_buffer.getvalue(),
                file_name="الاستمارات_الجاهزة.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"حدث خطأ أثناء المعالجة: {e}")
