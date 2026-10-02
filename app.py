import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import arabic_reshaper
from bidi.algorithm import get_display
import io
import zipfile

# دالة معالجة النصوص العربية
def format_arabic(text):
    if pd.isna(text) or text == "" or str(text).lower() == "nan":
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
            # قراءة الملف وتخطي أول صفين (عناوين الجدول في الصف الثالث)
            df = pd.read_excel(uploaded_file, skiprows=2)
            
            # تنظيف أسماء الأعمدة من المسافات الزائدة (مهم جداً لتجنب خطأ عدم العثور على العمود)
            df.columns = df.columns.str.strip()
            
            # إزالة الصفوف الفارغة بالكامل
            df = df.dropna(how='all')
            
            # تأكد من أن اسم القالب مطابق لاسم الملف المرفوع في GitHub
            template_pdf = 'استمارة ترحيل بضاعة .pdf'
            font_path = "arial.ttf"
            
            # الإحداثيات المحدثة لتتناسب مع الملف عالي الدقة (2379x3362)
            COORDS = {
                'رقم_الاستمارة': (220, 440),
                'التاريخ': (1650, 630), 
                'رقم_اللوحة': (750, 630),
                'اسم_السائق': (1550, 710), 
                'رقم_الجوال': (1650, 785), 
                'الميناء': (1650, 860),
                
                # بيانات الجدول
                'اسم_الصنف': (1600, 1190), 
                'الكمية_رقم': (1250, 1190), 
                'الكمية_وحدة': (1000, 1190),
                
                # بيانات التعهد في أسفل الصفحة
                'الكمية_تعهد': (1450, 2370),
                'نوع_البضاعة_تعهد': (950, 2370), 
                'اسم_السائق_تعهد': (1000, 2590), 
                'رقم_اللوحة_تعهد': (350, 2590)
            }

            # إنشاء ملف Zip في الذاكرة لتجميع الاستمارات
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for index, row in df.iterrows():
                    doc = fitz.open(template_pdf)
                    page = doc[0]
                    
                    # استخراج البيانات بأمان
                    form_num = str(row.get('رقم الاستمارة', '')).replace('.0', '')
                    date_val = str(row.get('التاريخ', '')).split()[0]
                    driver_name = str(row.get('اسم السائق', ''))
                    
                    # فصل الكمية والوحدة
                    qty_full = str(row.get('الكمية', '')).split()
                    qty_num = qty_full[0] if len(qty_full) > 0 else ""
                    qty_unit = qty_full[1] if len(qty_full) > 1 else ""
                    
                    # تحميل الخط العربي
                    try:
                        page.insert_font(fontname="arab", fontfile=font_path)
                        font_name = "arab"
                    except:
                        font_name = "helv"
                    
                    # دالة طباعة النصوص باستخدام حجم الخط 45
                    def add_text(text, key):
                        if text == "" or pd.isna(text): return
                        page.insert_text(COORDS[key], format_arabic(text), fontname=font_name, fontsize=45, color=(1, 0, 0))

                    # تعبئة الحقول
                    add_text(form_num, 'رقم_الاستمارة')
                    add_text(date_val, 'التاريخ')
                    add_text(driver_name, 'اسم_السائق')
                    add_text(str(row.get('رقم الجوال', '')).replace('.0', ''), 'رقم_الجوال')
                    add_text(str(row.get('الميناء', '')), 'الميناء')
                    add_text(str(row.get('رقم السيارة', '')), 'رقم_اللوحة')
                    
                    add_text(str(row.get('نوع البضاعة', '')), 'اسم_الصنف')
                    add_text(qty_num, 'الكمية_رقم')
                    add_text(qty_unit, 'الكمية_وحدة')
                    
                    add_text(str(row.get('نوع البضاعة', '')), 'نوع_البضاعة_تعهد')
                    add_text(f"{qty_num} {qty_unit}", 'الكمية_تعهد')
                    add_text(driver_name, 'اسم_السائق_تعهد')
                    add_text(str(row.get('رقم السيارة', '')), 'رقم_اللوحة_تعهد')
                    
                    # حفظ الـ PDF في الذاكرة ثم إضافته لملف الـ Zip
                    pdf_bytes = doc.write()
                    doc.close()
                    
                    # تنظيف اسم الملف من الحروف غير الصالحة
                    safe_driver_name = driver_name.replace('/', '-').replace('\\', '-')
                    zip_file.writestr(f"استمارة_{form_num}_{safe_driver_name}.pdf", pdf_bytes)

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
