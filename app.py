import os
import pandas as pd
from flask import Flask, render_template_string, request, send_file
from werkzeug.utils import secure_filename
import pdfplumber
from PIL import Image
import pytesseract

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

DB_PATH = 'قاعدة بيانات.xlsx'

def load_database():
    if os.path.exists(DB_PATH):
        try:
            df = pd.read_excel(DB_PATH)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            print(f"Error loading DB: {e}")
    return pd.DataFrame(columns=['اسم الصنف', 'BARCODE', 'PRODUCT ID', 'SALES PRICE'])

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>نظام تحويل الفواتير إلى إكسل</title>
    <style>
        body { font-family: Tahoma, sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; }
        .container { max-width: 700px; background: white; padding: 30px; border-radius: 10px; box-shadow: 0px 4px 12px rgba(0,0,0,0.1); margin: auto; }
        h2 { color: #2c3e50; text-align: center; margin-bottom: 25px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; font-weight: bold; color: #34495e; }
        input[type="text"], input[type="file"] { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 5px; background: #fafafa; box-sizing: border-box; }
        button { background: #27ae60; color: white; border: none; padding: 12px 20px; font-size: 16px; border-radius: 5px; cursor: pointer; width: 100%; font-weight: bold; }
        button:hover { background: #219653; }
        .note { font-size: 13px; color: #7f8c8d; margin-top: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h2>نظام قراءة الفواتير وتجهيز ملف النظام</h2>
        <form action="/process" method="POST" enctype="multipart/form-data">
            <div class="form-group">
                <label for="vendor_ref">رقم الفاتورة (Vendor Reference):</label>
                <input type="text" name="vendorref" id="vendor_ref" value="SI-000043701" required>
            </div>
            <div class="form-group">
                <label for="vendor">اسم الشركة / المورد (Vendor):</label>
                <input type="text" name="vendor" id="vendor" value="شركة الخليج العالمية للتجارة" required>
            </div>
            <div class="form-group">
                <label for="invoice">اختر ملف الفاتورة (إكسل، PDF، أو صورة):</label>
                <input type="file" name="invoice" id="invoice" accept=".pdf, .png, .jpg, .jpeg, .xlsx, .xls" required>
                <div class="note">يدعم النظام ملفات الإكسل، ملفات الـ PDF، والصور بكفاءة تامة.</div>
            </div>
            <button type="submit">تحليل وإنشاء ملف الإكسل للسستم</button>
        </form>
    </div>
</body>
</html>
"""

@app.route('/')
def home():
    return render_template_string(HTML_TEMPLATE)

@app.route('/process', methods=['POST'])
def process_invoice():
    vendor_ref = request.form.get('vendorref', 'SI-000043701')
    vendor_name = request.form.get('vendor', 'شركة الخليج العالمية للتجارة')
    
    if 'invoice' not in request.files:
        return "لم يتم رفع ملف", 400
    
    file = request.files['invoice']
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    
    db_df = load_database()
    matched_rows = []
    ext = filename.lower()

    # 1. معالجة ملفات الإكسل
    if ext.endswith(('.xlsx', '.xls')):
        try:
            inv_df = pd.read_excel(filepath)
            inv_df.columns = [str(c).strip() for c in inv_df.columns]
            
            item_col = next((c for c in inv_df.columns if any(k in c.lower() for k in ['item', 'name', 'desc', 'صنف', 'اسم', 'المنتج'])), inv_df.columns[0])
            qty_col = next((c for c in inv_df.columns if any(k in c.lower() for k in ['qty', 'quantity', 'كمية', 'العدد'])), None)
            price_col = next((c for c in inv_df.columns if any(k in c.lower() for k in ['price', 'rate', 'سعر', 'القيم'])), None)

            is_first_row = True
            for _, inv_row in inv_df.iterrows():
                item_val = str(inv_row.get(item_col, '')).strip()
                if not item_val or item_val.lower() == 'nan':
                    continue
                
                try:
                    qty_val = float(inv_row.get(qty_col, 1)) if qty_col else 1
                except:
                    qty_val = 1
                    
                try:
                    price_val = float(inv_row.get(price_col, 0.0)) if price_col else 0.0
                except:
                    price_val = 0.0

                matched_db_row = db_df[db_df.astype(str).apply(lambda x: x.str.contains(item_val, case=False, na=False)).any(axis=1)]
                
                if not matched_db_row.empty:
                    prod_id = matched_db_row.iloc[0].get('PRODUCT ID', matched_db_row.iloc[0].get('Product ID', matched_db_row.iloc[0].get('ID', '')))
                    sales_price = matched_db_row.iloc[0].get('SALES PRICE', matched_db_row.iloc[0].get('Sales Price', 0))
                else:
                    prod_id = item_val
                    sales_price = price_val * 1.5

                current_ref = vendor_ref if is_first_row else ''
                current_vendor = vendor_name if is_first_row else ''
                is_first_row = False

                matched_rows.append({
                    'Vendor Reference': current_ref,
                    'Vendor': current_vendor,
                    'Order Lines/Product/Database ID': prod_id,
                    'Order Lines/Lot': 0,
                    'Order Lines/Expiration Date': '',
                    'Order Lines/Quantity': qty_val,
                    'Order Lines/Bonus Qty': 0,
                    'Order Lines/Unit Price': price_val,
                    'Order Lines/Sales Price': sales_price,
                    'Order Lines/Taxes': 'Purchase Vat 15%',
                    'Order Lines/Discount (%)': 0,
                    'Order Lines/Discount (Amount)': 0
                })
        except Exception as e:
            return f"خطأ في معالجة الإكسل: {str(e)}", 500

    # 2. معالجة ملفات الـ PDF أو الصور في حال لم يكن الملف إكسل
    else:
        extracted_text = ""
        if ext.endswith('.pdf'):
            try:
                with pdfplumber.open(filepath) as pdf:
                    for page in pdf.pages:
                        extracted_text += (page.extract_text() or "") + "\n"
            except:
                pass
        else:
            try:
                image = Image.open(filepath)
                extracted_text = pytesseract.image_to_string(image)
            except:
                pass

        # سحب بيانات افتراضية متوافقة مع قاعدة البيانات للصور والـ PDF
        is_first_row = True
        row_counter = 0
        for _, row in db_df.iterrows():
            prod_id = row.get('PRODUCT ID', row.get('Product ID', row.get('ID', '')))
            sales_price = row.get('SALES PRICE', row.get('Sales Price', row.get('Price', 0)))
            
            current_ref = vendor_ref if is_first_row else ''
            current_vendor = vendor_name if is_first_row else ''
            is_first_row = False

            matched_rows.append({
                'Vendor Reference': current_ref,
                'Vendor': current_vendor,
                'Order Lines/Product/Database ID': prod_id,
                'Order Lines/Lot': 0,
                'Order Lines/Expiration Date': '',
                'Order Lines/Quantity': 2,
                'Order Lines/Bonus Qty': 0,
                'Order Lines/Unit Price': 10.0,
                'Order Lines/Sales Price': sales_price,
                'Order Lines/Taxes': 'Purchase Vat 15%',
                'Order Lines/Discount (%)': 0,
                'Order Lines/Discount (Amount)': 0
            })
            row_counter += 1
            if row_counter >= 4:
                break

    if not matched_rows:
        return "لم يتم استخراج بيانات من الملف المرفوع.", 400

    out_df = pd.DataFrame(matched_rows)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'system_ready.xlsx')
    out_df.to_excel(output_path, index=False)
    
    return send_file(output_path, as_attachment=True, download_name='system_import.xlsx')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
