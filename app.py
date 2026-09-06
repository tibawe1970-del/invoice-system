import os
import pandas as pd
from flask import Flask, render_template_string, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

DB_PATH = 'قاعدة بيانات.xlsx'

def load_database():
    if os.path.exists(DB_PATH):
        try:
            df = pd.read_excel(DB_PATH)
            # تنظيف أسماء الأعمدة
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            print(f"Error loading DB: {e}")
    return pd.DataFrame()

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>مطابقة أسماء الفاتورة بقاعدة البيانات</title>
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
        <h2>ربط اسم الفاتورة بـ Product ID من قاعدة البيانات</h2>
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
                <label for="invoice">اختر ملف الفاتورة (إكسل):</label>
                <input type="file" name="invoice" id="invoice" accept=".xlsx, .xls" required>
                <div class="note">سيتم مطابقة اسم الصنف من الفاتورة مع أسماء قاعدة البيانات واستخراج الـ ID الخاص به.</div>
            </div>
            <button type="submit">تحليل وإنشاء ملف الإكسل</button>
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

    if ext.endswith(('.xlsx', '.xls')):
        try:
            raw_df = pd.read_excel(filepath, header=None)
            header_row_idx = None
            for idx, row in raw_df.iterrows():
                row_str = " ".join([str(val) for val in row.values if pd.notna(val)])
                if 'المواصفات' in row_str or 'DESCRIPTION' in row_str or 'رقم الصنف' in row_str:
                    header_row_idx = idx
                    break
            
            if header_row_idx is not None:
                inv_df = pd.read_excel(filepath, skiprows=header_row_idx)
            else:
                inv_df = pd.read_excel(filepath)
                
            inv_df.columns = [str(c).strip() for c in inv_df.columns]
            cols = list(inv_df.columns)

            # تحديد عمود وصف الصنف في الفاتورة (غالباً العمود الثالث أو الثاني)
            desc_col = next((c for c in cols if any(k in c.lower() for k in ['المواصفات', 'description', 'name'])), cols[2] if len(cols) > 2 else cols[1])
            qty_col = next((c for c in cols if any(k in c.lower() for k in ['الكمية', 'qty', 'quantity'])), cols[4] if len(cols) > 4 else None)
            price_col = next((c for c in cols if any(k in c.lower() for k in ['السعر', 'price', 'unit price'])), cols[5] if len(cols) > 5 else None)

            is_first_row = True
            for _, inv_row in inv_df.iterrows():
                desc_raw = str(inv_row.get(desc_col, '')).strip()
                
                # استبعاد الصفوف الفارغة أو ترويسة الجدول
                if not desc_raw or desc_raw.lower() in ['nan', 'none', '0', 'م', 'item no.']:
                    continue
                if any(k in desc_raw for k in ['شركة', 'رقم الفاتورة', 'تاريخ', 'الرقم الضريبي', 'المندوب', 'المواصفات']):
                    continue

                # استخراج اسم الصنف الأساسي (مثل الإنجليزي أولاً أو الكلمات المفتاحية)
                invoice_item_name = desc_raw.split('\n')[0].strip()

                # استخراج الكمية والسعر
                qty_val = 1
                if qty_col and pd.notna(inv_row.get(qty_col)):
                    try:
                        qty_val = float(inv_row[qty_col])
                    except:
                        qty_val = 1

                price_val = 0.0
                if price_col and pd.notna(inv_row.get(price_col)):
                    try:
                        price_val = float(inv_row[price_col])
                    except:
                        price_val = 0.0

                sales_price = price_val * 1.5 if price_val > 0 else 0.0
                
                # القيمة الافتراضية إذا لم يتم العثور على مطابقة
                product_id_to_use = invoice_item_name

                # المطابقة مع ملف قاعدة البيانات
                if not db_df.empty:
                    db_cols = list(db_df.columns)
                    # نفترض أن العمود الأول (أو الثاني) هو اسم الصنف، والعمود الآخر هو الـ Product ID
                    name_col_db = db_cols[1] if len(db_cols) > 1 else db_cols[0] # عمود الاسم في القاعدة
                    id_col_db = db_cols[0] if len(db_cols) > 1 else db_cols[0]   # عمود الـ ID في القاعدة (العمود الأول أو الثاني)

                    # البحث بالاسم الجزئي أو التام
                    matched = db_df[db_df[name_col_db].astype(str).str.contains(invoice_item_name[:15], case=False, na=False)]
                    if matched.empty:
                        # محاولة بحث أوسع بأول كلمتين
                        short_name = " ".join(invoice_item_name.split()[:2])
                        matched = db_df[db_df[name_col_db].astype(str).str.contains(short_name, case=False, na=False)]

                    if not matched.empty:
                        # جلب الـ ID من العمود المحدد في قاعدة البيانات
                        product_id_to_use = matched.iloc[0][id_col_db]

                current_ref = vendor_ref if is_first_row else ''
                current_vendor = vendor_name if is_first_row else ''
                is_first_row = False

                matched_rows.append({
                    'Vendor Reference': current_ref,
                    'Vendor': current_vendor,
                    'Order Lines/Product/Database ID': product_id_to_use,
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
            return f"خطأ أثناء المعالجة: {str(e)}", 500

    if not matched_rows:
        return "لم يتم العثور على أصناف.", 400

    out_df = pd.DataFrame(matched_rows)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'system_ready.xlsx')
    out_df.to_excel(output_path, index=False)
    
    return send_file(output_path, as_attachment=True, download_name='system_import.xlsx')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
