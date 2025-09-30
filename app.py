from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import pandas as pd
import os
from datetime import datetime
import jdatetime
import requests
import json
import numpy as np

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import timedelta
from math import radians, sin, cos, sqrt, atan2


app = Flask(__name__)
app.secret_key = 'your-secret-key-2024'

# مسیر فایل‌های Excel
USERS_FILE = 'users.xlsx'
CUSTOMERS_FILE = 'customers.xlsx'
VISITS_FILE = 'visits.xlsx'
EXAMS_FILE = 'azmon.xlsx'  # ← این خط را اضافه کنید

def calculate_distance(lat1, lon1, lat2, lon2):
    """محاسبه فاصله بین دو نقطه جغرافیایی به متر (فرمول Haversine)"""
    try:
        # تبدیل به رادیان
        lat1, lon1, lat2, lon2 = map(radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
        
        # فرمول Haversine
        dlat = lat2 - lat1
        dlon = lon2 - lon1
        a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
        c = 2 * atan2(sqrt(a), sqrt(1-a))
        
        # شعاع زمین به متر
        radius = 6371000
        distance = radius * c
        
        return distance
    except Exception as e:
        print(f"خطا در محاسبه فاصله: {e}")
        return None

# تابع کمکی برای تبدیل امن به JSON
def safe_json_response(data):
    """تبدیل امن داده‌ها به JSON response"""
    def convert_numpy_types(obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif pd.isna(obj):
            return None
        return obj
    
    # تبدیل recursive همه مقادیر
    def recursive_convert(data):
        if isinstance(data, dict):
            return {k: recursive_convert(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [recursive_convert(item) for item in data]
        else:
            return convert_numpy_types(data)
    
    clean_data = recursive_convert(data)
    return jsonify(clean_data)

def load_brand_order_from_excel():
    """بارگذاری ترتیب برندها از شیت brand در فایل products.xlsx"""
    try:
        if not os.path.exists('products.xlsx'):
            return None
            
        # بررسی وجود شیت brand
        with pd.ExcelFile('products.xlsx') as xls:
            if 'brand' not in xls.sheet_names:
                return None
                
        df = pd.read_excel('products.xlsx', sheet_name='brand')
        
        # پاک کردن فاصله‌های اضافی
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        # مرتب‌سازی بر اساس Radif
        df = df.sort_values('Radif', ascending=True)
        
        # برگرداندن لیست برندها
        brand_order = df['Brand'].tolist()
        
        print(f"✅ Brand order loaded: {brand_order}")
        return brand_order
        
    except Exception as e:
        print(f"❌ Error loading brand order: {e}")
        return None

def save_brand_order_to_excel(brand_order):
    """ذخیره ترتیب برندها در شیت brand فایل products.xlsx"""
    try:
        # ایجاد DataFrame با ترتیب جدید
        brand_data = []
        for index, brand in enumerate(brand_order):
            brand_data.append({
                'Brand': brand,
                'Radif': index + 1
            })
        
        df = pd.DataFrame(brand_data)
        
        # خواندن فایل موجود
        if os.path.exists('products.xlsx'):
            with pd.ExcelWriter('products.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='brand', index=False)
        else:
            # اگر فایل وجود نداشت، ایجاد کن
            with pd.ExcelWriter('products.xlsx', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='brand', index=False)
        
        print(f"✅ Brand order saved: {brand_order}")
        return True
        
    except Exception as e:
        print(f"❌ Error saving brand order: {e}")
        return False

@app.route('/get_brand_order')
def get_brand_order():
    """دریافت ترتیب برندها"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    try:
        brand_order = load_brand_order_from_excel()
        
        if brand_order:
            return jsonify({
                'success': True,
                'brand_order': brand_order
            })
        else:
            # اگر ترتیب ثبت نشده، لیست خالی برگردان
            return jsonify({
                'success': True,
                'brand_order': []
            })
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/save_brand_order', methods=['POST'])
def save_brand_order():
    """ذخیره ترتیب برندها"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    # بررسی سطح دسترسی (فقط ادمین)
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'Access denied. Admin only.'}), 403
    
    try:
        data = request.get_json()
        brand_order = data.get('brand_order', [])
        
        if not brand_order or not isinstance(brand_order, list):
            return jsonify({'error': 'Invalid brand order data'}), 400
        
        # ذخیره در فایل Excel
        if save_brand_order_to_excel(brand_order):
            return jsonify({
                'success': True,
                'message': 'Brand order saved successfully'
            })
        else:
            return jsonify({'error': 'Failed to save brand order'}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def jalali_to_gregorian(jalali_date_str):
    """تبدیل تاریخ شمسی به میلادی"""
    try:
        # فرمت‌های مختلف شمسی
        if len(jalali_date_str) == 8:  # 14040101
            year = int(jalali_date_str[:4])
            month = int(jalali_date_str[4:6])
            day = int(jalali_date_str[6:8])
        elif '/' in jalali_date_str:  # 1404/01/01
            parts = jalali_date_str.split('/')
            year = int(parts[0])
            month = int(parts[1])
            day = int(parts[2])
        elif '-' in jalali_date_str:  # 1404-01-01
            parts = jalali_date_str.split('-')
            year = int(parts[0])
            month = int(parts[1])
            day = int(parts[2])
        else:
            return None
            
        # تبدیل به میلادی
        jalali_date = jdatetime.date(year, month, day)
        gregorian_date = jalali_date.togregorian()
        return gregorian_date.strftime('%Y-%m-%d')
    except Exception as e:
        print(f"خطا در تبدیل تاریخ {jalali_date_str}: {e}")
        return None

def gregorian_to_jalali(gregorian_date_str):
    """تبدیل تاریخ میلادی به شمسی"""
    try:
        if isinstance(gregorian_date_str, str):
            gregorian_date = datetime.strptime(gregorian_date_str, '%Y-%m-%d').date()
        else:
            gregorian_date = gregorian_date_str
            
        jalali_date = jdatetime.date.fromgregorian(date=gregorian_date)
        return jalali_date.strftime('%Y/%m/%d')
    except Exception as e:
        print(f"خطا در تبدیل تاریخ {gregorian_date_str}: {e}")
        return gregorian_date_str

def jalali_date_compact(gregorian_date_str):
    """تبدیل تاریخ میلادی به شمسی فشرده (14040101)"""
    try:
        if isinstance(gregorian_date_str, str):
            gregorian_date = datetime.strptime(gregorian_date_str, '%Y-%m-%d').date()
        else:
            gregorian_date = gregorian_date_str
            
        jalali_date = jdatetime.date.fromgregorian(date=gregorian_date)
        return jalali_date.strftime('%Y%m%d')
    except Exception as e:
        print(f"خطا در تبدیل تاریخ {gregorian_date_str}: {e}")
        return gregorian_date_str

def load_users_from_excel():
    """بارگذاری کاربران از فایل Excel"""
    try:
        if not os.path.exists(USERS_FILE):
            print("❌ Users file not found:", USERS_FILE)
            return None
            
        df = pd.read_excel(USERS_FILE, sheet_name='users')
        print("✅ Users file loaded successfully")
        
        # پاک کردن فاصله‌های اضافی
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        return df
    except Exception as e:
        print("❌ Error loading users file:", e)
        return None

def load_customers_from_excel():
    """بارگذاری مشتریان از فایل Excel - اصلاح شده برای خطای NaN"""
    try:
        if not os.path.exists(CUSTOMERS_FILE):
            print("❌ Customers file not found:", CUSTOMERS_FILE)
            return None
            
        df = pd.read_excel(CUSTOMERS_FILE, sheet_name='customers')
        print("✅ Customers file loaded successfully")
        
        # پاک کردن فاصله‌های اضافی
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        # 🔧 FIX: تبدیل مقادیر NaN به مقادیر قابل استفاده
        # اگر ستون LocationSet وجود داره، NaN ها رو به False تبدیل کن
        if 'LocationSet' in df.columns:
            df['LocationSet'] = df['LocationSet'].fillna(False)
            # تبدیل string values به boolean
            df['LocationSet'] = df['LocationSet'].apply(lambda x: 
                True if str(x).lower() in ['true', '1', 'yes', 'بله'] else False
            )
        
        # اگر ستون‌های Latitude/Longitude وجود دارن، NaN ها رو به 0 تبدیل کن
        if 'Latitude' in df.columns:
            df['Latitude'] = df['Latitude'].fillna(0)
        
        if 'Longitude' in df.columns:
            df['Longitude'] = df['Longitude'].fillna(0)
        
        print(f"📊 Customers data cleaned: {len(df)} records")
        return df
        
    except Exception as e:
        print("❌ Error loading customers file:", e)
        return None

def load_visits_from_excel():
    """بارگذاری مراجعات از فایل Excel"""
    try:
        if not os.path.exists(VISITS_FILE):
            print("❌ Visits file not found:", VISITS_FILE)
            return None
            
        df = pd.read_excel(VISITS_FILE, sheet_name='visits')
        print("✅ Visits file loaded successfully")
        
        return df
    except Exception as e:
        print("❌ Error loading visits file:", e)
        return None

def save_customers_to_excel(df):
    """ذخیره مشتریان در فایل Excel"""
    try:
        df.to_excel(CUSTOMERS_FILE, sheet_name='customers', index=False)
        print("✅ Customers file saved successfully")
        return True
    except Exception as e:
        print("❌ Error saving customers file:", e)
        return False

def save_visits_to_excel(df):
    """ذخیره مراجعات در فایل Excel"""
    try:
        df.to_excel(VISITS_FILE, sheet_name='visits', index=False)
        print("✅ Visits file saved successfully")
        return True
    except Exception as e:
        print("❌ Error saving visits file:", e)
        return False

def authenticate_user(username, password):
    """احراز هویت کاربر"""
    try:
        users_df = load_users_from_excel()
        if users_df is None:
            print("❌ Cannot load users for authentication")
            return None
        
        print(f"🔍 Looking for user: '{username}'")
        
        username = str(username).strip()
        password = str(password).strip()
        
        user = users_df[users_df['Userv'].astype(str).str.strip() == username]
        
        if not user.empty:
            stored_password = str(user.iloc[0]['Passv']).strip()
            
            if stored_password == password:
                print("✅ Authentication successful!")
                return {
                    'Codev': str(user.iloc[0]['Codev']).strip(),
                    'Namev': str(user.iloc[0]['Namev']).strip(),
                    'Userv': str(user.iloc[0]['Userv']).strip(),
                    'Typev': str(user.iloc[0]['Typev']).strip()
                }
        
        print("❌ Authentication failed")
        return None
    except Exception as e:
        print(f"❌ Authentication error: {e}")
        return None

@app.route('/')
def index():
    """صفحه اصلی"""
    if 'user_id' in session:
        return render_template('dashboard.html', user=session['user_info'])
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    """صفحه ورود"""
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        user = authenticate_user(username, password)
        if user:
            session['user_id'] = user['Codev']
            session['user_info'] = user
            flash('ورود موفقیت‌آمیز بود!', 'success')
            return redirect(url_for('index'))
        else:
            flash('نام کاربری یا رمز عبور اشتباه است!', 'error')
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    """خروج از حساب کاربری"""
    session.pop('user_id', None)
    session.pop('user_info', None)
    flash('با موفقیت خارج شدید!', 'info')
    return redirect(url_for('login'))

@app.route('/profile')
def profile():
    """صفحه پروفایل کاربر"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    return render_template('profile.html', user=session['user_info'])

@app.route('/users')
def users_list():
    """لیست کاربران (فقط برای ادمین)"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    users_df = load_users_from_excel()
    if users_df is not None:
        users = users_df.to_dict('records')
        return render_template('users.html', users=users)
    else:
        flash('خطا در بارگذاری لیست کاربران!', 'error')
        return redirect(url_for('index'))

@app.route('/customers')
def customers_list():
    """لیست مشتریان بازاریاب"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # بارگذاری مشتریان
    customers_df = load_customers_from_excel()
    if customers_df is None:
        flash('خطا در بارگذاری لیست مشتریان!', 'error')
        return redirect(url_for('index'))
    
    # فیلتر کردن مشتریان بر اساس کد بازاریاب
    bazaryab_code = session['user_info']['Codev']
    my_customers = customers_df[customers_df['BazaryabCode'] == bazaryab_code]
    
    customers = my_customers.to_dict('records')
    
    return render_template('customers.html', customers=customers, user=session['user_info'])

@app.route('/set_location/<customer_code>')
def set_location(customer_code):
    """صفحه تنظیم مکان مشتری"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    customers_df = load_customers_from_excel()
    if customers_df is None:
        flash('خطا در بارگذاری اطلاعات!', 'error')
        return redirect(url_for('customers_list'))
    
    # پیدا کردن مشتری
    customer = customers_df[customers_df['CustomerCode'] == customer_code]
    if customer.empty:
        flash('مشتری یافت نشد!', 'error')
        return redirect(url_for('customers_list'))
    
    customer_info = customer.iloc[0].to_dict()
    
    return render_template('set_location.html', customer=customer_info, user=session['user_info'])

@app.route('/save_location', methods=['POST'])
def save_location():
    """ذخیره مکان مشتری"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    customer_code = request.form['customer_code']
    latitude = request.form['latitude']
    longitude = request.form['longitude']
    
    customers_df = load_customers_from_excel()
    if customers_df is None:
        flash('خطا در بارگذاری اطلاعات!', 'error')
        return redirect(url_for('customers_list'))
    
    # بروزرسانی مکان مشتری
    customers_df.loc[customers_df['CustomerCode'] == customer_code, 'Latitude'] = latitude
    customers_df.loc[customers_df['CustomerCode'] == customer_code, 'Longitude'] = longitude
    customers_df.loc[customers_df['CustomerCode'] == customer_code, 'LocationSet'] = True
    
    # ذخیره فایل
    if save_customers_to_excel(customers_df):
        flash('مکان مشتری با موفقیت ثبت شد!', 'success')
    else:
        flash('خطا در ذخیره اطلاعات!', 'error')
    
    return redirect(url_for('customers_list'))

@app.route('/record_visit', methods=['POST'])
def record_visit():
    """ثبت مراجعه به مشتری با بررسی موقعیت جغرافیایی"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    try:
        # دریافت اطلاعات از request
        data = request.get_json() if request.is_json else request.form
        
        customer_code = data.get('customer_code')
        current_lat = data.get('current_latitude')
        current_lon = data.get('current_longitude')
        
        if not customer_code:
            return jsonify({'error': 'کد مشتری الزامی است'}), 400
        
        if not current_lat or not current_lon:
            return jsonify({'error': 'موقعیت جغرافیایی فعلی شما دریافت نشد'}), 400
        
        # بارگذاری اطلاعات مشتری
        customers_df = load_customers_from_excel()
        if customers_df is None:
            return jsonify({'error': 'خطا در بارگذاری اطلاعات'}), 500
        
        # پیدا کردن مشتری
        customer = customers_df[customers_df['CustomerCode'] == customer_code]
        
        if customer.empty:
            return jsonify({'error': 'مشتری یافت نشد'}), 404
        
        customer_info = customer.iloc[0]
        
        # بررسی اینکه آیا موقعیت مشتری ثبت شده است
        if not customer_info.get('LocationSet') or not customer_info.get('Latitude') or not customer_info.get('Longitude'):
            return jsonify({'error': 'موقعیت جغرافیایی این مشتری ثبت نشده است'}), 400
        
        customer_lat = float(customer_info['Latitude'])
        customer_lon = float(customer_info['Longitude'])
        
        # محاسبه فاصله
        distance = calculate_distance(current_lat, current_lon, customer_lat, customer_lon)
        
        if distance is None:
            return jsonify({'error': 'خطا در محاسبه فاصله'}), 500
        
        print(f"🔍 فاصله محاسبه شده: {distance:.2f} متر")
        
        # بررسی فاصله (5 متر)
        MAX_DISTANCE = 5.0
        
        if distance > MAX_DISTANCE:
            return jsonify({
                'error': f'شما در موقعیت مشتری نیستید! فاصله: {distance:.1f} متر',
                'distance': round(distance, 1),
                'max_distance': MAX_DISTANCE,
                'too_far': True
            }), 403
        
        # اگر فاصله مجاز باشد، مراجعه را ثبت کنید
        visits_df = load_visits_from_excel()
        if visits_df is None:
            visits_df = pd.DataFrame(columns=['VisitCode', 'BazaryabCode', 'CustomerCode', 'VisitDate', 'VisitTime', 'Latitude', 'Longitude', 'Distance'])
        
        # ایجاد کد مراجعه جدید
        visit_count = len(visits_df) + 1
        visit_code = f"V{visit_count:03d}"
        
        # ایجاد رکورد جدید
        now = datetime.now()
        new_visit = {
            'VisitCode': visit_code,
            'BazaryabCode': session['user_info']['Codev'],
            'CustomerCode': customer_code,
            'VisitDate': now.strftime('%Y-%m-%d'),
            'VisitTime': now.strftime('%H:%M'),
            'Latitude': current_lat,
            'Longitude': current_lon,
            'Distance': round(distance, 2)
        }
        
        # اضافه کردن به DataFrame
        visits_df = pd.concat([visits_df, pd.DataFrame([new_visit])], ignore_index=True)
        
        # ذخیره فایل
        if save_visits_to_excel(visits_df):
            return jsonify({
                'success': True,
                'message': f'مراجعه با موفقیت ثبت شد (فاصله: {distance:.1f} متر)',
                'distance': round(distance, 1),
                'visit_code': visit_code
            }), 200
        else:
            return jsonify({'error': 'خطا در ثبت مراجعه'}), 500
            
    except Exception as e:
        print(f"❌ خطا در ثبت مراجعه: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500
        
@app.route('/show_map/<customer_code>')
def show_map(customer_code):
    """نمایش مکان مشتری روی نقشه"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    customers_df = load_customers_from_excel()
    if customers_df is None:
        flash('خطا در بارگذاری اطلاعات!', 'error')
        return redirect(url_for('customers_list'))
    
    # پیدا کردن مشتری
    customer = customers_df[customers_df['CustomerCode'] == customer_code]
    if customer.empty:
        flash('مشتری یافت نشد!', 'error')
        return redirect(url_for('customers_list'))
    
    customer_info = customer.iloc[0].to_dict()
    
    # بررسی وجود مختصات
    if not customer_info['Latitude'] or not customer_info['Longitude']:
        flash('مکان این مشتری هنوز ثبت نشده است!', 'error')
        return redirect(url_for('customers_list'))
    
    return render_template('map_view.html', customer=customer_info, user=session['user_info'])

@app.route('/customer_report/<customer_code>')
def customer_report(customer_code):
    """گزارش مراجعات یک مشتری"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # بارگذاری اطلاعات مشتری
    customers_df = load_customers_from_excel()
    if customers_df is None:
        flash('خطا در بارگذاری اطلاعات مشتری!', 'error')
        return redirect(url_for('customers_list'))
    
    # پیدا کردن مشتری
    customer = customers_df[customers_df['CustomerCode'] == customer_code]
    if customer.empty:
        flash('مشتری یافت نشد!', 'error')
        return redirect(url_for('customers_list'))
    
    customer_info = customer.iloc[0].to_dict()
    
    # بررسی دسترسی (فقط بازاریاب همین مشتری یا ادمین)
    if (session['user_info']['Typev'] != 'admin' and 
        customer_info['BazaryabCode'] != session['user_info']['Codev']):
        flash('شما اجازه مشاهده گزارش این مشتری را ندارید!', 'error')
        return redirect(url_for('customers_list'))
    
    # بارگذاری مراجعات
    visits_df = load_visits_from_excel()
    if visits_df is None:
        customer_visits = []
    else:
        # فیلتر مراجعات این مشتری
        customer_visits = visits_df[visits_df['CustomerCode'] == customer_code]
        customer_visits = customer_visits.sort_values('VisitDate', ascending=False)
        customer_visits = customer_visits.to_dict('records')
    
    # بارگذاری اطلاعات بازاریاب
    users_df = load_users_from_excel()
    bazaryab_name = "نامشخص"
    if users_df is not None:
        bazaryab = users_df[users_df['Codev'] == customer_info['BazaryabCode']]
        if not bazaryab.empty:
            bazaryab_name = bazaryab.iloc[0]['Namev']
    
    # آمار کلی
    total_visits = len(customer_visits)
    last_visit = customer_visits[0] if customer_visits else None
    
    return render_template('customer_report.html', 
                         customer=customer_info,
                         visits=customer_visits,
                         bazaryab_name=bazaryab_name,
                         total_visits=total_visits,
                         last_visit=last_visit,
                         user=session['user_info'])

def load_products_from_excel():
    """بارگذاری کالاها از فایل Excel - اصلاح شده برای خطای NaN"""
    try:
        if not os.path.exists('products.xlsx'):
            print("❌ Products file not found!")
            return None
            
        df = pd.read_excel('products.xlsx', sheet_name='products')
        print("✅ Products file loaded successfully")
        
        # پاک کردن فاصله‌های اضافی
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        # 🔧 FIX: تبدیل مقادیر NaN به مقادیر قابل استفاده
        # برای ستون‌های متنی: NaN -> ""
        text_columns = ['ProductCode', 'ProductName', 'Brand', 'Category', 'ImageFile', 'Description']
        for col in text_columns:
            if col in df.columns:
                df[col] = df[col].fillna('')
        
        # برای ستون‌های عددی: NaN -> 0
        numeric_columns = ['Price', 'Stock']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].fillna(0)
        
        # برای ستون‌های offer: NaN -> ""
        offer_columns = ['Offer1', 'Offer2', 'Offer3']
        for col in offer_columns:
            if col in df.columns:
                df[col] = df[col].fillna('')
        
        print(f"📊 Products data cleaned: {len(df)} records")
        return df
        
    except Exception as e:
        print("❌ Error loading products file:", e)
        return None

def load_sales_from_excel():
    """بارگذاری فروش از فایل Excel - نسخه اصلاح شده"""
    try:
        if not os.path.exists('sales.xlsx'):
            print("❌ Sales file not found!")
            return None
            
        # بررسی شیت‌های موجود
        with pd.ExcelFile('sales.xlsx') as xls:
            sheet_names = xls.sheet_names
            print(f"📋 Available sheets in sales.xlsx: {sheet_names}")
            
            # اگر شیت 'sales' موجود نیست، اولین شیت را استفاده کن
            if 'sales' in sheet_names:
                sheet_name = 'sales'
            elif len(sheet_names) > 0:
                sheet_name = sheet_names[0]
                print(f"⚠️ Using sheet '{sheet_name}' instead of 'sales'")
            else:
                print("❌ No sheets found in sales file")
                return None
        
        df = pd.read_excel('sales.xlsx', sheet_name=sheet_name)
        print(f"✅ Sales file loaded successfully with {len(df)} records")
        print(f"📑 Columns: {list(df.columns)}")
        
        # پاک کردن فاصه‌های اضافی
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        # 🔧 FIX: تبدیل مقادیر NaN برای اجتناب از خطای JSON
        # برای ستون‌های عددی
        numeric_columns = ['Quantity', 'UnitPrice', 'TotalAmount']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].fillna(0)
        
        # برای ستون‌های متنی
        text_columns = ['CustomerCode', 'ProductCode', 'InvoiceDate', 'Status', 'Notes']
        for col in text_columns:
            if col in df.columns:
                df[col] = df[col].fillna('')
        
        print(f"📊 Sales data cleaned: {len(df)} records")
        return df
        
    except Exception as e:
        print(f"❌ Error loading sales file: {e}")
        return None

@app.route('/product_report/<customer_code>')
def product_report(customer_code):
    """گزارش کالاهای خریداری شده و نشده مشتری"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # بارگذاری اطلاعات مشتری
    customers_df = load_customers_from_excel()
    if customers_df is None:
        flash('خطا در بارگذاری اطلاعات مشتری!', 'error')
        return redirect(url_for('customers_list'))
    
    # پیدا کردن مشتری
    customer = customers_df[customers_df['CustomerCode'] == customer_code]
    if customer.empty:
        flash('مشتری یافت نشد!', 'error')
        return redirect(url_for('customers_list'))
    
    customer_info = customer.iloc[0].to_dict()
    
    # بررسی دسترسی
    if (session['user_info']['Typev'] != 'admin' and 
        customer_info['BazaryabCode'] != session['user_info']['Codev']):
        flash('شما اجازه مشاهده گزارش این مشتری را ندارید!', 'error')
        return redirect(url_for('customers_list'))
    
    return render_template('product_report.html', 
                         customer=customer_info,
                         user=session['user_info'])

@app.route('/get_product_data/<customer_code>')
def get_product_data(customer_code):
    """دریافت داده‌های کالا برای مشتری در بازه زمانی"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    date_from = request.args.get('date_from')
    date_to = request.args.get('date_to')
    date_type = request.args.get('date_type', 'gregorian')
    
    if not date_from or not date_to:
        return jsonify({'error': 'Date range required'}), 400
    
    # تبدیل تاریخ شمسی به میلادی در صورت نیاز
    if date_type == 'jalali':
        date_from_gregorian = jalali_to_gregorian(date_from)
        date_to_gregorian = jalali_to_gregorian(date_to)
        
        if not date_from_gregorian or not date_to_gregorian:
            return jsonify({'error': 'Invalid date format'}), 400
    else:
        date_from_gregorian = date_from
        date_to_gregorian = date_to
    
    # بارگذاری داده‌ها
    products_df = load_products_from_excel()
    sales_df = load_sales_from_excel()
    
    if products_df is None or sales_df is None:
        return jsonify({'error': 'Failed to load data'}), 500
    
    # تبدیل تاریخ‌های فروش به میلادی اگر شمسی هستند
    def convert_sale_date_to_gregorian(date_value):
        """تبدیل تاریخ فروش به میلادی"""
        if pd.isna(date_value):
            return None
        
        date_str = str(date_value).strip()
        
        # اگر شمسی است (شامل / است)
        if '/' in date_str and len(date_str.split('/')) == 3:
            return jalali_to_gregorian(date_str)
        
        # اگر قبلاً میلادی است
        if '-' in date_str and len(date_str) == 10:
            return date_str
        
        return date_str
    
    # تبدیل تمام تاریخ‌های فروش به میلادی
    sales_df_copy = sales_df.copy()
    sales_df_copy['InvoiceDateConverted'] = sales_df_copy['InvoiceDate'].apply(convert_sale_date_to_gregorian)
    
    # فیلتر فروش در بازه زمانی و مشتری
    customer_sales = sales_df_copy[
        (sales_df_copy['CustomerCode'] == customer_code) &
        (sales_df_copy['InvoiceDateConverted'] >= date_from_gregorian) &
        (sales_df_copy['InvoiceDateConverted'] <= date_to_gregorian)
    ]
    
    # محاسبه کل مبلغ خرید
    total_amount = customer_sales['TotalAmount'].sum()
    
    # کالاهای خریداری شده
    purchased_products = customer_sales['ProductCode'].unique()
    
    # تفکیک برند و مرتب‌سازی
    purchased_list = []
    not_purchased_list = []
    
    # گروه‌بندی بر اساس برند
    brands = products_df['Brand'].unique()
    
    for brand in sorted(brands):
        brand_products = products_df[products_df['Brand'] == brand].sort_values('Category')
        
        for _, product in brand_products.iterrows():
            # بررسی وجود عکس
            image_path = f"static/images/{product['ImageFile']}"
            if not os.path.exists(image_path):
                image_file = "null.jpg"
            else:
                image_file = product['ImageFile']
            
            product_data = {
                'ProductCode': product['ProductCode'],
                'ProductName': product['ProductName'],
                'Brand': product['Brand'],
                'Category': product['Category'],
                'Price': product['Price'],
                'ImageFile': image_file,
                'Description': product['Description'],
                'Offer1': product['Offer1'],
                'Offer2': product['Offer2'],
                'Offer3': product['Offer3']
            }
            
            if product['ProductCode'] in purchased_products:
                # محاسبه آمار خرید
                product_sales = customer_sales[customer_sales['ProductCode'] == product['ProductCode']]
                total_qty = product_sales['Quantity'].sum()
                product_amount = product_sales['TotalAmount'].sum()
                percentage = (product_amount / total_amount * 100) if total_amount > 0 else 0
                
                # تاریخ‌های خرید (نمایش اصلی)
                purchase_dates = []
                for _, sale in product_sales.iterrows():
                    original_date = sale['InvoiceDate']
                    # اگر تاریخ اصلی شمسی است، همان را نشان بده
                    if '/' in str(original_date):
                        display_date = str(original_date)
                        compact_date = str(original_date).replace('/', '')
                    else:
                        # اگر میلادی است، به شمسی تبدیل کن
                        display_date = gregorian_to_jalali(original_date)
                        compact_date = jalali_date_compact(original_date)
                    
                    purchase_dates.append({
                        'date': display_date,
                        'compact': compact_date,
                        'quantity': sale['Quantity'],
                        'amount': sale['TotalAmount']
                    })
                
                product_data.update({
                    'Purchased': True,
                    'TotalQuantity': int(total_qty),
                    'TotalAmount': int(product_amount),
                    'Percentage': round(percentage, 2),
                    'PurchaseDates': purchase_dates
                })
                purchased_list.append(product_data)
            else:
                product_data['Purchased'] = False
                not_purchased_list.append(product_data)
    
    return jsonify({
        'purchased': purchased_list,
        'not_purchased': not_purchased_list,
        'total_amount': int(total_amount),
        'date_from': date_from,
        'date_to': date_to,
        'date_from_jalali': date_from if date_type == 'jalali' else gregorian_to_jalali(date_from_gregorian),
        'date_to_jalali': date_to if date_type == 'jalali' else gregorian_to_jalali(date_to_gregorian),
        'date_type': date_type
    })

@app.route('/brand_report')
def brand_report():
    """گزارش برندی کالاها و مشتریان"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    return render_template('brand_report.html', user=session['user_info'])

@app.route('/get_brand_data')
def get_brand_data():
    """درիافت داده‌های برند و کالاها - اصلاح شده برای ترتیب بر اساس Radif"""
    try:
        if 'user_id' not in session:
            print("❌ Unauthorized access to get_brand_data")
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        print("📂 Loading brand data...")
        
        # بارگذاری داده‌ها
        products_df = load_products_from_excel()
        if products_df is None:
            print("❌ Products file not found")
            return jsonify({'error': 'فایل محصولات یافت نشد'}), 500
        
        print(f"✅ Products loaded: {len(products_df)} products")
        print(f"🔑 Product columns: {list(products_df.columns)}")
        
        # بارگذاری ترتیب برندها از شیت brand
        brand_order = load_brand_order_from_excel()
        
        if brand_order:
            print(f"📋 Brand order loaded from Excel: {brand_order}")
            # استفاده از ترتیب موجود در شیت brand
            ordered_brands = brand_order
        else:
            print("⚠️ No brand order found, using alphabetical order")
            # اگر شیت brand وجود ندارد، ترتیب الفبایی
            ordered_brands = sorted(products_df['Brand'].unique())
        
        print(f"🏷️ Final brand order: {ordered_brands}")
        
        # ایجاد دیکشنری کالاها برای هر برند
        brand_products = {}
        for brand in ordered_brands:
            brand_items = products_df[products_df['Brand'] == brand]
            products_list = []
            
            for _, product in brand_items.iterrows():
                # 🔧 FIX: اطمینان از عدم وجود NaN در هر فیلد
                product_data = {
                    'ProductCode': str(product.get('ProductCode', '')),
                    'ProductName': str(product.get('ProductName', '')),
                    'Category': str(product.get('Category', '')),
                    'Price': float(product.get('Price', 0)) if not pd.isna(product.get('Price', 0)) else 0,
                    'ImageFile': str(product.get('ImageFile', 'null.jpg'))
                }
                
                products_list.append(product_data)
            
            if products_list:  # فقط اگر برند دارای محصول باشد
                brand_products[brand] = products_list
                print(f"   {brand}: {len(products_list)} products")
        
        # فقط برندهایی که محصول دارند را برگردان
        final_brands = list(brand_products.keys())
        
        response_data = {
            'brands': final_brands,
            'brand_products': brand_products
        }
        
        # 🔧 DEBUG: چاپ response برای اطمینان
        import json
        response_json = json.dumps(response_data, ensure_ascii=False)
        print(f"📦 Brand data response size: {len(response_json)} characters")
        print(f"📊 Total brands with products: {len(final_brands)}")
        
        return jsonify(response_data)
        
    except Exception as e:
        print(f"❌ Error in get_brand_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/get_customers_by_product')
def get_customers_by_product():
    """دریافت مشتریانی که کالای خاص را خریده/نخریده‌اند - اصلاح شده برای خطای NaN"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            print("❌ Unauthorized access attempt")
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # دریافت پارامترها
        product_code = request.args.get('product_code')
        date_from = request.args.get('date_from', '')
        date_to = request.args.get('date_to', '')
        date_type = request.args.get('date_type', 'jalali')
        
        print(f"🔍 Request params: product_code={product_code}, date_from={date_from}, date_to={date_to}")
        
        if not product_code:
            return jsonify({'error': 'کد کالا الزامی است'}), 400
        
        # بارگذاری داده‌های اصلی
        print("📂 Loading data files...")
        customers_df = load_customers_from_excel()
        products_df = load_products_from_excel()
        
        # بررسی وجود فایل‌های ضروری
        if customers_df is None:
            print("❌ Customers file not found")
            return jsonify({'error': 'فایل مشتریان یافت نشد'}), 500
            
        if products_df is None:
            print("❌ Products file not found")
            return jsonify({'error': 'فایل محصولات یافت نشد'}), 500
        
        # بررسی وجود کالا
        product_info = products_df[products_df['ProductCode'] == product_code]
        if product_info.empty:
            print(f"❌ Product not found: {product_code}")
            return jsonify({'error': f'کالا با کد {product_code} یافت نشد'}), 404
        
        product_details = product_info.iloc[0].to_dict()
        print(f"✅ Product found: {product_details['ProductName']}")
        
        # فیلتر مشتریان بر اساس بازاریاب
        bazaryab_code = session['user_info']['Codev']
        if session['user_info']['Typev'] != 'admin':
            my_customers = customers_df[customers_df['BazaryabCode'] == bazaryab_code]
            print(f"👤 Filtering by bazaryab: {bazaryab_code}, found {len(my_customers)} customers")
        else:
            my_customers = customers_df
            print(f"👑 Admin access: showing all {len(my_customers)} customers")
        
        # بارگذاری فروش (اختیاری - اگر نباشه مشکلی نیست)
        sales_df = load_sales_from_excel()
        purchased_customer_codes = []
        customer_purchase_data = {}
        
        if sales_df is not None and not sales_df.empty:
            print("📊 Processing sales data...")
            print(f"📋 Sales columns: {list(sales_df.columns)}")
            
            # بررسی وجود ستون‌های مورد نیاز
            required_columns = ['CustomerCode', 'ProductCode', 'InvoiceDate']
            missing_columns = [col for col in required_columns if col not in sales_df.columns]
            
            if missing_columns:
                print(f"⚠️ Missing columns in sales file: {missing_columns}")
                print("📝 Available columns:", list(sales_df.columns))
                # ادامه می‌دیم بدون داده‌های فروش
            else:
                # تبدیل تاریخ اگر نیاز باشد
                date_from_gregorian = None
                date_to_gregorian = None
                
                if date_from and date_to:
                    if date_type == 'jalali':
                        date_from_gregorian = jalali_to_gregorian(date_from)
                        date_to_gregorian = jalali_to_gregorian(date_to)
                        print(f"📅 Date conversion: {date_from} -> {date_from_gregorian}, {date_to} -> {date_to_gregorian}")
                    else:
                        date_from_gregorian = date_from
                        date_to_gregorian = date_to
                
                # تبدیل تاریخ‌های فروش
                def convert_sale_date_to_gregorian(date_value):
                    if pd.isna(date_value):
                        return None
                    date_str = str(date_value).strip()
                    if '/' in date_str and len(date_str.split('/')) == 3:
                        return jalali_to_gregorian(date_str)
                    if '-' in date_str and len(date_str) == 10:
                        return date_str
                    return date_str
                
                sales_df_copy = sales_df.copy()
                sales_df_copy['InvoiceDateConverted'] = sales_df_copy['InvoiceDate'].apply(convert_sale_date_to_gregorian)
                
                # فیلتر فروش‌ها
                if date_from_gregorian and date_to_gregorian:
                    product_sales = sales_df_copy[
                        (sales_df_copy['ProductCode'] == product_code) &
                        (sales_df_copy['InvoiceDateConverted'] >= date_from_gregorian) &
                        (sales_df_copy['InvoiceDateConverted'] <= date_to_gregorian)
                    ]
                    print(f"📈 Filtered sales records: {len(product_sales)}")
                else:
                    product_sales = sales_df_copy[sales_df_copy['ProductCode'] == product_code]
                    print(f"📈 All sales records for product: {len(product_sales)}")
                
                # مشتریانی که این کالا را خریده‌اند
                purchased_customer_codes = product_sales['CustomerCode'].unique()
                print(f"👥 Customers who bought this product: {len(purchased_customer_codes)}")
                
                # محاسبه داده‌های خرید برای هر مشتری
                for customer_code in purchased_customer_codes:
                    customer_purchases = product_sales[product_sales['CustomerCode'] == customer_code]
                    
                    # محاسبه مجموع
                    total_qty = 0
                    total_amount = 0
                    purchase_dates = []
                    
                    for _, sale in customer_purchases.iterrows():
                        # مقادیر با مقدار پیش‌فرض
                        qty = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
                        amount = int(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
                        
                        total_qty += qty
                        total_amount += amount
                        
                        # تاریخ نمایش
                        original_date = sale.get('InvoiceDate', '')
                        if '/' in str(original_date):
                            display_date = str(original_date)
                        else:
                            display_date = gregorian_to_jalali(str(original_date)) if original_date else ''
                        
                        purchase_dates.append({
                            'date': display_date,
                            'quantity': qty,
                            'amount': amount
                        })
                    
                    customer_purchase_data[customer_code] = {
                        'TotalQuantity': total_qty,
                        'TotalAmount': total_amount,
                        'PurchaseDates': purchase_dates
                    }
        else:
            print("⚠️ No sales data found - showing customers without purchase history")
        
        # تفکیک مشتریان
        purchased_customers = []
        not_purchased_customers = []
        
        # بارگذاری اطلاعات کاربران برای نام‌های بازاریاب
        users_df = load_users_from_excel()
        
        for _, customer in my_customers.iterrows():
            customer_code = customer['CustomerCode']
            
            # نام بازاریاب
            bazaryab_name = "نامشخص"
            if users_df is not None:
                bazaryab = users_df[users_df['Codev'] == customer['BazaryabCode']]
                if not bazaryab.empty:
                    bazaryab_name = bazaryab.iloc[0]['Namev']
            
            # 🔧 FIX: تبدیل مقادیر NaN به boolean
            location_set = customer.get('LocationSet', False)
            if pd.isna(location_set):
                location_set = False
            elif isinstance(location_set, str):
                location_set = location_set.lower() in ['true', '1', 'yes', 'بله']
            else:
                location_set = bool(location_set)
            
            customer_data = {
                'CustomerCode': str(customer['CustomerCode']),  # تبدیل به string
                'CustomerName': str(customer['CustomerName']),
                'BazaryabCode': str(customer['BazaryabCode']),
                'BazaryabName': bazaryab_name,
                'LocationSet': location_set  # ✅ حالا boolean است، نه NaN
            }
            
            # اگر این مشتری کالا رو خریده
            if customer_code in purchased_customer_codes and customer_code in customer_purchase_data:
                customer_data.update(customer_purchase_data[customer_code])
                purchased_customers.append(customer_data)
            else:
                not_purchased_customers.append(customer_data)
        
        print(f"✅ Final result: {len(purchased_customers)} purchased, {len(not_purchased_customers)} not purchased")
        
        # 🔧 FIX: اطمینان از عدم وجود NaN در product_details
        clean_product_details = {}
        for key, value in product_details.items():
            if pd.isna(value):
                if key in ['Price', 'Stock']:
                    clean_product_details[key] = 0
                else:
                    clean_product_details[key] = ""
            else:
                clean_product_details[key] = value
        
        response_data = {
            'product': clean_product_details,
            'purchased_customers': purchased_customers,
            'not_purchased_customers': not_purchased_customers,
            'date_from': date_from,
            'date_to': date_to,
            'date_type': date_type,
            'total_purchased': len(purchased_customers),
            'total_not_purchased': len(not_purchased_customers)
        }
        
        # 🔧 DEBUG: چاپ تعداد کاراکترهای response برای اطمینان
        import json
        response_json = json.dumps(response_data, ensure_ascii=False)
        print(f"📦 Response size: {len(response_json)} characters")
        
        return jsonify(response_data)
        
    except Exception as e:
        print(f"❌ Error in get_customers_by_product: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

def load_orders_from_excel():
    """بارگذاری سفارشات از فایل Excel"""
    try:
        if not os.path.exists('orders.xlsx'):
            # ایجاد فایل خالی اگر وجود نداشته باشد
            empty_df = pd.DataFrame(columns=[
                'OrderNumber', 'DocumentNumber', 'BazaryabCode', 'CustomerCode', 
                'ProductCode', 'Quantity', 'UnitPrice', 'TotalAmount', 
                'OrderDate', 'OrderTime', 'Status', 'Notes'
            ])
            empty_df.to_excel('orders.xlsx', sheet_name='orders', index=False)
            return empty_df
            
        df = pd.read_excel('orders.xlsx', sheet_name='orders')
        print("✅ Orders file loaded successfully")
        return df
    except Exception as e:
        print("❌ Error loading orders file:", e)
        return None

def save_orders_to_excel(df):
    """ذخیره سفارشات در فایل Excel"""
    try:
        df.to_excel('orders.xlsx', sheet_name='orders', index=False)
        print("✅ Orders file saved successfully")
        return True
    except Exception as e:
        print("❌ Error saving orders file:", e)
        return False

def generate_order_number():
    """تولید شماره سفارش منحصر به فرد"""
    now = datetime.now()
    jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
    date_str = jalali_now.strftime('%Y%m%d')
    
    # بررسی آخرین شماره سفارش امروز
    orders_df = load_orders_from_excel()
    if orders_df is not None and len(orders_df) > 0:
        today_orders = orders_df[orders_df['OrderNumber'].str.contains(f'ORD-{date_str}')]
        if len(today_orders) > 0:
            last_number = len(today_orders) + 1
        else:
            last_number = 1
    else:
        last_number = 1
    
    return f"ORD-{date_str}{last_number:03d}"

def generate_document_number():
    """تولید شماره سند منحصر به فرد"""
    now = datetime.now()
    date_str = now.strftime('%y%m%d')
    
    # بررسی آخرین شماره سند امروز
    orders_df = load_orders_from_excel()
    if orders_df is not None and len(orders_df) > 0:
        today_docs = orders_df[orders_df['DocumentNumber'].str.contains(f'DOC-{date_str}')]
        if len(today_docs) > 0:
            last_number = len(today_docs) + 1
        else:
            last_number = 1
    else:
        last_number = 1
    
    return f"DOC-{date_str}{last_number:03d}"

@app.route('/catalog')
def catalog():
    """صفحه کاتالوگ کالاها"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    return render_template('catalog.html', user=session['user_info'])

@app.route('/get_catalog_data')
def get_catalog_data():
    """دریافت داده‌های کاتالوگ"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    # بارگذاری داده‌ها
    products_df = load_products_from_excel()
    customers_df = load_customers_from_excel()
    
    if products_df is None or customers_df is None:
        return jsonify({'error': 'Failed to load data'}), 500
    
    # فیلتر مشتریان بر اساس بازاریاب
    bazaryab_code = session['user_info']['Codev']
    if session['user_info']['Typev'] != 'admin':
        my_customers = customers_df[customers_df['BazaryabCode'] == bazaryab_code]
    else:
        my_customers = customers_df
    
    # تنظیم کالاها بر اساس برند
    brands = {}
    for _, product in products_df.iterrows():
        brand = product['Brand']
        if brand not in brands:
            brands[brand] = []
        
        # بررسی وجود عکس
        image_path = f"static/images/{product['ImageFile']}"
        if not os.path.exists(image_path):
            image_file = "null.jpg"
        else:
            image_file = product['ImageFile']
        
        brands[brand].append({
            'ProductCode': product['ProductCode'],
            'ProductName': product['ProductName'],
            'Category': product['Category'],
            'Brand': product['Brand'],
            'Price': product['Price'],
            'Stock': product['Stock'],
            'ImageFile': image_file,
            'Description': product['Description'],
            'Offer1': product['Offer1'],
            'Offer2': product['Offer2'],
            'Offer3': product['Offer3'],
            'radif': product.get('radif', product.get('Radif', product.get('RADIF', 999999)))
        })
    
    # لیست مشتریان
    customers_list = []
    for _, customer in my_customers.iterrows():
        customers_list.append({
            'CustomerCode': customer['CustomerCode'],
            'CustomerName': customer['CustomerName']
        })
    
    return jsonify({
        'brands': brands,
        'customers': customers_list
    })

# ✅ اصلاح شده - داشبورد فروش کاربر
@app.route('/user_dashboard')
def user_dashboard():
    """داشبورد فروش کاربر"""
    # چک کردن لاگین
    if 'user_id' not in session:  # ✅ درست شد
        return redirect(url_for('login'))
    
    # چک کردن نوع کاربر
    user_type = session['user_info'].get('Typev', '')  # ✅ درست شد
    if user_type != 'user':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('user_dashboard.html', user=session['user_info'])

# ✅ اصلاح شده - گزارش فروش ماهانه  
@app.route('/get_sales_report', methods=['POST'])
def get_sales_report():
    """گزارش فروش ماهانه"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'})
        
        # چک نوع کاربر
        if session['user_info'].get('Typev') != 'user':
            return jsonify({'error': 'دسترسی غیرمجاز'})
        
        # دریافت داده‌های POST
        data = request.get_json()
        year = data.get('year')
        month = data.get('month')
        
        if not year or not month:
            return jsonify({'error': 'سال و ماه الزامی است'})
        
        # بارگذاری داده‌ها از فایل‌های اصلی
        try:
            # بارگذاری فایل Sales
            sales_df = load_sales_from_excel()
            if sales_df is None:
                return jsonify({'error': 'فایل فروش یافت نشد'})
            
            # بارگذاری فایل Customers برای نام مشتریان
            customers_df = load_customers_from_excel()
            if customers_df is None:
                return jsonify({'error': 'فایل مشتریان یافت نشد'})
            
        except Exception as e:
            return jsonify({'error': f'خطا در بارگذاری فایل: {str(e)}'})
        
        # 🔥 اضافه کردن فیلتر بازاریاب اینجا!
        bazaryab_code = session['user_info']['Codev']
        
        # فیلتر مشتریان این بازاریاب
        my_customers = customers_df[customers_df['BazaryabCode'] == bazaryab_code]
        my_customer_codes = my_customers['CustomerCode'].tolist()
        
        # فیلتر فروش‌های مربوط به مشتریان این بازاریاب
        my_sales = sales_df[sales_df['CustomerCode'].isin(my_customer_codes)]
        
        # فیلتر کردن داده‌ها بر اساس تاریخ شمسی
        filtered_sales = filter_sales_by_jalali_date(my_sales, year, month)
        
        if filtered_sales.empty:
            return jsonify({
                'customers': [],
                'total_sales': 0,
                'year': year,
                'month': month
            })
        
        # محاسبه فروش هر مشتری (فقط برای مشتریان این بازاریاب)
        sales_summary = calculate_customer_sales_summary(filtered_sales, my_customers)
        
        return jsonify({
            'customers': sales_summary['customers'],
            'total_sales': sales_summary['total_sales'],
            'year': year,
            'month': month
        })
        
    except Exception as e:
        return jsonify({'error': f'خطا در پردازش: {str(e)}'})

def filter_sales_by_jalali_date(sales_df, year, month):
    """فیلتر کردن فروش بر اساس سال و ماه شمسی"""
    try:
        if sales_df.empty:
            return pd.DataFrame()
        
        # فرض: ستون تاریخ InvoiceDate نام داره
        if 'InvoiceDate' not in sales_df.columns:
            return pd.DataFrame()
        
        filtered_rows = []
        
        for index, row in sales_df.iterrows():
            try:
                invoice_date = row['InvoiceDate']
                
                if pd.isna(invoice_date):
                    continue
                
                # تبدیل تاریخ به شمسی
                if isinstance(invoice_date, str):
                    # اگر تاریخ رشته‌ای است
                    if '/' in invoice_date:
                        # فرمت شمسی: 1403/01/15
                        date_parts = invoice_date.split('/')
                        if len(date_parts) == 3:
                            invoice_year = int(date_parts[0])
                            invoice_month = int(date_parts[1])
                            
                            if invoice_year == year and invoice_month == month:
                                filtered_rows.append(row)
                    elif '-' in invoice_date:
                        # فرمت میلادی: 2024-03-21
                        gregorian_date = datetime.strptime(invoice_date, '%Y-%m-%d').date()
                        jalali_date = jdatetime.date.fromgregorian(date=gregorian_date)
                        
                        if jalali_date.year == year and jalali_date.month == month:
                            filtered_rows.append(row)
                
                elif hasattr(invoice_date, 'year'):
                    # اگر datetime object است
                    jalali_date = jdatetime.date.fromgregorian(
                        year=invoice_date.year,
                        month=invoice_date.month,
                        day=invoice_date.day
                    )
                    
                    if jalali_date.year == year and jalali_date.month == month:
                        filtered_rows.append(row)
                        
            except (ValueError, AttributeError):
                continue
        
        return pd.DataFrame(filtered_rows) if filtered_rows else pd.DataFrame()
        
    except Exception as e:
        print(f"Error in filter_sales_by_jalali_date: {e}")
        return pd.DataFrame()

def calculate_customer_sales_summary(sales_df, customers_df):
    """محاسبه خلاصه فروش مشتریان"""
    try:
        if sales_df.empty:
            return {'customers': [], 'total_sales': 0}
        
        # محاسبه فروش هر مشتری
        customer_sales = sales_df.groupby('CustomerCode')['TotalAmount'].sum().to_dict()
        
        # ایجاد لیست نهایی با نام مشتریان
        customers_list = []
        total_sales = 0
        
        for customer_code, sales_amount in customer_sales.items():
            if sales_amount > 0:
                # پیدا کردن نام مشتری
                customer_name = 'نامشخص'
                customer_row = customers_df[customers_df['CustomerCode'] == customer_code]
                if not customer_row.empty:
                    customer_name = customer_row.iloc[0]['CustomerName']
                
                customers_list.append({
                    'customer_code': customer_code,
                    'customer_name': customer_name,
                    'sales_amount': int(sales_amount)
                })
                
                total_sales += sales_amount
        
        # مرتب‌سازی بر اساس مقدار فروش (از زیاد به کم)
        customers_list.sort(key=lambda x: x['sales_amount'], reverse=True)
        
        return {
            'customers': customers_list,
            'total_sales': int(total_sales)
        }
        
    except Exception as e:
        print(f"Error in calculate_customer_sales_summary: {e}")
        return {'customers': [], 'total_sales': 0}

# این کدها رو به فایل app.py اضافه کنید

@app.route('/sales_performance_report')
def sales_performance_report():
    """گزارش عملکرد بازاریابان - فقط برای ادمین"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط ادمین می‌تونه این گزارش رو ببینه
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('sales_performance_report.html', user=session['user_info'])

@app.route('/get_performance_report')
def get_performance_report():
    """دریافت داده‌های گزارش عملکرد بازاریابان"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط ادمین
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # دریافت پارامترها
        date_from = request.args.get('date_from', '').strip()
        date_to = request.args.get('date_to', '').strip()
        
        if not date_from or not date_to:
            return jsonify({'error': 'بازه زمانی الزامی است'}), 400
        
        print(f"🔍 Performance report request: {date_from} to {date_to}")
        
        # تبدیل تاریخ شمسی به میلادی
        date_from_gregorian = jalali_to_gregorian(date_from)
        date_to_gregorian = jalali_to_gregorian(date_to)
        
        if not date_from_gregorian or not date_to_gregorian:
            return jsonify({'error': 'فرمت تاریخ نامعتبر است'}), 400
        
        # بارگذاری داده‌ها
        users_df = load_users_from_excel()
        customers_df = load_customers_from_excel() 
        visits_df = load_visits_from_excel()
        sales_df = load_sales_from_excel()
        
        if users_df is None or customers_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل‌ها'}), 500
        
        # فیلتر بازاریابان (فقط کاربران با نوع user)
        salespeople = users_df[users_df['Typev'] == 'user']
        
        if salespeople.empty:
            return jsonify({'error': 'هیچ بازاریابی یافت نشد'}), 404
        
        performance_data = []
        
        for _, salesperson in salespeople.iterrows():
            salesperson_code = salesperson['Codev']
            salesperson_name = salesperson['Namev']
            
            print(f"📊 Processing salesperson: {salesperson_name} ({salesperson_code})")
            
            # تعداد مشتریان این بازاریاب
            sp_customers = customers_df[customers_df['BazaryabCode'] == salesperson_code]
            total_customers = len(sp_customers)
            
            # مراجعات در بازه زمانی
            total_visits = 0
            if visits_df is not None and not visits_df.empty:
                sp_visits = visits_df[visits_df['BazaryabCode'] == salesperson_code]
                
                # فیلتر بر اساس تاریخ
                filtered_visits = []
                for _, visit in sp_visits.iterrows():
                    visit_date = visit.get('VisitDate', '')
                    if visit_date:
                        # تبدیل تاریخ ویزیت به میلادی برای مقایسه
                        if isinstance(visit_date, str) and len(visit_date) == 10:
                            # اگر میلادی است: 2024-03-21
                            visit_gregorian = visit_date
                        else:
                            # اگر شمسی است، تبدیل کن
                            visit_gregorian = jalali_to_gregorian(str(visit_date))
                        
                        if (visit_gregorian and 
                            visit_gregorian >= date_from_gregorian and 
                            visit_gregorian <= date_to_gregorian):
                            filtered_visits.append(visit)
                
                total_visits = len(filtered_visits)
            
            # فروش در بازه زمانی
            total_sales = 0
            if sales_df is not None and not sales_df.empty:
                # فیلتر فروش‌های مشتریان این بازاریاب
                customer_codes = sp_customers['CustomerCode'].tolist()
                sp_sales = sales_df[sales_df['CustomerCode'].isin(customer_codes)]
                
                # فیلتر بر اساس تاریخ
                filtered_sales = []
                for _, sale in sp_sales.iterrows():
                    sale_date = sale.get('InvoiceDate', '')
                    if sale_date:
                        # تبدیل تاریخ فروش
                        if '/' in str(sale_date):
                            # شمسی: 1404/01/15
                            sale_gregorian = jalali_to_gregorian(str(sale_date))
                        elif '-' in str(sale_date) and len(str(sale_date)) == 10:
                            # میلادی: 2024-03-21
                            sale_gregorian = str(sale_date)
                        else:
                            sale_gregorian = None
                        
                        if (sale_gregorian and 
                            sale_gregorian >= date_from_gregorian and 
                            sale_gregorian <= date_to_gregorian):
                            filtered_sales.append(sale)
                
                # محاسبه مجموع فروش
                for sale in filtered_sales:
                    amount = sale.get('TotalAmount', 0)
                    if not pd.isna(amount):
                        total_sales += float(amount)
            
            # محاسبه نرخ تبدیل
            conversion_rate = 0
            if total_visits > 0:
                # فرض: هر فروش یعنی یک مراجعه موفق
                successful_visits = len(filtered_sales) if 'filtered_sales' in locals() else 0
                conversion_rate = (successful_visits / total_visits) * 100
            
            performance_data.append({
                'salesperson_code': salesperson_code,
                'salesperson_name': salesperson_name,
                'total_customers': int(total_customers),
                'total_visits': int(total_visits),
                'total_sales': int(total_sales),
                'conversion_rate': round(conversion_rate, 1)
            })
            
            print(f"   👥 Customers: {total_customers}")
            print(f"   🚪 Visits: {total_visits}")
            print(f"   💰 Sales: {total_sales:,}")
            print(f"   📊 Conversion: {conversion_rate:.1f}%")
        
        # مرتب‌سازی بر اساس مجموع فروش (بالا به پایین)
        performance_data.sort(key=lambda x: x['total_sales'], reverse=True)
        
        print(f"✅ Performance report generated for {len(performance_data)} salespeople")
        
        return jsonify({
            'salespeople': performance_data,
            'date_from': date_from,
            'date_to': date_to,
            'period_info': f"{date_from} تا {date_to}"
        })
        
    except Exception as e:
        print(f"❌ Error in get_performance_report: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500


# این کدها رو به فایل app.py اضافه کنید

@app.route('/admin_brand_sales_report')
def admin_brand_sales_report():
    """گزارش فروش برندی همه بازاریابان - فقط برای ادمین"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط ادمین می‌تونه این گزارش رو ببینه
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('admin_brand_sales_report.html', user=session['user_info'])

@app.route('/get_admin_brand_sales_data', methods=['POST'])
def get_admin_brand_sales_data():
    """دریافت داده‌های فروش برندی همه بازاریابان"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط ادمین
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # دریافت داده‌های POST
        data = request.get_json()
        date_from = data.get('date_from', '').strip()
        date_to = data.get('date_to', '').strip()
        date_type = data.get('date_type', 'jalali')
        
        if not date_from or not date_to:
            return jsonify({'error': 'بازه زمانی الزامی است'}), 400
        
        print(f"🔍 Admin brand sales report: {date_from} to {date_to} ({date_type})")
        
        # تبدیل تاریخ به میلادی اگر لازم باشه
        if date_type == 'jalali':
            date_from_gregorian = jalali_to_gregorian(date_from)
            date_to_gregorian = jalali_to_gregorian(date_to)
            
            if not date_from_gregorian or not date_to_gregorian:
                return jsonify({'error': 'فرمت تاریخ نامعتبر است'}), 400
        else:
            date_from_gregorian = date_from
            date_to_gregorian = date_to
        
        # بارگذاری داده‌ها
        products_df = load_products_from_excel()
        customers_df = load_customers_from_excel()
        sales_df = load_sales_from_excel()
        users_df = load_users_from_excel()
        
        if products_df is None or customers_df is None or sales_df is None or users_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل‌ها'}), 500
        
        # فیلتر بازاریابان (فقط کاربران با نوع user)
        salespeople = users_df[users_df['Typev'] == 'user']
        
        if salespeople.empty:
            return jsonify({'error': 'هیچ بازاریابی یافت نشد'}), 404
        
        print(f"👥 Found {len(salespeople)} salespeople")
        
        # فیلتر فروش‌ها بر اساس تاریخ
        def convert_sale_date_to_gregorian(date_value):
            if pd.isna(date_value):
                return None
            date_str = str(date_value).strip()
            if '/' in date_str and len(date_str.split('/')) == 3:
                return jalali_to_gregorian(date_str)
            elif '-' in date_str and len(date_str) == 10:
                return date_str
            return None
        
        # تبدیل تاریخ‌های فروش
        sales_df_copy = sales_df.copy()
        sales_df_copy['InvoiceDateConverted'] = sales_df_copy['InvoiceDate'].apply(convert_sale_date_to_gregorian)
        
        # فیلتر بر اساس بازه زمانی
        filtered_sales = sales_df_copy[
            (sales_df_copy['InvoiceDateConverted'] >= date_from_gregorian) &
            (sales_df_copy['InvoiceDateConverted'] <= date_to_gregorian)
        ]
        
        if filtered_sales.empty:
            return jsonify({
                'brands': [],
                'salespeople': [],
                'total_sales': 0,
                'date_from': date_from,
                'date_to': date_to,
                'date_type': date_type
            })
        
        print(f"📊 Filtered sales: {len(filtered_sales)} records")
        
        # ایجاد دیکشنری برای نگاشت مشتری به بازاریاب
        customer_to_salesperson = {}
        for _, customer in customers_df.iterrows():
            customer_to_salesperson[customer['CustomerCode']] = customer['BazaryabCode']
        
        # محاسبه فروش هر بازاریاب از هر کالا
        salesperson_product_sales = {}
        
        for _, sale in filtered_sales.iterrows():
            customer_code = sale['CustomerCode']
            product_code = sale['ProductCode']
            amount = float(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
            quantity = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
            
            # پیدا کردن بازاریاب این مشتری
            salesperson_code = customer_to_salesperson.get(customer_code)
            
            if salesperson_code:
                if salesperson_code not in salesperson_product_sales:
                    salesperson_product_sales[salesperson_code] = {}
                
                if product_code not in salesperson_product_sales[salesperson_code]:
                    salesperson_product_sales[salesperson_code][product_code] = {'amount': 0, 'quantity': 0}
                
                salesperson_product_sales[salesperson_code][product_code]['amount'] += amount
                salesperson_product_sales[salesperson_code][product_code]['quantity'] += quantity
        
        # تفکیک بر اساس برند
        brand_data = {}
        
        # دریافت همه برندها و مرتب‌سازی بر اساس Radif
        brands_radif = {}
        for _, product in products_df.iterrows():
            brand = product['Brand']
            radif = int(product.get('Radif', 999999))
            if brand not in brands_radif or radif < brands_radif[brand]:
                brands_radif[brand] = radif
        
        # مرتب‌سازی برندها بر اساس Radif
        sorted_brands = sorted(brands_radif.keys(), key=lambda x: brands_radif[x])
        
        for brand in sorted_brands:
            brand_data[brand] = {
                'brand_name': brand,
                'radif': brands_radif[brand],
                'total_amount': 0,
                'total_quantity': 0,
                'salespeople_sales': [],
                'products': []
            }
            
            # پیدا کردن کالاهای این برند
            brand_products = products_df[products_df['Brand'] == brand]['ProductCode'].tolist()
            
            # محاسبه فروش هر بازاریاب از این برند
            for _, salesperson in salespeople.iterrows():
                sp_code = salesperson['Codev']
                sp_name = salesperson['Namev']
                
                sp_brand_amount = 0
                sp_brand_quantity = 0
                sp_products = []
                
                if sp_code in salesperson_product_sales:
                    for product_code in brand_products:
                        if product_code in salesperson_product_sales[sp_code]:
                            product_sales = salesperson_product_sales[sp_code][product_code]
                            sp_brand_amount += product_sales['amount']
                            sp_brand_quantity += product_sales['quantity']
                            
                            # اطلاعات کالا
                            product_info = products_df[products_df['ProductCode'] == product_code]
                            if not product_info.empty:
                                product_detail = product_info.iloc[0]
                                sp_products.append({
                                    'product_code': product_code,
                                    'product_name': product_detail['ProductName'],
                                    'category': product_detail.get('Category', ''),
                                    'amount': int(product_sales['amount']),
                                    'quantity': int(product_sales['quantity'])
                                })
                
                # اگر این بازاریاب از این برند فروش داشته
                if sp_brand_amount > 0:
                    brand_data[brand]['salespeople_sales'].append({
                        'salesperson_code': sp_code,
                        'salesperson_name': sp_name,
                        'total_amount': int(sp_brand_amount),
                        'total_quantity': int(sp_brand_quantity),
                        'products': sorted(sp_products, key=lambda x: x['amount'], reverse=True)
                    })
                    
                    brand_data[brand]['total_amount'] += sp_brand_amount
                    brand_data[brand]['total_quantity'] += sp_brand_quantity
            
            # مرتب‌سازی بازاریابان بر اساس فروش (بالا به پایین)
            brand_data[brand]['salespeople_sales'].sort(key=lambda x: x['total_amount'], reverse=True)
        
        # حذف برندهایی که فروش نداشتن
        filtered_brands = []
        total_sales = 0
        
        for brand in sorted_brands:
            if brand_data[brand]['total_amount'] > 0:
                brand_data[brand]['total_amount'] = int(brand_data[brand]['total_amount'])
                brand_data[brand]['total_quantity'] = int(brand_data[brand]['total_quantity'])
                filtered_brands.append(brand_data[brand])
                total_sales += brand_data[brand]['total_amount']
        
        # آمار کلی بازاریابان
        salespeople_summary = []
        for _, salesperson in salespeople.iterrows():
            sp_code = salesperson['Codev']
            sp_name = salesperson['Namev']
            
            sp_total = 0
            if sp_code in salesperson_product_sales:
                for product_sales in salesperson_product_sales[sp_code].values():
                    sp_total += product_sales['amount']
            
            salespeople_summary.append({
                'salesperson_code': sp_code,
                'salesperson_name': sp_name,
                'total_sales': int(sp_total)
            })
        
        # مرتب‌سازی بازاریابان بر اساس فروش
        salespeople_summary.sort(key=lambda x: x['total_sales'], reverse=True)
        
        print(f"✅ Admin brand report: {len(filtered_brands)} brands, total: {total_sales:,}")
        
        return jsonify({
            'brands': filtered_brands,
            'salespeople': salespeople_summary,
            'total_sales': int(total_sales),
            'date_from': date_from,
            'date_to': date_to,
            'date_type': date_type,
            'period_info': f"{date_from} تا {date_to}"
        })
        
    except Exception as e:
        print(f"❌ Error in get_admin_brand_sales_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

# این کدها رو به فایل app.py اضافه کنید

@app.route('/user_brand_sales_report')
def user_brand_sales_report():
    """گزارش فروش برندی برای کاربر عادی"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط کاربران عادی (user) می‌تونن این گزارش رو ببینن
    if session['user_info']['Typev'] != 'user':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('user_brand_sales_report.html', user=session['user_info'])



@app.route('/get_user_brand_sales_data', methods=['POST'])
def get_user_brand_sales_data():
    """دریافت داده‌های فروش برندی برای کاربر عادی"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط کاربر عادی
        if session['user_info']['Typev'] != 'user':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # دریافت داده‌های POST
        data = request.get_json()
        date_from = data.get('date_from', '').strip()
        date_to = data.get('date_to', '').strip()
        date_type = data.get('date_type', 'jalali')
        
        if not date_from or not date_to:
            return jsonify({'error': 'بازه زمانی الزامی است'}), 400
        
        print(f"🔍 Brand sales report request: {date_from} to {date_to} ({date_type})")
        
        # تبدیل تاریخ به میلادی اگر لازم باشه
        if date_type == 'jalali':
            date_from_gregorian = jalali_to_gregorian(date_from)
            date_to_gregorian = jalali_to_gregorian(date_to)
            
            if not date_from_gregorian or not date_to_gregorian:
                return jsonify({'error': 'فرمت تاریخ نامعتبر است'}), 400
        else:
            date_from_gregorian = date_from
            date_to_gregorian = date_to
        
        # بارگذاری داده‌ها
        products_df = load_products_from_excel()
        customers_df = load_customers_from_excel()
        sales_df = load_sales_from_excel()
        
        if products_df is None or customers_df is None or sales_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل‌ها'}), 500
        
        # فیلتر مشتریان این بازاریاب
        bazaryab_code = session['user_info']['Codev']
        my_customers = customers_df[customers_df['BazaryabCode'] == bazaryab_code]
        customer_codes = my_customers['CustomerCode'].tolist()
        
        if not customer_codes:
            return jsonify({'error': 'هیچ مشتری برای شما تعریف نشده است'}), 404
        
        print(f"👥 Found {len(customer_codes)} customers for bazaryab {bazaryab_code}")
        
        # فیلتر فروش‌های مشتریان این بازاریاب
        my_sales = sales_df[sales_df['CustomerCode'].isin(customer_codes)]
        
        if my_sales.empty:
            return jsonify({
                'brands': [],
                'total_sales': 0,
                'date_from': date_from,
                'date_to': date_to,
                'date_type': date_type
            })
        
        # فیلتر بر اساس تاریخ
        def convert_sale_date_to_gregorian(date_value):
            if pd.isna(date_value):
                return None
            date_str = str(date_value).strip()
            if '/' in date_str and len(date_str.split('/')) == 3:
                return jalali_to_gregorian(date_str)
            elif '-' in date_str and len(date_str) == 10:
                return date_str
            return None
        
        # تبدیل تاریخ‌های فروش
        my_sales_copy = my_sales.copy()
        my_sales_copy['InvoiceDateConverted'] = my_sales_copy['InvoiceDate'].apply(convert_sale_date_to_gregorian)
        
        # فیلتر بر اساس بازه زمانی
        filtered_sales = my_sales_copy[
            (my_sales_copy['InvoiceDateConverted'] >= date_from_gregorian) &
            (my_sales_copy['InvoiceDateConverted'] <= date_to_gregorian)
        ]
        
        if filtered_sales.empty:
            return jsonify({
                'brands': [],
                'total_sales': 0,
                'date_from': date_from,
                'date_to': date_to,
                'date_type': date_type
            })
        
        print(f"📊 Filtered sales: {len(filtered_sales)} records")
        
        # محاسبه فروش هر کالا
        product_sales = {}
        for _, sale in filtered_sales.iterrows():
            product_code = sale['ProductCode']
            amount = float(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
            quantity = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
            
            if product_code not in product_sales:
                product_sales[product_code] = {'amount': 0, 'quantity': 0}
            
            product_sales[product_code]['amount'] += amount
            product_sales[product_code]['quantity'] += quantity
        
        # تفکیک بر اساس برند و محاسبه مجموع هر برند
        brand_sales = {}
        
        for product_code, sales_data in product_sales.items():
            # پیدا کردن اطلاعات کالا
            product_info = products_df[products_df['ProductCode'] == product_code]
            
            if not product_info.empty:
                product_detail = product_info.iloc[0]
                brand = product_detail['Brand']
                radif = int(product_detail.get('Radif', 999999))  # اگر Radif نداشته باشه، آخر قرار بگیره
                
                if brand not in brand_sales:
                    brand_sales[brand] = {
                        'brand_name': brand,
                        'radif': radif,
                        'total_amount': 0,
                        'total_quantity': 0,
                        'products': []
                    }
                
                # اضافه کردن به مجموع برند
                brand_sales[brand]['total_amount'] += sales_data['amount']
                brand_sales[brand]['total_quantity'] += sales_data['quantity']
                
                # اضافه کردن جزئیات کالا
                brand_sales[brand]['products'].append({
                    'product_code': product_code,
                    'product_name': product_detail['ProductName'],
                    'category': product_detail.get('Category', ''),
                    'amount': int(sales_data['amount']),
                    'quantity': int(sales_data['quantity'])
                })
        
        # مرتب‌سازی برندها بر اساس Radif
        sorted_brands = sorted(brand_sales.values(), key=lambda x: x['radif'])
        
        # مرتب‌سازی کالاهای هر برند بر اساس مقدار فروش
        for brand in sorted_brands:
            brand['products'].sort(key=lambda x: x['amount'], reverse=True)
            brand['total_amount'] = int(brand['total_amount'])
            brand['total_quantity'] = int(brand['total_quantity'])
        
        # محاسبه مجموع کل فروش
        total_sales = sum([brand['total_amount'] for brand in sorted_brands])
        
        print(f"✅ Brand sales report: {len(sorted_brands)} brands, total: {total_sales:,}")
        
        return jsonify({
            'brands': sorted_brands,
            'total_sales': int(total_sales),
            'date_from': date_from,
            'date_to': date_to,
            'date_type': date_type,
            'period_info': f"{date_from} تا {date_to}"
        })
        
    except Exception as e:
        print(f"❌ Error in get_user_brand_sales_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/submit_order', methods=['POST'])
def submit_order():
    """ثبت سفارش جدید"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    try:
        data = request.get_json()
        
        customer_code = data.get('customer_code')
        product_code = data.get('product_code')
        quantity = int(data.get('quantity', 1))
        notes = data.get('notes', '')
        
        if not customer_code or not product_code:
            return jsonify({'error': 'Customer and product required'}), 400
        
        # دریافت اطلاعات کالا
        products_df = load_products_from_excel()
        product = products_df[products_df['ProductCode'] == product_code]
        
        if product.empty:
            return jsonify({'error': 'Product not found'}), 404
        
        product_info = product.iloc[0]
        unit_price = product_info['Price']
        total_amount = unit_price * quantity
        
        # تولید شماره‌های منحصر به فرد
        order_number = generate_order_number()
        document_number = generate_document_number()
        
        # تاریخ و ساعت فعلی
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        order_date = jalali_now.strftime('%Y/%m/%d')
        order_time = now.strftime('%H:%M')
        
        # ایجاد سفارش جدید
        new_order = {
            'OrderNumber': order_number,
            'DocumentNumber': document_number,
            'BazaryabCode': session['user_info']['Codev'],
            'CustomerCode': customer_code,
            'ProductCode': product_code,
            'Quantity': quantity,
            'UnitPrice': unit_price,
            'TotalAmount': total_amount,
            'OrderDate': order_date,
            'OrderTime': order_time,
            'Status': 'ثبت شده',
            'Notes': notes
        }
        
        # اضافه کردن به فایل
        orders_df = load_orders_from_excel()
        if orders_df is None:
            orders_df = pd.DataFrame(columns=list(new_order.keys()))
        
        new_row = pd.DataFrame([new_order])
        orders_df = pd.concat([orders_df, new_row], ignore_index=True)
        
        # ذخیره فایل
        if save_orders_to_excel(orders_df):
            return jsonify({
                'success': True,
                'order_number': order_number,
                'document_number': document_number,
                'total_amount': total_amount,
                'message': 'سفارش با موفقیت ثبت شد'
            })
        else:
            return jsonify({'error': 'Failed to save order'}), 500
            
    except Exception as e:
        print(f"Error in submit_order: {e}")
        return jsonify({'error': 'Server error'}), 500

@app.route('/orders_report')
def orders_report():
    """گزارش سفارشات"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # بارگذاری سفارشات
    orders_df = load_orders_from_excel()
    if orders_df is None or len(orders_df) == 0:
        orders = []
    else:
        # فیلتر بر اساس بازاریاب (اگر ادمین نیست)
        if session['user_info']['Typev'] != 'admin':
            bazaryab_code = session['user_info']['Codev']
            my_orders = orders_df[orders_df['BazaryabCode'] == bazaryab_code]
        else:
            my_orders = orders_df
        
        # مرتب‌سازی بر اساس تاریخ (جدیدترین اول)
        my_orders = my_orders.sort_values(['OrderDate', 'OrderTime'], ascending=[False, False])
        orders = my_orders.to_dict('records')
    
    return render_template('orders_report.html', orders=orders, user=session['user_info'])

# این کد را به فایل app.py اضافه کنید

@app.route('/get_salesperson_brand_detail', methods=['POST'])
def get_salesperson_brand_detail():
    """گزارش تفصیلی یک بازاریاب از یک برند خاص"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط ادمین
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # دریافت داده‌های POST
        data = request.get_json()
        salesperson_code = data.get('salesperson_code', '').strip()
        brand_name = data.get('brand_name', '').strip()
        date_from = data.get('date_from', '').strip()
        date_to = data.get('date_to', '').strip()
        date_type = data.get('date_type', 'jalali')
        
        if not salesperson_code or not brand_name or not date_from or not date_to:
            return jsonify({'error': 'همه فیلدها الزامی هستند'}), 400
        
        print(f"🔍 Detailed analysis: {salesperson_code} - {brand_name} ({date_from} to {date_to})")
        
        # تبدیل تاریخ به میلادی اگر لازم باشه
        if date_type == 'jalali':
            date_from_gregorian = jalali_to_gregorian(date_from)
            date_to_gregorian = jalali_to_gregorian(date_to)
            
            if not date_from_gregorian or not date_to_gregorian:
                return jsonify({'error': 'فرمت تاریخ نامعتبر است'}), 400
        else:
            date_from_gregorian = date_from
            date_to_gregorian = date_to
        
        # بارگذاری داده‌ها
        products_df = load_products_from_excel()
        customers_df = load_customers_from_excel()
        sales_df = load_sales_from_excel()
        users_df = load_users_from_excel()
        
        if products_df is None or customers_df is None or sales_df is None or users_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل‌ها'}), 500
        
        # پیدا کردن نام بازاریاب
        salesperson_info = users_df[users_df['Codev'] == salesperson_code]
        if salesperson_info.empty:
            return jsonify({'error': 'بازاریاب یافت نشد'}), 404
        
        salesperson_name = salesperson_info.iloc[0]['Namev']
        
        # پیدا کردن همه کالاهای این برند
        brand_products = products_df[products_df['Brand'] == brand_name]
        if brand_products.empty:
            return jsonify({'error': 'هیچ کالایی برای این برند یافت نشد'}), 404
        
        print(f"📦 Found {len(brand_products)} products for brand {brand_name}")
        
        # پیدا کردن مشتریان این بازاریاب
        salesperson_customers = customers_df[customers_df['BazaryabCode'] == salesperson_code]
        customer_codes = salesperson_customers['CustomerCode'].tolist()
        
        if not customer_codes:
            return jsonify({'error': 'هیچ مشتری برای این بازاریاب تعریف نشده'}), 404
        
        print(f"👥 Found {len(customer_codes)} customers for salesperson {salesperson_code}")
        
        # فیلتر فروش‌های این بازاریاب
        salesperson_sales = sales_df[sales_df['CustomerCode'].isin(customer_codes)]
        
        # فیلتر بر اساس تاریخ
        def convert_sale_date_to_gregorian(date_value):
            if pd.isna(date_value):
                return None
            date_str = str(date_value).strip()
            if '/' in date_str and len(date_str.split('/')) == 3:
                return jalali_to_gregorian(date_str)
            elif '-' in date_str and len(date_str) == 10:
                return date_str
            return None
        
        # تبدیل تاریخ‌های فروش
        salesperson_sales_copy = salesperson_sales.copy()
        salesperson_sales_copy['InvoiceDateConverted'] = salesperson_sales_copy['InvoiceDate'].apply(convert_sale_date_to_gregorian)
        
        # فیلتر بر اساس بازه زمانی
        filtered_sales = salesperson_sales_copy[
            (salesperson_sales_copy['InvoiceDateConverted'] >= date_from_gregorian) &
            (salesperson_sales_copy['InvoiceDateConverted'] <= date_to_gregorian)
        ]
        
        print(f"💰 Found {len(filtered_sales)} sales records in date range")
        
        # فیلتر فروش‌های این برند
        brand_product_codes = brand_products['ProductCode'].tolist()
        brand_sales = filtered_sales[filtered_sales['ProductCode'].isin(brand_product_codes)]
        
        print(f"🎯 Found {len(brand_sales)} sales for this brand")
        
        # محاسبه فروش هر کالا
        product_sales = {}
        total_brand_sales = 0
        
        for _, sale in brand_sales.iterrows():
            product_code = sale['ProductCode']
            amount = float(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
            quantity = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
            
            if product_code not in product_sales:
                product_sales[product_code] = {'amount': 0, 'quantity': 0}
            
            product_sales[product_code]['amount'] += amount
            product_sales[product_code]['quantity'] += quantity
            total_brand_sales += amount
        
        # تفکیک کالاهای فروخته شده و نشده
        sold_products = []
        unsold_products = []
        
        for _, product in brand_products.iterrows():
            product_code = product['ProductCode']
            product_name = product['ProductName']
            product_price = float(product.get('Price', 0)) if not pd.isna(product.get('Price', 0)) else 0
            product_category = product.get('Category', '')
            
            if product_code in product_sales:
                # کالای فروخته شده
                sales_data = product_sales[product_code]
                percentage = (sales_data['amount'] / total_brand_sales * 100) if total_brand_sales > 0 else 0
                
                sold_products.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'category': product_category,
                    'price': product_price,
                    'total_amount': int(sales_data['amount']),
                    'total_quantity': int(sales_data['quantity']),
                    'percentage': percentage
                })
            else:
                # کالای فروخته نشده
                unsold_products.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'category': product_category,
                    'price': product_price
                })
        
        # مرتب‌سازی کالاهای فروخته شده بر اساس مقدار فروش
        sold_products.sort(key=lambda x: x['total_amount'], reverse=True)
        
        # مرتب‌سازی کالاهای فروخته نشده بر اساس قیمت
        unsold_products.sort(key=lambda x: x['price'], reverse=True)
        
        print(f"✅ Analysis complete: {len(sold_products)} sold, {len(unsold_products)} unsold")
        
        return jsonify({
            'salesperson_code': salesperson_code,
            'salesperson_name': salesperson_name,
            'brand_name': brand_name,
            'sold_products': sold_products,
            'unsold_products': unsold_products,
            'total_sales': int(total_brand_sales),
            'date_from': date_from,
            'date_to': date_to,
            'date_type': date_type,
            'period_info': f"{date_from} تا {date_to}"
        })
        
    except Exception as e:
        print(f"❌ Error in get_salesperson_brand_detail: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500


# این کدها را به فایل app.py اضافه کنید

@app.route('/product_analysis')
def product_analysis():
    """صفحه تحلیل کالایی بازاریاب"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط ادمین میتونه این صفحه رو ببینه
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('product_analysis.html', user=session['user_info'])

@app.route('/get_salespeople_list')
def get_salespeople_list():
    """دریافت لیست همه بازاریابان"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط ادمین
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # بارگذاری کاربران
        users_df = load_users_from_excel()
        if users_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل کاربران'}), 500
        
        # فیلتر بازاریابان (نوع user)
        salespeople = users_df[users_df['Typev'] == 'user']
        
        salespeople_list = []
        for _, sp in salespeople.iterrows():
            salespeople_list.append({
                'code': sp['Codev'],
                'name': sp['Namev']
            })
        
        # مرتب‌سازی بر اساس نام
        salespeople_list.sort(key=lambda x: x['name'])
        
        return jsonify({'salespeople': salespeople_list})
        
    except Exception as e:
        print(f"❌ Error in get_salespeople_list: {str(e)}")
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/get_product_analysis', methods=['POST'])
def get_product_analysis():
    """تحلیل کالایی یک بازاریاب در بازه زمانی مشخص"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط ادمین
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # دریافت داده‌های POST
        data = request.get_json()
        salesperson_code = data.get('salesperson_code', '').strip()
        date_from = data.get('date_from', '').strip()
        date_to = data.get('date_to', '').strip()
        date_type = data.get('date_type', 'jalali')
        
        if not salesperson_code or not date_from or not date_to:
            return jsonify({'error': 'همه فیلدها الزامی هستند'}), 400
        
        print(f"🔍 Product analysis: {salesperson_code} ({date_from} to {date_to})")
        
        # تبدیل تاریخ به میلادی اگر لازم باشه
        if date_type == 'jalali':
            date_from_gregorian = jalali_to_gregorian(date_from)
            date_to_gregorian = jalali_to_gregorian(date_to)
            
            if not date_from_gregorian or not date_to_gregorian:
                return jsonify({'error': 'فرمت تاریخ نامعتبر است'}), 400
        else:
            date_from_gregorian = date_from
            date_to_gregorian = date_to
        
        # بارگذاری داده‌ها
        products_df = load_products_from_excel()
        customers_df = load_customers_from_excel()
        sales_df = load_sales_from_excel()
        users_df = load_users_from_excel()
        
        if products_df is None or customers_df is None or sales_df is None or users_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل‌ها'}), 500
        
        # پیدا کردن نام بازاریاب
        salesperson_info = users_df[users_df['Codev'] == salesperson_code]
        if salesperson_info.empty:
            return jsonify({'error': 'بازاریاب یافت نشد'}), 404
        
        salesperson_name = salesperson_info.iloc[0]['Namev']
        
        # فیلتر فروش‌ها بر اساس تاریخ
        def convert_sale_date_to_gregorian(date_value):
            if pd.isna(date_value):
                return None
            date_str = str(date_value).strip()
            if '/' in date_str and len(date_str.split('/')) == 3:
                return jalali_to_gregorian(date_str)
            elif '-' in date_str and len(date_str) == 10:
                return date_str
            return None
        
        # تبدیل تاریخ‌های فروش
        sales_df_copy = sales_df.copy()
        sales_df_copy['InvoiceDateConverted'] = sales_df_copy['InvoiceDate'].apply(convert_sale_date_to_gregorian)
        
        # فیلتر بر اساس بازه زمانی
        filtered_sales = sales_df_copy[
            (sales_df_copy['InvoiceDateConverted'] >= date_from_gregorian) &
            (sales_df_copy['InvoiceDateConverted'] <= date_to_gregorian)
        ]
        
        print(f"📊 Found {len(filtered_sales)} sales in date range")
        
        # پیدا کردن مشتریان این بازاریاب
        salesperson_customers = customers_df[customers_df['BazaryabCode'] == salesperson_code]
        customer_codes = salesperson_customers['CustomerCode'].tolist()
        
        # فروش‌های این بازاریاب
        salesperson_sales = filtered_sales[filtered_sales['CustomerCode'].isin(customer_codes)]
        
        # فروش‌های سایر بازاریابان
        other_sales = filtered_sales[~filtered_sales['CustomerCode'].isin(customer_codes)]
        
        print(f"👤 Salesperson sales: {len(salesperson_sales)}")
        print(f"👥 Other sales: {len(other_sales)}")
        
        # محاسبه فروش هر کالا برای این بازاریاب
        salesperson_product_sales = {}
        total_salesperson_sales = 0
        
        for _, sale in salesperson_sales.iterrows():
            product_code = sale['ProductCode']
            amount = float(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
            quantity = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
            
            if product_code not in salesperson_product_sales:
                salesperson_product_sales[product_code] = {'amount': 0, 'quantity': 0}
            
            salesperson_product_sales[product_code]['amount'] += amount
            salesperson_product_sales[product_code]['quantity'] += quantity
            total_salesperson_sales += amount
        
        # محاسبه فروش هر کالا برای سایر بازاریابان
        other_product_sales = {}
        for _, sale in other_sales.iterrows():
            customer_code = sale['CustomerCode']
            product_code = sale['ProductCode']
            amount = float(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
            quantity = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
            
            # پیدا کردن بازاریاب این مشتری
            customer_info = customers_df[customers_df['CustomerCode'] == customer_code]
            if not customer_info.empty:
                other_salesperson_code = customer_info.iloc[0]['BazaryabCode']
                other_salesperson_info = users_df[users_df['Codev'] == other_salesperson_code]
                other_salesperson_name = other_salesperson_info.iloc[0]['Namev'] if not other_salesperson_info.empty else 'نامشخص'
            else:
                other_salesperson_name = 'نامشخص'
            
            if product_code not in other_product_sales:
                other_product_sales[product_code] = {}
            
            if other_salesperson_name not in other_product_sales[product_code]:
                other_product_sales[product_code][other_salesperson_name] = {'amount': 0, 'quantity': 0}
            
            other_product_sales[product_code][other_salesperson_name]['amount'] += amount
            other_product_sales[product_code][other_salesperson_name]['quantity'] += quantity
        
        # تفکیک کالاها
        sold_by_salesperson = []
        sold_by_others = []
        not_sold = []
        
        for _, product in products_df.iterrows():
            product_code = product['ProductCode']
            product_name = product['ProductName']
            brand = product.get('Brand', '')
            category = product.get('Category', '')
            price = float(product.get('Price', 0)) if not pd.isna(product.get('Price', 0)) else 0
            
            # کالاهای فروخته شده توسط این بازاریاب
            if product_code in salesperson_product_sales:
                sales_data = salesperson_product_sales[product_code]
                sold_by_salesperson.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'brand': brand,
                    'category': category,
                    'price': price,
                    'total_amount': int(sales_data['amount']),
                    'total_quantity': int(sales_data['quantity'])
                })
            
            # کالاهای فروخته شده توسط سایر بازاریابان
            elif product_code in other_product_sales:
                other_sales_list = []
                total_lost_amount = 0
                
                for sp_name, sales_data in other_product_sales[product_code].items():
                    other_sales_list.append({
                        'salesperson_name': sp_name,
                        'amount': int(sales_data['amount']),
                        'quantity': int(sales_data['quantity'])
                    })
                    total_lost_amount += sales_data['amount']
                
                # مرتب‌سازی بر اساس مقدار فروش
                other_sales_list.sort(key=lambda x: x['amount'], reverse=True)
                
                sold_by_others.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'brand': brand,
                    'category': category,
                    'price': price,
                    'total_lost_amount': int(total_lost_amount),
                    'other_sales': other_sales_list
                })
            
            # کالاهای فروخته نشده
            else:
                not_sold.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'brand': brand,
                    'category': category,
                    'price': price
                })
        
        # مرتب‌سازی
        sold_by_salesperson.sort(key=lambda x: x['total_amount'], reverse=True)
        sold_by_others.sort(key=lambda x: x['total_lost_amount'], reverse=True)
        not_sold.sort(key=lambda x: x['price'], reverse=True)
        
        print(f"✅ Analysis complete:")
        print(f"   Sold by salesperson: {len(sold_by_salesperson)}")
        print(f"   Sold by others: {len(sold_by_others)}")
        print(f"   Not sold: {len(not_sold)}")
        
        return jsonify({
            'salesperson_code': salesperson_code,
            'salesperson_name': salesperson_name,
            'sold_by_salesperson': sold_by_salesperson,
            'sold_by_others': sold_by_others,
            'not_sold': not_sold,
            'total_sales': int(total_salesperson_sales),
            'total_products': len(products_df),
            'date_from': date_from,
            'date_to': date_to,
            'date_type': date_type,
            'period_info': f"{date_from} تا {date_to}"
        })
        
    except Exception as e:
        print(f"❌ Error in get_product_analysis: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500


# این کدها را به فایل app.py اضافه کنید

@app.route('/my_product_analysis')
def my_product_analysis():
    """صفحه تحلیل عملکرد فروش برای کاربر عادی"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط کاربران عادی میتونن این صفحه رو ببینن
    if session['user_info']['Typev'] != 'user':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('my_product_analysis.html', user=session['user_info'])

@app.route('/get_my_product_analysis', methods=['POST'])
def get_my_product_analysis():
    """تحلیل کالایی برای کاربر لاگین شده"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط کاربر عادی
        if session['user_info']['Typev'] != 'user':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # دریافت داده‌های POST
        data = request.get_json()
        date_from = data.get('date_from', '').strip()
        date_to = data.get('date_to', '').strip()
        date_type = data.get('date_type', 'jalali')
        stock_filter = data.get('stock_filter', 'in_stock')  # پیش‌فرض: فقط موجود
        
        if not date_from or not date_to:
            return jsonify({'error': 'بازه زمانی الزامی است'}), 400
        
        # کد بازاریاب از session
        salesperson_code = session['user_info']['Codev']
        salesperson_name = session['user_info']['Namev']
        
        print(f"🔍 My product analysis: {salesperson_name} ({salesperson_code}) - {date_from} to {date_to}")
        print(f"📦 Stock filter: {stock_filter}")
        
        # تبدیل تاریخ به میلادی اگر لازم باشه
        if date_type == 'jalali':
            date_from_gregorian = jalali_to_gregorian(date_from)
            date_to_gregorian = jalali_to_gregorian(date_to)
            
            if not date_from_gregorian or not date_to_gregorian:
                return jsonify({'error': 'فرمت تاریخ نامعتبر است'}), 400
        else:
            date_from_gregorian = date_from
            date_to_gregorian = date_to
        
        # بارگذاری داده‌ها
        products_df = load_products_from_excel()
        customers_df = load_customers_from_excel()
        sales_df = load_sales_from_excel()
        users_df = load_users_from_excel()
        
        if products_df is None or customers_df is None or sales_df is None or users_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل‌ها'}), 500
        
        # بارگذاری ترتیب برندها برای مرتب‌سازی
        brand_order = load_brand_order_from_excel()
        brand_radif = {}
        if brand_order:
            for index, brand in enumerate(brand_order):
                brand_radif[brand] = index + 1
        
        print(f"📋 Brand order loaded: {len(brand_radif)} brands")
        
        # فیلتر فروش‌ها بر اساس تاریخ
        def convert_sale_date_to_gregorian(date_value):
            if pd.isna(date_value):
                return None
            date_str = str(date_value).strip()
            if '/' in date_str and len(date_str.split('/')) == 3:
                return jalali_to_gregorian(date_str)
            elif '-' in date_str and len(date_str) == 10:
                return date_str
            return None
        
        # تبدیل تاریخ‌های فروش
        sales_df_copy = sales_df.copy()
        sales_df_copy['InvoiceDateConverted'] = sales_df_copy['InvoiceDate'].apply(convert_sale_date_to_gregorian)
        
        # فیلتر بر اساس بازه زمانی
        filtered_sales = sales_df_copy[
            (sales_df_copy['InvoiceDateConverted'] >= date_from_gregorian) &
            (sales_df_copy['InvoiceDateConverted'] <= date_to_gregorian)
        ]
        
        print(f"📊 Found {len(filtered_sales)} sales in date range")
        
        # پیدا کردن مشتریان این بازاریاب
        salesperson_customers = customers_df[customers_df['BazaryabCode'] == salesperson_code]
        customer_codes = salesperson_customers['CustomerCode'].tolist()
        
        print(f"👥 Found {len(customer_codes)} customers for this salesperson")
        
        # فروش‌های این بازاریاب
        my_sales = filtered_sales[filtered_sales['CustomerCode'].isin(customer_codes)]
        
        # فروش‌های سایر بازاریابان
        other_sales = filtered_sales[~filtered_sales['CustomerCode'].isin(customer_codes)]
        
        print(f"👤 My sales: {len(my_sales)}")
        print(f"👥 Other sales: {len(other_sales)}")
        
        # محاسبه فروش هر کالا برای این بازاریاب
        my_product_sales = {}
        total_my_sales = 0
        
        for _, sale in my_sales.iterrows():
            product_code = sale['ProductCode']
            amount = float(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
            quantity = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
            
            if product_code not in my_product_sales:
                my_product_sales[product_code] = {'amount': 0, 'quantity': 0}
            
            my_product_sales[product_code]['amount'] += amount
            my_product_sales[product_code]['quantity'] += quantity
            total_my_sales += amount
        
        # محاسبه فروش هر کالا برای سایر بازاریابان
        other_product_sales = {}
        for _, sale in other_sales.iterrows():
            customer_code = sale['CustomerCode']
            product_code = sale['ProductCode']
            amount = float(sale.get('TotalAmount', 0)) if not pd.isna(sale.get('TotalAmount', 0)) else 0
            quantity = int(sale.get('Quantity', 0)) if not pd.isna(sale.get('Quantity', 0)) else 0
            
            # پیدا کردن بازاریاب این مشتری
            customer_info = customers_df[customers_df['CustomerCode'] == customer_code]
            if not customer_info.empty:
                other_salesperson_code = customer_info.iloc[0]['BazaryabCode']
                other_salesperson_info = users_df[users_df['Codev'] == other_salesperson_code]
                other_salesperson_name = other_salesperson_info.iloc[0]['Namev'] if not other_salesperson_info.empty else 'نامشخص'
            else:
                other_salesperson_name = 'نامشخص'
            
            if product_code not in other_product_sales:
                other_product_sales[product_code] = {}
            
            if other_salesperson_name not in other_product_sales[product_code]:
                other_product_sales[product_code][other_salesperson_name] = {'amount': 0, 'quantity': 0}
            
            other_product_sales[product_code][other_salesperson_name]['amount'] += amount
            other_product_sales[product_code][other_salesperson_name]['quantity'] += quantity
        
        # تفکیک کالاها با مرتب‌سازی بر اساس برند
        sold_by_me = []
        sold_by_others = []
        not_sold = []
        total_lost_opportunities = 0
        
        for _, product in products_df.iterrows():
            product_code = product['ProductCode']
            product_name = product['ProductName']
            brand = product.get('Brand', '')
            category = product.get('Category', '')
            price = float(product.get('Price', 0)) if not pd.isna(product.get('Price', 0)) else 0
            stock = int(product.get('Stock', 0)) if not pd.isna(product.get('Stock', 0)) else 0
            
            # دریافت ردیف برند برای مرتب‌سازی
            radif = brand_radif.get(brand, 999)
            
            # کالاهای فروخته شده توسط من
            if product_code in my_product_sales:
                sales_data = my_product_sales[product_code]
                sold_by_me.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'brand': brand,
                    'category': category,
                    'price': price,
                    'radif': radif,
                    'total_amount': int(sales_data['amount']),
                    'total_quantity': int(sales_data['quantity'])
                })
            
            # کالاهای فروخته شده توسط سایر بازاریابان
            elif product_code in other_product_sales:
                other_sales_list = []
                total_lost_amount = 0
                
                for sp_name, sales_data in other_product_sales[product_code].items():
                    other_sales_list.append({
                        'salesperson_name': sp_name,
                        'amount': int(sales_data['amount']),
                        'quantity': int(sales_data['quantity'])
                    })
                    total_lost_amount += sales_data['amount']
                
                # مرتب‌سازی بر اساس مقدار فروش
                other_sales_list.sort(key=lambda x: x['amount'], reverse=True)
                
                sold_by_others.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'brand': brand,
                    'category': category,
                    'price': price,
                    'radif': radif,
                    'total_lost_amount': int(total_lost_amount),
                    'other_sales': other_sales_list
                })
                
                total_lost_opportunities += total_lost_amount
            
            # کالاهای فروخته نشده (با موجودی و قیمت)
            else:
                # فیلتر بر اساس موجودی
                if stock_filter == 'in_stock' and stock <= 0:
                    continue  # رد کردن کالاهای بدون موجودی
                
                not_sold.append({
                    'product_code': product_code,
                    'product_name': product_name,
                    'brand': brand,
                    'category': category,
                    'price': price,
                    'stock': stock,
                    'radif': radif
                })
        
        # مرتب‌سازی بر اساس ردیف برند و سپس مقدار فروش
        sold_by_me.sort(key=lambda x: (x['radif'], -x['total_amount']))
        sold_by_others.sort(key=lambda x: (x['radif'], -x['total_lost_amount']))
        not_sold.sort(key=lambda x: (x['radif'], -x['price']))
        
        print(f"✅ My analysis complete:")
        print(f"   Sold by me: {len(sold_by_me)}")
        print(f"   Sold by others: {len(sold_by_others)}")
        print(f"   Not sold: {len(not_sold)} ({'in stock only' if stock_filter == 'in_stock' else 'all products'})")
        print(f"   Total lost opportunities: {total_lost_opportunities:,}")
        
        return jsonify({
            'salesperson_code': salesperson_code,
            'salesperson_name': salesperson_name,
            'sold_by_me': sold_by_me,
            'sold_by_others': sold_by_others,
            'not_sold': not_sold,
            'total_sales': int(total_my_sales),
            'total_lost_opportunities': int(total_lost_opportunities),
            'total_products': len(products_df),
            'stock_filter': stock_filter,
            'date_from': date_from,
            'date_to': date_to,
            'date_type': date_type,
            'period_info': f"{date_from} تا {date_to}"
        })
        
    except Exception as e:
        print(f"❌ Error in get_my_product_analysis: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500
        
@app.route('/all_reports')
def all_reports():
    """گزارش کلی همه مراجعات (برای ادمین یا گزارش بازاریاب)"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # بارگذاری داده‌ها
    customers_df = load_customers_from_excel()
    visits_df = load_visits_from_excel()
    users_df = load_users_from_excel()
    
    if customers_df is None or visits_df is None or users_df is None:
        flash('خطا در بارگذاری اطلاعات!', 'error')
        return redirect(url_for('index'))
    
    # فیلتر بر اساس نوع کاربر
    bazaryab_code = session['user_info']['Codev']
    if session['user_info']['Typev'] != 'admin':
        # بازاریاب فقط مراجعات خودش رو می‌بینه
        my_visits = visits_df[visits_df['BazaryabCode'] == bazaryab_code]
        my_customers = customers_df[customers_df['BazaryabCode'] == bazaryab_code]
    else:
        # ادمین همه رو می‌بینه
        my_visits = visits_df
        my_customers = customers_df
    
    # ترکیب اطلاعات
    report_data = []
    for _, visit in my_visits.iterrows():
        # پیدا کردن اطلاعات مشتری
        customer = my_customers[my_customers['CustomerCode'] == visit['CustomerCode']]
        customer_name = customer.iloc[0]['CustomerName'] if not customer.empty else 'نامشخص'
        
        # پیدا کردن اطلاعات بازاریاب
        bazaryab = users_df[users_df['Codev'] == visit['BazaryabCode']]
        bazaryab_name = bazaryab.iloc[0]['Namev'] if not bazaryab.empty else 'نامشخص'
        
        report_data.append({
            'VisitCode': visit['VisitCode'],
            'CustomerCode': visit['CustomerCode'],
            'CustomerName': customer_name,
            'BazaryabCode': visit['BazaryabCode'],
            'BazaryabName': bazaryab_name,
            'VisitDate': visit['VisitDate'],
            'VisitTime': visit['VisitTime']
        })
    
    # مرتب‌سازی بر اساس تاریخ (جدیدترین اول)
    report_data.sort(key=lambda x: (x['VisitDate'], x['VisitTime']), reverse=True)
    
    # آمار کلی
    total_visits = len(report_data)
    unique_customers = len(set([r['CustomerCode'] for r in report_data]))
    
    return render_template('all_reports.html',
                         reports=report_data,
                         total_visits=total_visits,
                         unique_customers=unique_customers,
                         user=session['user_info'])

# 2. توابع مدیریت فایل آزمون:

def create_exam_file_if_not_exists():
    """ایجاد فایل azmon.xlsx اگر وجود نداشته باشد"""
    if not os.path.exists(EXAMS_FILE):
        try:
            df = pd.DataFrame(columns=[
                'ExamCode', 'ExamName', 'BrandName', 'CreatedDate', 'CreatedTime', 'CreatedBy'
            ])
            df.to_excel(EXAMS_FILE, sheet_name='list', index=False)
            print("✅ فایل azmon.xlsx ایجاد شد")
            return True
        except Exception as e:
            print(f"❌ خطا در ایجاد فایل آزمون: {e}")
            return False
    return True

def load_exams_from_excel():
    """بارگذاری آزمون‌ها از فایل Excel"""
    try:
        # ابتدا مطمئن شویم فایل وجود دارد
        create_exam_file_if_not_exists()
        
        if not os.path.exists(EXAMS_FILE):
            return pd.DataFrame(columns=[
                'ExamCode', 'ExamName', 'ExamType', 'BrandName', 'Description',
                'CreatedDate', 'CreatedTime', 'CreatedBy'
            ])
            
        df = pd.read_excel(EXAMS_FILE, sheet_name='list')
        print("✅ فایل آزمون با موفقیت بارگذاری شد")
        
        # اگر ستون‌های جدید وجود ندارند، اضافه کن
        required_columns = ['ExamCode', 'ExamName', 'ExamType', 'BrandName', 'Description', 
                          'CreatedDate', 'CreatedTime', 'CreatedBy']
        
        for col in required_columns:
            if col not in df.columns:
                df[col] = ''
                print(f"➕ ستون {col} اضافه شد")
        
        # پاک کردن فاصله‌های اضافی
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        return df
    except Exception as e:
        print(f"❌ خطا در بارگذاری فایل آزمون: {e}")
        # در صورت خطا، یک DataFrame خالی برگردان
        return pd.DataFrame(columns=[
            'ExamCode', 'ExamName', 'ExamType', 'BrandName', 'Description',
            'CreatedDate', 'CreatedTime', 'CreatedBy'
        ])

def save_exams_to_excel(df):
    """ذخیره آزمون‌ها در فایل Excel"""
    try:
        # استفاده از ExcelWriter برای کنترل بهتر
        with pd.ExcelWriter(EXAMS_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='list', index=False)
        print("✅ فایل آزمون با موفقیت ذخیره شد")
        return True
    except Exception as e:
        print(f"❌ خطا در ذخیره فایل آزمون: {e}")
        return False

def generate_exam_code():
    """تولید کد آزمون منحصر به فرد"""
    try:
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        date_str = jalali_now.strftime('%Y%m%d')
        
        # بررسی آخرین کد آزمون امروز
        exams_df = load_exams_from_excel()
        if exams_df is not None and len(exams_df) > 0:
            today_exams = exams_df[exams_df['ExamCode'].str.contains(f'EX-{date_str}', na=False)]
            if len(today_exams) > 0:
                last_number = len(today_exams) + 1
            else:
                last_number = 1
        else:
            last_number = 1
        
        exam_code = f"EX-{date_str}{last_number:03d}"
        print(f"🆕 کد آزمون جدید: {exam_code}")
        return exam_code
        
    except Exception as e:
        print(f"❌ خطا در تولید کد آزمون: {e}")
        # در صورت خطا، از timestamp استفاده کن
        fallback_code = f"EX-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        return fallback_code

# 3. Route های آزمون:

@app.route('/exam_management')
def exam_management():
    """صفحه مدیریت آزمون - فقط برای ادمین"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط ادمین می‌تونه این صفحه رو ببینه
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    # ایجاد فایل اگر وجود نداشته باشد
    create_exam_file_if_not_exists()
    
    # استفاده از فایل template
    return render_template('exam_management.html', user=session['user_info'])

@app.route('/create_exam_simple', methods=['POST'])
def create_exam_simple():
    """ایجاد آزمون ساده"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        data = request.get_json()
        exam_name = data.get('exam_name', '').strip()
        brand_name = data.get('brand_name', '').strip()
        
        if not exam_name or not brand_name:
            return jsonify({'error': 'نام آزمون و برند الزامی است'}), 400
        
        print(f"🆕 Creating exam: {exam_name} for brand: {brand_name}")
        
        # تولید کد آزمون
        exam_code = generate_exam_code()
        
        # تاریخ و ساعت فعلی
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        created_date = jalali_now.strftime('%Y/%m/%d')
        created_time = now.strftime('%H:%M')
        
        # بارگذاری آزمون‌های موجود
        exams_df = load_exams_from_excel()
        
        # ایجاد رکورد جدید
        new_exam = pd.DataFrame([{
            'ExamCode': exam_code,
            'ExamName': exam_name,
            'BrandName': brand_name,
            'CreatedDate': created_date,
            'CreatedTime': created_time,
            'CreatedBy': session['user_info']['Codev']
        }])
        
        # اضافه کردن به DataFrame موجود
        exams_df = pd.concat([exams_df, new_exam], ignore_index=True)
        
        # ذخیره فایل
        if save_exams_to_excel(exams_df):
            print(f"✅ Exam created successfully: {exam_code}")
            return jsonify({
                'success': True,
                'exam_code': exam_code,
                'message': 'آزمون با موفقیت ایجاد شد'
            })
        else:
            return jsonify({'error': 'خطا در ذخیره آزمون'}), 500
        
    except Exception as e:
        print(f"❌ Error in create_exam_simple: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/get_exams_simple')
def get_exams_simple():
    """دریافت لیست آزمون‌ها - ساده"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # بارگذاری آزمون‌ها
        exams_df = load_exams_from_excel()
        
        if len(exams_df) == 0:
            return jsonify({'exams': []})
        
        # مرتب‌سازی بر اساس تاریخ (جدیدترین اول)
        exams_df = exams_df.sort_values(['CreatedDate', 'CreatedTime'], ascending=[False, False])
        
        exams = []
        for _, row in exams_df.iterrows():
            exams.append({
                'code': row['ExamCode'],
                'name': row['ExamName'],
                'brand': row['BrandName'],
                'date': f"{row['CreatedDate']} {row.get('CreatedTime', '')}"
            })
        
        return jsonify({'exams': exams})
        
    except Exception as e:
        print(f"❌ Error in get_exams_simple: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

# این کد را به فایل app.py اضافه کنید

@app.route('/get_brands_for_exam')
def get_brands_for_exam():
    """دریافت لیست برندها برای استفاده در آزمون"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # بارگذاری محصولات
        products_df = load_products_from_excel()
        if products_df is None:
            return jsonify({'error': 'فایل محصولات یافت نشد'}), 500
        
        # دریافت لیست برندها (حذف تکراری و مرتب‌سازی)
        brands = sorted(products_df['Brand'].unique().tolist())
        
        # حذف مقادیر خالی یا NaN
        brands = [brand for brand in brands if str(brand) not in ['', 'nan', 'None']]
        
        print(f"🏷️ Brands found for exam: {brands}")
        
        return jsonify({
            'success': True,
            'brands': brands
        })
        
    except Exception as e:
        print(f"❌ Error in get_brands_for_exam: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/get_exam_list')
def get_exam_list():
    """دریافت لیست آزمون‌های ایجاد شده"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # بارگذاری آزمون‌ها
        exams_df = load_exams_from_excel()
        
        if len(exams_df) == 0:
            return jsonify({'exams': []})
        
        # مرتب‌سازی بر اساس تاریخ (جدیدترین اول)
        exams_df = exams_df.sort_values(['CreatedDate', 'CreatedTime'], ascending=[False, False])
        
        exams = []
        for _, row in exams_df.iterrows():
            exam_data = {
                'exam_code': row.get('ExamCode', ''),
                'exam_name': row.get('ExamName', ''),
                'exam_type': row.get('ExamType', 'عمومی'),
                'brand_name': row.get('BrandName', ''),
                'description': row.get('Description', ''),
                'created_date': f"{row.get('CreatedDate', '')} {row.get('CreatedTime', '')}"
            }
            exams.append(exam_data)
        
        return jsonify({
            'success': True,
            'exams': exams
        })
        
    except Exception as e:
        print(f"❌ Error in get_exam_list: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500


@app.route('/create_exam', methods=['POST'])
def create_exam():
    """ایجاد آزمون جدید با نوع آزمون"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        data = request.get_json()
        exam_name = data.get('exam_name', '').strip()
        exam_type = data.get('exam_type', '').strip()
        brand_name = data.get('brand_name', '').strip()
        description = data.get('description', '').strip()
        
        if not exam_name or not exam_type or not brand_name:
            return jsonify({'error': 'نام آزمون، نوع آزمون و برند الزامی است'}), 400
        
        print(f"🆕 Creating exam: {exam_name} ({exam_type}) for brand: {brand_name}")
        
        # تولید کد آزمون
        exam_code = generate_exam_code()
        
        # تاریخ و ساعت فعلی
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        created_date = jalali_now.strftime('%Y/%m/%d')
        created_time = now.strftime('%H:%M')
        
        # بارگذاری آزمون‌های موجود
        exams_df = load_exams_from_excel()
        
        # ایجاد رکورد جدید
        new_exam = pd.DataFrame([{
            'ExamCode': exam_code,
            'ExamName': exam_name,
            'ExamType': exam_type,
            'BrandName': brand_name,
            'Description': description,
            'CreatedDate': created_date,
            'CreatedTime': created_time,
            'CreatedBy': session['user_info']['Codev']
        }])
        
        # اضافه کردن به DataFrame موجود
        exams_df = pd.concat([exams_df, new_exam], ignore_index=True)
        
        # ذخیره فایل
        if save_exams_to_excel(exams_df):
            print(f"✅ Exam created successfully: {exam_code}")
            return jsonify({
                'success': True,
                'exam_code': exam_code,
                'message': 'آزمون با موفقیت ایجاد شد'
            })
        else:
            return jsonify({'error': 'خطا در ذخیره آزمون'}), 500
        
    except Exception as e:
        print(f"❌ Error in create_exam: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

# این کدها را به فایل app.py اضافه کنید

@app.route('/user_exam_list')
def user_exam_list():
    """صفحه لیست آزمون‌ها برای کاربران عادی"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط کاربران عادی می‌توانند این صفحه را ببینند
    if session['user_info']['Typev'] != 'user':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('user_exam_list.html', user=session['user_info'])

@app.route('/get_user_exams')
def get_user_exams():
    """دریافت لیست آزمون‌های موجود برای کاربران عادی"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'user':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        print(f"🎯 Loading exams for user: {session['user_info']['Namev']}")
        
        # بارگذاری آزمون‌ها
        exams_df = load_exams_from_excel()
        
        if len(exams_df) == 0:
            return jsonify({
                'success': True,
                'exams': [],
                'message': 'هیچ آزمونی موجود نیست'
            })
        
        # مرتب‌سازی بر اساس تاریخ (جدیدترین اول)
        exams_df = exams_df.sort_values(['CreatedDate', 'CreatedTime'], ascending=[False, False])
        
        # تبدیل به لیست برای نمایش به کاربر
        user_exams = []
        for _, row in exams_df.iterrows():
            exam_data = {
                'exam_code': row.get('ExamCode', ''),
                'exam_name': row.get('ExamName', ''),
                'exam_type': row.get('ExamType', 'عمومی'),
                'brand_name': row.get('BrandName', ''),
                'description': row.get('Description', ''),
                'created_date': row.get('CreatedDate', ''),
                'created_time': row.get('CreatedTime', ''),
                'status': 'available'  # فعلاً همه آزمون‌ها در دسترس هستند
            }
            user_exams.append(exam_data)
        
        print(f"📋 Found {len(user_exams)} exams for user")
        
        return jsonify({
            'success': True,
            'exams': user_exams,
            'total_count': len(user_exams)
        })
        
    except Exception as e:
        print(f"❌ Error in get_user_exams: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500


@app.route('/exam_info/<exam_code>')
def exam_info(exam_code):
    """نمایش جزئیات آزمون"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    try:
        # بارگذاری اطلاعات آزمون
        exams_df = load_exams_from_excel()
        exam = exams_df[exams_df['ExamCode'] == exam_code]
        
        if exam.empty:
            return jsonify({'error': 'آزمون یافت نشد'}), 404
        
        exam_info = exam.iloc[0]
        
        return jsonify({
            'success': True,
            'exam': {
                'exam_code': exam_info.get('ExamCode', ''),
                'exam_name': exam_info.get('ExamName', ''),
                'exam_type': exam_info.get('ExamType', 'عمومی'),
                'brand_name': exam_info.get('BrandName', ''),
                'description': exam_info.get('Description', ''),
                'created_date': exam_info.get('CreatedDate', ''),
                'created_time': exam_info.get('CreatedTime', ''),
                'created_by': exam_info.get('CreatedBy', '')
            }
        })
        
    except Exception as e:
        print(f"❌ Error in exam_info: {str(e)}")
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500
# این کدها را به فایل app.py اضافه کنید

def save_exam_result_to_excel(result_data):
    """ذخیره نتیجه آزمون در فایل azmon.xlsx شیت azmon"""
    try:
        # بررسی وجود فایل و شیت
        if os.path.exists(EXAMS_FILE):
            with pd.ExcelFile(EXAMS_FILE) as xls:
                if 'azmon' in xls.sheet_names:
                    # بارگذاری داده‌های موجود
                    results_df = pd.read_excel(EXAMS_FILE, sheet_name='azmon')
                else:
                    # ایجاد DataFrame جدید
                    results_df = pd.DataFrame(columns=[
                        'ExamResultCode', 'ExamCode', 'BazaryabCode', 'BazaryabName',
                        'ExamDate', 'ExamTime', 'TotalQuestions', 'CorrectAnswers', 
                        'WrongAnswers', 'Score', 'Percentage', 'TimeTaken', 'ExamType',
                        'BrandName', 'ResultDescription'
                    ])
        else:
            # ایجاد DataFrame جدید
            results_df = pd.DataFrame(columns=[
                'ExamResultCode', 'ExamCode', 'BazaryabCode', 'BazaryabName',
                'ExamDate', 'ExamTime', 'TotalQuestions', 'CorrectAnswers', 
                'WrongAnswers', 'Score', 'Percentage', 'TimeTaken', 'ExamType',
                'BrandName', 'ResultDescription'
            ])
        
        # ایجاد رکورد جدید
        new_result = pd.DataFrame([result_data])
        results_df = pd.concat([results_df, new_result], ignore_index=True)
        
        # ذخیره در فایل
        if os.path.exists(EXAMS_FILE):
            # بارگذاری سایر شیت‌ها
            with pd.ExcelFile(EXAMS_FILE) as xls:
                sheets_dict = {}
                for sheet_name in xls.sheet_names:
                    if sheet_name != 'azmon':
                        sheets_dict[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
            
            # ذخیره همه شیت‌ها
            with pd.ExcelWriter(EXAMS_FILE, engine='openpyxl') as writer:
                for sheet_name, df in sheets_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                results_df.to_excel(writer, sheet_name='azmon', index=False)
        else:
            # ایجاد فایل جدید
            with pd.ExcelWriter(EXAMS_FILE, engine='openpyxl') as writer:
                results_df.to_excel(writer, sheet_name='azmon', index=False)
                # ایجاد شیت list خالی
                pd.DataFrame().to_excel(writer, sheet_name='list', index=False)
        
        print(f"✅ Exam result saved: {result_data['ExamResultCode']}")
        return True
        
    except Exception as e:
        print(f"❌ Error saving exam result: {e}")
        import traceback
        traceback.print_exc()
        return False

def generate_exam_result_code():
    """تولید کد نتیجه آزمون منحصر به فرد"""
    try:
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        date_str = jalali_now.strftime('%Y%m%d')
        time_str = now.strftime('%H%M%S')
        
        return f"ER-{date_str}{time_str}"
        
    except Exception as e:
        print(f"❌ Error generating result code: {e}")
        return f"ER-{datetime.now().strftime('%Y%m%d%H%M%S')}"

@app.route('/take_exam/<exam_code>')
def take_exam(exam_code):
    """ورود به آزمون - تشخیص نوع آزمون"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط کاربران عادی می‌توانند آزمون بدهند
    if session['user_info']['Typev'] != 'user':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    try:
        # بارگذاری اطلاعات آزمون
        exams_df = load_exams_from_excel()
        exam = exams_df[exams_df['ExamCode'] == exam_code]
        
        if exam.empty:
            flash('آزمون مورد نظر یافت نشد!', 'error')
            return redirect(url_for('user_exam_list'))
        
        exam_info = exam.iloc[0].to_dict()
        exam_type = exam_info.get('ExamType', 'عمومی')
        
        print(f"🎯 User {session['user_info']['Namev']} starting exam: {exam_code} (Type: {exam_type})")
        
        # تشخیص نوع آزمون و هدایت به صفحه مناسب
        if exam_type == 'محصولات':
            return render_template('product_exam.html', 
                                 exam=exam_info, 
                                 user=session['user_info'])
        else:
            # سایر انواع آزمون (فعلاً placeholder)
            return render_template('take_exam.html', 
                                 exam=exam_info, 
                                 user=session['user_info'])
        
    except Exception as e:
        print(f"❌ Error in take_exam: {str(e)}")
        flash('خطا در بارگذاری آزمون!', 'error')
        return redirect(url_for('user_exam_list'))

@app.route('/get_exam_products/<exam_code>')
def get_exam_products(exam_code):
    """دریافت محصولات برند آزمون برای آزمون محصولات"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'user':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        print(f"📦 Loading products for exam: {exam_code}")
        
        # بارگذاری اطلاعات آزمون
        exams_df = load_exams_from_excel()
        exam = exams_df[exams_df['ExamCode'] == exam_code]
        
        if exam.empty:
            return jsonify({'error': 'آزمون یافت نشد'}), 404
        
        exam_info = exam.iloc[0]
        brand_name = exam_info.get('BrandName', '')
        
        if not brand_name:
            return jsonify({'error': 'برند آزمون مشخص نیست'}), 400
        
        # بارگذاری محصولات این برند
        products_df = load_products_from_excel()
        if products_df is None:
            return jsonify({'error': 'فایل محصولات یافت نشد'}), 500
        
        # فیلتر محصولات بر اساس برند
        brand_products = products_df[products_df['Brand'] == brand_name]
        
        if brand_products.empty:
            return jsonify({'error': f'هیچ محصولی برای برند {brand_name} یافت نشد'}), 404
        
        # تبدیل به لیست
        products_list = []
        for _, product in brand_products.iterrows():
            # تنها محصولاتی که دارای عکس هستند (اختیاری)
            products_list.append({
                'ProductCode': product.get('ProductCode', ''),
                'ProductName': product.get('ProductName', ''),
                'Category': product.get('Category', ''),
                'Brand': product.get('Brand', ''),
                'Price': float(product.get('Price', 0)) if not pd.isna(product.get('Price', 0)) else 0,
                'ImageFile': product.get('ImageFile', 'null.jpg'),
                'Description': product.get('Description', '')
            })
        
        print(f"✅ Found {len(products_list)} products for brand {brand_name}")
        
        return jsonify({
            'success': True,
            'products': products_list,
            'brand_name': brand_name,
            'exam_code': exam_code
        })
        
    except Exception as e:
        print(f"❌ Error in get_exam_products: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/submit_product_exam', methods=['POST'])
def submit_product_exam():
    """ثبت نتیجه آزمون محصولات"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'user':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        data = request.get_json()
        exam_code = data.get('exam_code')
        matches = data.get('matches', {})
        time_taken = data.get('time_taken', 0)
        
        if not exam_code or not matches:
            return jsonify({'error': 'داده‌های آزمون ناقص است'}), 400
        
        print(f"📝 Processing exam submission: {exam_code}")
        print(f"👤 User: {session['user_info']['Namev']}")
        print(f"⏱️ Time taken: {time_taken} seconds")
        
        # بارگذاری اطلاعات آزمون
        exams_df = load_exams_from_excel()
        exam = exams_df[exams_df['ExamCode'] == exam_code]
        
        if exam.empty:
            return jsonify({'error': 'آزمون یافت نشد'}), 404
        
        exam_info = exam.iloc[0]
        brand_name = exam_info.get('BrandName', '')
        
        # بارگذاری محصولات برند
        products_df = load_products_from_excel()
        brand_products = products_df[products_df['Brand'] == brand_name]
        
        # محاسبه نتایج
        total_questions = len(brand_products)
        correct_answers = 0
        wrong_answers = 0
        
        # بررسی پاسخ‌ها
        for target_code, selected_code in matches.items():
            if target_code == selected_code:
                correct_answers += 1
            else:
                wrong_answers += 1
        
        # محاسبه امتیاز
        percentage = (correct_answers / total_questions * 100) if total_questions > 0 else 0
        score = round(percentage)
        
        # تعیین وضعیت و توضیحات
        if percentage >= 80:
            result_description = f"عالی! شما {correct_answers} از {total_questions} محصول را به درستی تشخیص دادید."
        elif percentage >= 60:
            result_description = f"خوب! شما {correct_answers} از {total_questions} محصول را به درستی تشخیص دادید."
        else:
            result_description = f"نیاز به تلاش بیشتر. شما {correct_answers} از {total_questions} محصول را به درستی تشخیص دادید."
        
        # ایجاد کد نتیجه
        result_code = generate_exam_result_code()
        
        # تاریخ و ساعت فعلی
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        exam_date = jalali_now.strftime('%Y/%m/%d')
        exam_time = now.strftime('%H:%M')
        
        # ایجاد داده نتیجه
        result_data = {
            'ExamResultCode': result_code,
            'ExamCode': exam_code,
            'BazaryabCode': session['user_info']['Codev'],
            'BazaryabName': session['user_info']['Namev'],
            'ExamDate': exam_date,
            'ExamTime': exam_time,
            'TotalQuestions': total_questions,
            'CorrectAnswers': correct_answers,
            'WrongAnswers': wrong_answers,
            'Score': score,
            'Percentage': round(percentage, 1),
            'TimeTaken': f"{time_taken // 60}:{time_taken % 60:02d}",
            'ExamType': exam_info.get('ExamType', 'محصولات'),
            'BrandName': brand_name,
            'ResultDescription': result_description
        }
        
        # ذخیره در فایل
        if save_exam_result_to_excel(result_data):
            print(f"✅ Exam result saved successfully for {session['user_info']['Namev']}")
            
            return jsonify({
                'success': True,
                'result_code': result_code,
                'total_questions': total_questions,
                'correct_answers': correct_answers,
                'wrong_answers': wrong_answers,
                'score': score,
                'percentage': round(percentage, 1),
                'time_taken': f"{time_taken // 60}:{time_taken % 60:02d}",
                'description': result_description
            })
        else:
            return jsonify({'error': 'خطا در ذخیره نتیجه آزمون'}), 500
        
    except Exception as e:
        print(f"❌ Error in submit_product_exam: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

# این کدها را به فایل app.py اضافه کنید

@app.route('/exam_performance_report')
def exam_performance_report():
    """گزارش عملکرد بازاریابان در آزمون‌ها - فقط برای ادمین"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط ادمین می‌تونه این گزارش رو ببینه
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('exam_performance_report.html', user=session['user_info'])

def load_exam_results_from_excel():
    """بارگذاری نتایج آزمون‌ها از فایل azmon.xlsx شیت azmon"""
    try:
        if not os.path.exists(EXAMS_FILE):
            return pd.DataFrame()
            
        # بررسی وجود شیت azmon
        with pd.ExcelFile(EXAMS_FILE) as xls:
            if 'azmon' not in xls.sheet_names:
                return pd.DataFrame()
        
        df = pd.read_excel(EXAMS_FILE, sheet_name='azmon')
        print(f"✅ Exam results loaded: {len(df)} records")
        
        # پاک کردن فاصله‌های اضافی
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
        
        return df
        
    except Exception as e:
        print(f"❌ Error loading exam results: {e}")
        return pd.DataFrame()

@app.route('/get_exam_performance_report', methods=['POST'])
def get_exam_performance_report():
    """دریافت داده‌های گزارش عملکرد آزمون بازاریابان"""
    try:
        # چک احراز هویت
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        # فقط ادمین
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # دریافت داده‌های POST
        data = request.get_json()
        date_from = data.get('date_from', '').strip()
        date_to = data.get('date_to', '').strip()
        exam_type_filter = data.get('exam_type', 'all')
        brand_filter = data.get('brand', 'all')
        
        print(f"🎯 Exam performance report: {date_from} to {date_to}")
        print(f"📝 Filters: type={exam_type_filter}, brand={brand_filter}")
        
        # تبدیل تاریخ به میلادی اگر لازم باشه
        date_from_gregorian = None
        date_to_gregorian = None
        
        if date_from and date_to:
            date_from_gregorian = jalali_to_gregorian(date_from)
            date_to_gregorian = jalali_to_gregorian(date_to)
            
            if not date_from_gregorian or not date_to_gregorian:
                return jsonify({'error': 'فرمت تاریخ نامعتبر است'}), 400
        
        # بارگذاری داده‌ها
        exam_results_df = load_exam_results_from_excel()
        users_df = load_users_from_excel()
        exams_df = load_exams_from_excel()
        
        if exam_results_df.empty:
            return jsonify({
                'success': True,
                'salespeople': [],
                'summary': {
                    'total_participants': 0,
                    'total_exams': 0,
                    'average_score': 0,
                    'pass_rate': 0
                },
                'message': 'هیچ نتیجه آزمونی یافت نشد'
            })
        
        print(f"📊 Found {len(exam_results_df)} exam results")
        
        # فیلتر بر اساس تاریخ
        if date_from_gregorian and date_to_gregorian:
            def convert_exam_date_to_gregorian(date_value):
                if pd.isna(date_value):
                    return None
                date_str = str(date_value).strip()
                if '/' in date_str and len(date_str.split('/')) == 3:
                    return jalali_to_gregorian(date_str)
                return None
            
            exam_results_df['ExamDateConverted'] = exam_results_df['ExamDate'].apply(convert_exam_date_to_gregorian)
            
            filtered_results = exam_results_df[
                (exam_results_df['ExamDateConverted'] >= date_from_gregorian) &
                (exam_results_df['ExamDateConverted'] <= date_to_gregorian)
            ]
        else:
            filtered_results = exam_results_df
        
        # فیلتر بر اساس نوع آزمون
        if exam_type_filter != 'all':
            filtered_results = filtered_results[filtered_results['ExamType'] == exam_type_filter]
        
        # فیلتر بر اساس برند
        if brand_filter != 'all':
            filtered_results = filtered_results[filtered_results['BrandName'] == brand_filter]
        
        print(f"🔍 After filtering: {len(filtered_results)} results")
        
        if filtered_results.empty:
            return jsonify({
                'success': True,
                'salespeople': [],
                'summary': {
                    'total_participants': 0,
                    'total_exams': 0,
                    'average_score': 0,
                    'pass_rate': 0
                },
                'message': 'هیچ نتیجه‌ای در این بازه زمانی یافت نشد'
            })
        
        # تجمیع نتایج بر اساس بازاریاب
        salesperson_performance = {}
        
        for _, result in filtered_results.iterrows():
            bazaryab_code = result['BazaryabCode']
            bazaryab_name = result.get('BazaryabName', 'نامشخص')
            
            if bazaryab_code not in salesperson_performance:
                salesperson_performance[bazaryab_code] = {
                    'salesperson_code': bazaryab_code,
                    'salesperson_name': bazaryab_name,
                    'total_exams': 0,
                    'total_score': 0,
                    'scores': [],
                    'exam_details': [],
                    'passed_exams': 0,
                    'excellent_scores': 0,  # نمرات بالای 80
                    'good_scores': 0,       # نمرات 60-80
                    'poor_scores': 0        # نمرات زیر 60
                }
            
            # اضافه کردن نتیجه
            score = float(result.get('Score', 0)) if not pd.isna(result.get('Score', 0)) else 0
            percentage = float(result.get('Percentage', 0)) if not pd.isna(result.get('Percentage', 0)) else 0
            
            salesperson_performance[bazaryab_code]['total_exams'] += 1
            salesperson_performance[bazaryab_code]['total_score'] += score
            salesperson_performance[bazaryab_code]['scores'].append(score)
            
            # دسته‌بندی نمرات
            if percentage >= 80:
                salesperson_performance[bazaryab_code]['excellent_scores'] += 1
            elif percentage >= 60:
                salesperson_performance[bazaryab_code]['good_scores'] += 1
            else:
                salesperson_performance[bazaryab_code]['poor_scores'] += 1
            
            # آزمون‌های قبولی (نمره بالای 60)
            if percentage >= 60:
                salesperson_performance[bazaryab_code]['passed_exams'] += 1
            
            # جزئیات آزمون
            salesperson_performance[bazaryab_code]['exam_details'].append({
                'exam_code': result.get('ExamCode', ''),
                'exam_date': result.get('ExamDate', ''),
                'exam_type': result.get('ExamType', ''),
                'brand_name': result.get('BrandName', ''),
                'score': int(score),
                'percentage': round(percentage, 1),
                'total_questions': int(result.get('TotalQuestions', 0)),
                'correct_answers': int(result.get('CorrectAnswers', 0)),
                'time_taken': result.get('TimeTaken', '')
            })
        
        # محاسبه آمار نهایی و مرتب‌سازی
        salespeople_list = []
        total_all_scores = 0
        total_all_exams = 0
        total_passed = 0
        
        for sp_data in salesperson_performance.values():
            # محاسبه میانگین
            avg_score = sp_data['total_score'] / sp_data['total_exams'] if sp_data['total_exams'] > 0 else 0
            pass_rate = (sp_data['passed_exams'] / sp_data['total_exams'] * 100) if sp_data['total_exams'] > 0 else 0
            
            # مرتب‌سازی جزئیات آزمون‌ها بر اساس تاریخ (جدیدترین اول)
            sp_data['exam_details'].sort(key=lambda x: x['exam_date'], reverse=True)
            
            salespeople_list.append({
                'salesperson_code': sp_data['salesperson_code'],
                'salesperson_name': sp_data['salesperson_name'],
                'total_exams': sp_data['total_exams'],
                'average_score': round(avg_score, 1),
                'passed_exams': sp_data['passed_exams'],
                'pass_rate': round(pass_rate, 1),
                'excellent_scores': sp_data['excellent_scores'],
                'good_scores': sp_data['good_scores'],
                'poor_scores': sp_data['poor_scores'],
                'exam_details': sp_data['exam_details']
            })
            
            # آمار کلی
            total_all_scores += sp_data['total_score']
            total_all_exams += sp_data['total_exams']
            total_passed += sp_data['passed_exams']
        
        # مرتب‌سازی بر اساس میانگین نمره (بالا به پایین)
        salespeople_list.sort(key=lambda x: x['average_score'], reverse=True)
        
        # آمار کلی
        overall_average = total_all_scores / total_all_exams if total_all_exams > 0 else 0
        overall_pass_rate = (total_passed / total_all_exams * 100) if total_all_exams > 0 else 0
        
        summary_stats = {
            'total_participants': len(salespeople_list),
            'total_exams': total_all_exams,
            'average_score': round(overall_average, 1),
            'pass_rate': round(overall_pass_rate, 1)
        }
        
        print(f"✅ Exam performance analysis complete:")
        print(f"   Participants: {len(salespeople_list)}")
        print(f"   Total exams: {total_all_exams}")
        print(f"   Average score: {overall_average:.1f}")
        print(f"   Pass rate: {overall_pass_rate:.1f}%")
        
        return jsonify({
            'success': True,
            'salespeople': salespeople_list,
            'summary': summary_stats,
            'date_from': date_from,
            'date_to': date_to,
            'filters': {
                'exam_type': exam_type_filter,
                'brand': brand_filter
            },
            'period_info': f"{date_from} تا {date_to}" if date_from and date_to else "تمام دوره‌ها"
        })
        
    except Exception as e:
        print(f"❌ Error in get_exam_performance_report: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/get_exam_filters')
def get_exam_filters():
    """دریافت فیلترهای آزمون (انواع آزمون و برندها)"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        if session['user_info']['Typev'] != 'admin':
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # بارگذاری آزمون‌ها
        exams_df = load_exams_from_excel()
        
        if exams_df.empty:
            return jsonify({
                'exam_types': [],
                'brands': []
            })
        
        # دریافت انواع آزمون
        exam_types = sorted(exams_df['ExamType'].dropna().unique().tolist())
        
        # دریافت برندها
        brands = sorted(exams_df['BrandName'].dropna().unique().tolist())
        
        return jsonify({
            'exam_types': exam_types,
            'brands': brands
        })
        
    except Exception as e:
        print(f"❌ Error in get_exam_filters: {str(e)}")
        return jsonify({'error': str(e)}), 500


# اضافه کردن این تابع به app.py
def get_location_by_ip(ip_address=None):
    """دریافت مکان بر اساس IP"""
    try:
        # سرویس‌های مختلف برای IP Location
        services = [
            f"https://ipapi.co/{ip_address}/json/" if ip_address else "https://ipapi.co/json/",
            "http://ip-api.com/json/",
            "https://ipinfo.io/json"
        ]
        
        for service_url in services:
            try:
                response = requests.get(service_url, timeout=5)
                if response.status_code == 200:
                    data = response.json()
                    
                    # استخراج coordinates از response های مختلف
                    lat, lon = None, None
                    
                    if 'latitude' in data and 'longitude' in data:
                        lat, lon = data['latitude'], data['longitude']
                    elif 'lat' in data and 'lon' in data:
                        lat, lon = data['lat'], data['lon']
                    elif 'loc' in data:  # ipinfo.io format
                        lat, lon = data['loc'].split(',')
                        lat, lon = float(lat), float(lon)
                    
                    if lat and lon:
                        return {
                            'latitude': float(lat),
                            'longitude': float(lon),
                            'city': data.get('city', 'نامشخص'),
                            'country': data.get('country_name', data.get('country', 'نامشخص')),
                            'accuracy': 'city_level',
                            'source': service_url
                        }
            except:
                continue
        
        return None
    except Exception as e:
        print(f"Error in IP location: {e}")
        return None

# اضافه کردن این route به app.py
@app.route('/api/location/ip')
def api_location_ip():
    """API endpoint برای دریافت مکان بر اساس IP"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    # دریافت IP کاربر
    user_ip = request.environ.get('HTTP_X_FORWARDED_FOR', request.remote_addr)
    if user_ip == '127.0.0.1':
        user_ip = None  # برای localhost از خودکار استفاده کن
    
    location_data = get_location_by_ip(user_ip)
    
    if location_data:
        return jsonify({
            'success': True,
            'location': location_data
        })
    else:
        return jsonify({
            'success': False,
            'error': 'نتوانستیم مکان شما را تشخیص دهیم'
        }), 404

# محل‌های پیش‌فرض برای شهرهای بزرگ ایران
DEFAULT_LOCATIONS = {
    'تهران': {'lat': 35.6892, 'lon': 51.3890},
    'اصفهان': {'lat': 32.6546, 'lon': 51.6680},
    'شیراز': {'lat': 29.5918, 'lon': 52.5837},
    'مشهد': {'lat': 36.2605, 'lon': 59.6168},
    'تبریز': {'lat': 38.0962, 'lon': 46.2738},
    'کرج': {'lat': 35.8327, 'lon': 50.9916},
    'اهواز': {'lat': 31.3183, 'lon': 48.6706},
    'رشت': {'lat': 37.4482, 'lon': 49.1267},
    'قم': {'lat': 34.6401, 'lon': 50.8764},
    'ساری': {'lat': 36.5659, 'lon': 53.0586}
}

# اضافه کردن این route به app.py
@app.route('/api/location/city/<city_name>')
def api_location_city(city_name):
    """API برای دریافت مختصات شهر"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    city_location = DEFAULT_LOCATIONS.get(city_name)
    
    if city_location:
        return jsonify({
            'success': True,
            'location': {
                'latitude': city_location['lat'],
                'longitude': city_location['lon'],
                'city': city_name,
                'accuracy': 'city_center'
            }
        })
    else:
        return jsonify({
            'success': False,
            'error': f'مختصات شهر {city_name} در دسترس نیست'
        }), 404

# اضافه کردن این route به app.py برای مدیریت ترتیب برندها

@app.route('/brand_management')
def brand_management():
    """صفحه مدیریت ترتیب برندها - فقط برای ادمین"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # فقط ادمین می‌تونه این صفحه رو ببینه
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('brand_management.html', user=session['user_info'])

@app.route('/get_current_brand_order')
def get_current_brand_order():
    """دریافت ترتیب فعلی برندها"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    
    try:
        # بارگذاری محصولات برای دریافت لیست کامل برندها
        products_df = load_products_from_excel()
        if products_df is None:
            return jsonify({'error': 'فایل محصولات یافت نشد'}), 500
        
        # تمام برندهای موجود
        all_brands = sorted(products_df['Brand'].dropna().unique().tolist())
        
        # ترتیب فعلی از شیت brand
        current_order = load_brand_order_from_excel()
        
        if current_order:
            # برندهای جدیدی که در شیت brand نیست را اضافه کن
            for brand in all_brands:
                if brand not in current_order:
                    current_order.append(brand)
            
            return jsonify({
                'success': True,
                'current_order': current_order,
                'all_brands': all_brands,
                'has_custom_order': True
            })
        else:
            # اگر شیت brand وجود ندارد، ترتیب الفبایی
            return jsonify({
                'success': True,
                'current_order': all_brands,
                'all_brands': all_brands,
                'has_custom_order': False
            })
            
    except Exception as e:
        print(f"❌ Error in get_current_brand_order: {str(e)}")
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/update_brand_order', methods=['POST'])
def update_brand_order():
    """به‌روزرسانی ترتیب برندها"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    
    try:
        data = request.get_json()
        new_order = data.get('brand_order', [])
        
        if not new_order or not isinstance(new_order, list):
            return jsonify({'error': 'ترتیب برندها نامعتبر است'}), 400
        
        print(f"🔄 Updating brand order to: {new_order}")
        
        # ذخیره ترتیب جدید
        if save_brand_order_to_excel(new_order):
            return jsonify({
                'success': True,
                'message': 'ترتیب برندها با موفقیت به‌روزرسانی شد'
            })
        else:
            return jsonify({'error': 'خطا در ذخیره ترتیب جدید'}), 500
            
    except Exception as e:
        print(f"❌ Error in update_brand_order: {str(e)}")
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/reset_brand_order', methods=['POST'])
def reset_brand_order():
    """بازنشانی ترتیب برندها به حالت الفبایی"""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'Access denied'}), 403
    
    try:
        # بارگذاری محصولات
        products_df = load_products_from_excel()
        if products_df is None:
            return jsonify({'error': 'فایل محصولات یافت نشد'}), 500
        
        # ترتیب الفبایی برندها
        alphabetical_order = sorted(products_df['Brand'].dropna().unique().tolist())
        
        print(f"🔄 Resetting brand order to alphabetical: {alphabetical_order}")
        
        # ذخیره ترتیب الفبایی
        if save_brand_order_to_excel(alphabetical_order):
            return jsonify({
                'success': True,
                'new_order': alphabetical_order,
                'message': 'ترتیب برندها به حالت الفبایی بازنشانی شد'
            })
        else:
            return jsonify({'error': 'خطا در بازنشانی ترتیب'}), 500
            
    except Exception as e:
        print(f"❌ Error in reset_brand_order: {str(e)}")
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

    
# مقایسه

# اضافه کردن این کدها به فایل app.py

@app.route('/comparative_sales_report')
def comparative_sales_report():
    """صفحه گزارش مقایسه‌ای فروش - برای ادمین و کاربران"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    return render_template('comparative_sales_report.html', user=session['user_info'])

def clean_dataframe_for_json(df):
    """تمیز کردن DataFrame از مقادیر NaN برای JSON serialization"""
    if df is None or df.empty:
        return df
    
    df = df.copy()
    
    # تبدیل همه مقادیر NaN به مقادیر مناسب
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].fillna('').astype(str)
        elif df[col].dtype in ['int64', 'float64', 'int32', 'float32']:
            df[col] = df[col].fillna(0)
        elif df[col].dtype == 'bool':
            df[col] = df[col].fillna(False)
    
    return df

def safe_convert_to_dict(data):
    """تبدیل امن داده‌ها به dict با حذف مقادیر NaN"""
    if isinstance(data, pd.DataFrame):
        # تمیز کردن DataFrame
        clean_data = clean_dataframe_for_json(data)
        return clean_data.to_dict('records')
    elif isinstance(data, dict):
        # تمیز کردن dictionary
        clean_dict = {}
        for key, value in data.items():
            if pd.isna(value):
                if isinstance(value, (int, float)):
                    clean_dict[key] = 0
                else:
                    clean_dict[key] = ''
            else:
                clean_dict[key] = value
        return clean_dict
    else:
        return data


def get_sales_comparison_data(periods, user_code=None, user_type='admin'):
    """
    محاسبه داده‌های مقایسه‌ای فروش برای چندین دوره
    periods: لیستی از دوره‌ها شامل سال و ماه‌ها
    user_code: کد کاربر (برای فیلتر کردن مشتریان)
    user_type: نوع کاربر (admin یا user)
    """
    try:
        print(f"🔄 Starting comparative sales analysis for {len(periods)} periods")
        
        # بارگذاری داده‌ها
        sales_df = load_sales_from_excel()
        customers_df = load_customers_from_excel()
        products_df = load_products_from_excel()
        
        if sales_df is None or customers_df is None or products_df is None:
            print("❌ Failed to load required data files")
            return None
        
        # تمیز کردن داده‌ها از NaN
        print("🧹 Cleaning data from NaN values...")
        sales_df = clean_dataframe_for_json(sales_df)
        customers_df = clean_dataframe_for_json(customers_df)
        products_df = clean_dataframe_for_json(products_df)
        
        print(f"📊 Data loaded: {len(sales_df)} sales, {len(customers_df)} customers")
        
        # فیلتر مشتریان بر اساس نوع کاربر
        if user_type != 'admin' and user_code:
            my_customers = customers_df[customers_df['BazaryabCode'] == user_code]
            customer_codes = my_customers['CustomerCode'].tolist()
            filtered_customers = my_customers
            print(f"👤 User filter applied: {len(customer_codes)} customers for user {user_code}")
        else:
            customer_codes = customers_df['CustomerCode'].tolist()
            filtered_customers = customers_df
            print(f"👑 Admin access: {len(customer_codes)} total customers")
        
        # فیلتر فروش‌های مربوط به مشتریان
        relevant_sales = sales_df[sales_df['CustomerCode'].isin(customer_codes)]
        print(f"💰 Relevant sales found: {len(relevant_sales)} records")
        
        # آماده کردن داده‌های مقایسه‌ای
        comparison_data = {}
        
        # پردازش هر دوره
        for period_index, period in enumerate(periods):
            year = int(period['year'])
            months = [int(m) for m in period['months']]
            period_key = f"{year}_{'-'.join(map(str, months))}"
            
            print(f"🔍 Processing period {period_index + 1}: Year {year}, Months {months}")
            
            period_sales = []
            
            # فیلتر فروش‌ها برای هر ماه در سال انتخابی
            for month in months:
                month_sales = filter_sales_by_jalali_date_range(
                    relevant_sales, year, month, year, month
                )
                if not month_sales.empty:
                    period_sales.append(month_sales)
            
            # ترکیب فروش‌های دوره
            if period_sales:
                combined_sales = pd.concat(period_sales, ignore_index=True)
                # تمیز کردن داده‌های ترکیبی
                combined_sales = clean_dataframe_for_json(combined_sales)
            else:
                combined_sales = pd.DataFrame()
            
            print(f"   📈 Period sales: {len(combined_sales)} records")
            
            # محاسبه آمار هر مشتری در این دوره
            period_customer_stats = {}
            
            for _, customer in filtered_customers.iterrows():
                customer_code = str(customer['CustomerCode']).strip()
                customer_name = str(customer['CustomerName']).strip()
                
                customer_sales = combined_sales[
                    combined_sales['CustomerCode'] == customer_code
                ] if not combined_sales.empty else pd.DataFrame()
                
                # محاسبه آمار با در نظر گیری مقادیر خالی
                total_amount = 0
                total_quantity = 0
                unique_products = 0
                order_count = 0
                
                if not customer_sales.empty:
                    # محاسبه امن مبلغ کل
                    amounts = customer_sales['TotalAmount'].fillna(0)
                    total_amount = float(amounts.sum()) if not amounts.empty else 0
                    
                    # محاسبه امن تعداد
                    quantities = customer_sales['Quantity'].fillna(0)
                    total_quantity = int(quantities.sum()) if not quantities.empty else 0
                    
                    # تعداد محصولات منحصر به فرد
                    unique_products = len(customer_sales['ProductCode'].dropna().unique())
                    order_count = len(customer_sales)
                
                period_customer_stats[customer_code] = {
                    'customer_name': customer_name,
                    'total_amount': float(total_amount),
                    'total_quantity': int(total_quantity),
                    'unique_products': int(unique_products),
                    'order_count': int(order_count)
                }
            
            # محاسبه مجموع دوره
            period_total = sum([
                float(stats['total_amount']) for stats in period_customer_stats.values()
            ])
            
            comparison_data[period_key] = {
                'year': int(year),
                'months': [int(m) for m in months],
                'customers': period_customer_stats,
                'period_total': float(period_total),
                'period_description': f"سال {year} - ماه‌های {', '.join(map(str, months))}"
            }
            
            print(f"   ✅ Period {period_index + 1} processed: {len(period_customer_stats)} customers, total: {period_total:,.0f}")
        
        print(f"🎉 Comparative analysis completed successfully!")
        return comparison_data
        
    except Exception as e:
        print(f"❌ Error in get_sales_comparison_data: {e}")
        import traceback
        traceback.print_exc()
        return None


#
def filter_sales_by_jalali_date_range(sales_df, start_year, start_month, end_year, end_month):
    """فیلتر کردن فروش در بازه تاریخی شمسی"""
    try:
        if sales_df.empty:
            return pd.DataFrame()
        
        filtered_rows = []
        
        for index, row in sales_df.iterrows():
            try:
                invoice_date = row['InvoiceDate']
                
                if pd.isna(invoice_date):
                    continue
                
                # تبدیل تاریخ به شمسی
                if isinstance(invoice_date, str):
                    if '/' in invoice_date:
                        # فرمت شمسی: 1403/01/15
                        date_parts = invoice_date.split('/')
                        if len(date_parts) == 3:
                            invoice_year = int(date_parts[0])
                            invoice_month = int(date_parts[1])
                            
                            # بررسی قرار گیری در بازه
                            if (invoice_year == start_year and invoice_month >= start_month and
                                invoice_year == end_year and invoice_month <= end_month) or \
                               (invoice_year > start_year and invoice_year < end_year) or \
                               (invoice_year == start_year and invoice_month >= start_month and invoice_year < end_year) or \
                               (invoice_year > start_year and invoice_year == end_year and invoice_month <= end_month):
                                filtered_rows.append(row)
                    elif '-' in invoice_date:
                        # فرمت میلادی: 2024-03-21
                        gregorian_date = datetime.strptime(invoice_date, '%Y-%m-%d').date()
                        jalali_date = jdatetime.date.fromgregorian(date=gregorian_date)
                        
                        invoice_year = jalali_date.year
                        invoice_month = jalali_date.month
                        
                        if (invoice_year == start_year and invoice_month >= start_month and
                            invoice_year == end_year and invoice_month <= end_month) or \
                           (invoice_year > start_year and invoice_year < end_year) or \
                           (invoice_year == start_year and invoice_month >= start_month and invoice_year < end_year) or \
                           (invoice_year > start_year and invoice_year == end_year and invoice_month <= end_month):
                            filtered_rows.append(row)
                            
            except (ValueError, AttributeError):
                continue
        
        return pd.DataFrame(filtered_rows) if filtered_rows else pd.DataFrame()
        
    except Exception as e:
        print(f"Error in filter_sales_by_jalali_date_range: {e}")
        return pd.DataFrame()

@app.route('/get_comparative_sales_data', methods=['POST'])
def get_comparative_sales_data():
    """API برای درافت داده‌های مقایسه‌ای فروش - اصلاح شده"""
    try:
        if 'user_id' not in session:
            return jsonify({'success': False, 'error': 'لطفاً وارد شوید'}), 401
        
        data = request.get_json()
        periods = data.get('periods', [])
        
        if not periods or len(periods) < 1:
            return jsonify({'success': False, 'error': 'حداقل یک دوره باید انتخاب شود'}), 400
        
        print(f"📊 Comparative sales analysis request for {len(periods)} periods")
        
        # درافت داده‌های مقایسه‌ای
        user_code = session['user_info']['Codev']
        user_type = session['user_info']['Typev']
        
        comparison_data = get_sales_comparison_data(periods, user_code, user_type)
        
        if comparison_data is None:
            return jsonify({'success': False, 'error': 'خطا در پردازش داده‌ها'}), 500
        
        # محاسبه آمار مقایسه‌ای
        period_keys = list(comparison_data.keys())
        customer_comparison = {}
        
        # لیست کلیه مشتریان در تمام دوره‌ها
        all_customers = set()
        for period_data in comparison_data.values():
            all_customers.update(period_data['customers'].keys())
        
        print(f"👥 Total unique customers across all periods: {len(all_customers)}")
        
        # مقایسه هر مشتری در دوره‌های مختلف
        for customer_code in all_customers:
            customer_periods = {}
            customer_name = 'نامشخص'
            
            for period_key, period_data in comparison_data.items():
                if customer_code in period_data['customers']:
                    customer_info = period_data['customers'][customer_code]
                    customer_name = customer_info['customer_name']
                    customer_periods[period_key] = customer_info
                else:
                    customer_periods[period_key] = {
                        'customer_name': customer_name,
                        'total_amount': 0,
                        'total_quantity': 0,
                        'unique_products': 0,
                        'order_count': 0
                    }
            
            # محاسبه تغییرات
            period_values = list(customer_periods.values())
            changes = []
            
            if len(period_values) >= 2:
                for i in range(1, len(period_values)):
                    current = float(period_values[i]['total_amount'])
                    previous = float(period_values[i-1]['total_amount'])
                    
                    if previous > 0:
                        change_percent = ((current - previous) / previous) * 100
                        change_amount = current - previous
                    else:
                        change_percent = 100.0 if current > 0 else 0.0
                        change_amount = current
                    
                    changes.append({
                        'change_percent': round(float(change_percent), 1),
                        'change_amount': int(change_amount),
                        'trend': 'رشد' if change_amount > 0 else 'افت' if change_amount < 0 else 'ثابت'
                    })
            
            total_across_periods = sum([float(p['total_amount']) for p in period_values])
            average_per_period = total_across_periods / len(period_values) if period_values else 0
            
            customer_comparison[customer_code] = {
                'customer_name': customer_name,
                'periods': customer_periods,
                'changes': changes,
                'total_across_periods': float(total_across_periods),
                'average_per_period': float(average_per_period)
            }
        
        # آمار کلی
        summary_stats = {}
        for period_key, period_data in comparison_data.items():
            active_customers = len([
                c for c in period_data['customers'].values() 
                if float(c['total_amount']) > 0
            ])
            
            summary_stats[period_key] = {
                'period_description': period_data['period_description'],
                'total_sales': int(float(period_data['period_total'])),
                'active_customers': int(active_customers),
                'total_customers': len(period_data['customers'])
            }
        
        print(f"✅ Analysis complete: {len(customer_comparison)} customers analyzed")
        
        # اطمینان از عدم وجود مقادیر NaN در response نهایی
        response_data = {
            'success': True,
            'periods': periods,
            'customer_comparison': customer_comparison,
            'summary_stats': summary_stats,
            'period_descriptions': {
                k: str(v['period_description']) for k, v in comparison_data.items()
            }
        }
        
        return jsonify(response_data)
        
    except Exception as e:
        print(f"❌ Error in get_comparative_sales_data: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False, 
            'error': f'خطای سرور: {str(e)}'
        }), 500

@app.route('/get_customer_detailed_comparison', methods=['POST'])
def get_customer_detailed_comparison():
    """API برای دریافت جزئیات مقایسه‌ای یک مشتری خاص"""
    try:
        if 'user_id' not in session:
            return jsonify({'error': 'لطفاً وارد شوید'}), 401
        
        data = request.get_json()
        customer_code = data.get('customer_code')
        periods = data.get('periods', [])
        
        if not customer_code or not periods:
            return jsonify({'error': 'کد مشتری و دوره‌ها الزامی است'}), 400
        
        print(f"🔍 Detailed analysis for customer: {customer_code}")
        
        # بارگذاری داده‌ها
        customers_df = load_customers_from_excel()
        products_df = load_products_from_excel()
        sales_df = load_sales_from_excel()
        
        if customers_df is None or products_df is None or sales_df is None:
            return jsonify({'error': 'خطا در بارگذاری فایل‌ها'}), 500
        
        # بررسی دسترسی کاربر به این مشتری
        user_code = session['user_info']['Codev']
        user_type = session['user_info']['Typev']
        
        if user_type != 'admin':
            customer_info = customers_df[customers_df['CustomerCode'] == customer_code]
            if customer_info.empty or customer_info.iloc[0]['BazaryabCode'] != user_code:
                return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # اطلاعات کلی مشتری
        customer_info = customers_df[customers_df['CustomerCode'] == customer_code]
        if customer_info.empty:
            return jsonify({'error': 'مشتری یافت نشد'}), 404
        
        customer_detail = customer_info.iloc[0].to_dict()
        
        # تحلیل هر دوره
        period_analysis = {}
        all_products_purchased = set()
        
        for period in periods:
            year = period['year']
            months = period['months']
            period_key = f"{year}_{'-'.join(map(str, months))}"
            
            # فیلتر فروش‌ها برای این دوره
            period_sales = []
            for month in months:
                month_sales = filter_sales_by_jalali_date_range(
                    sales_df[sales_df['CustomerCode'] == customer_code],
                    year, month, year, month
                )
                if not month_sales.empty:
                    period_sales.append(month_sales)
            
            if period_sales:
                combined_sales = pd.concat(period_sales, ignore_index=True)
            else:
                combined_sales = pd.DataFrame()
            
            # محاسبه فروش هر محصول
            product_sales = {}
            period_total = 0
            
            if not combined_sales.empty:
                for _, sale in combined_sales.iterrows():
                    product_code = sale['ProductCode']
                    amount = float(sale.get('TotalAmount', 0))
                    quantity = int(sale.get('Quantity', 0))
                    
                    if product_code not in product_sales:
                        product_sales[product_code] = {
                            'total_amount': 0,
                            'total_quantity': 0,
                            'purchase_dates': []
                        }
                    
                    product_sales[product_code]['total_amount'] += amount
                    product_sales[product_code]['total_quantity'] += quantity
                    product_sales[product_code]['purchase_dates'].append({
                        'date': sale.get('InvoiceDate', ''),
                        'amount': amount,
                        'quantity': quantity
                    })
                    
                    period_total += amount
                    all_products_purchased.add(product_code)
            
            # اطلاعات محصولات خریداری شده
            purchased_products = []
            for product_code, sales_data in product_sales.items():
                product_info = products_df[products_df['ProductCode'] == product_code]
                
                if not product_info.empty:
                    product_detail = product_info.iloc[0]
                    purchased_products.append({
                        'product_code': product_code,
                        'product_name': product_detail.get('ProductName', ''),
                        'brand': product_detail.get('Brand', ''),
                        'category': product_detail.get('Category', ''),
                        'price': float(product_detail.get('Price', 0)),
                        'total_amount': int(sales_data['total_amount']),
                        'total_quantity': int(sales_data['total_quantity']),
                        'purchase_dates': sales_data['purchase_dates']
                    })
            
            # مرتب‌سازی بر اساس مبلغ خرید
            purchased_products.sort(key=lambda x: x['total_amount'], reverse=True)
            
            period_analysis[period_key] = {
                'year': year,
                'months': months,
                'period_description': f"سال {year} - ماه‌های {', '.join(map(str, months))}",
                'purchased_products': purchased_products,
                'period_total': int(period_total),
                'unique_products_count': len(purchased_products)
            }
        
        # محصولات خریداری نشده (محصولات موجود که این مشتری نخریده)
        all_purchased_codes = list(all_products_purchased)
        not_purchased_products = []
        
        for _, product in products_df.iterrows():
            if product['ProductCode'] not in all_purchased_codes:
                not_purchased_products.append({
                    'product_code': product['ProductCode'],
                    'product_name': product.get('ProductName', ''),
                    'brand': product.get('Brand', ''),
                    'category': product.get('Category', ''),
                    'price': float(product.get('Price', 0))
                })
        
        # مرتب‌سازی بر اساس قیمت
        not_purchased_products.sort(key=lambda x: x['price'], reverse=True)
        
        print(f"✅ Detailed analysis complete for customer {customer_code}")
        
        return jsonify({
            'success': True,
            'customer': customer_detail,
            'periods': periods,
            'period_analysis': period_analysis,
            'not_purchased_products': not_purchased_products[:50],  # محدود کردن به 50 محصول
            'summary': {
                'total_across_periods': sum([p['period_total'] for p in period_analysis.values()]),
                'unique_products_purchased': len(all_purchased_codes),
                'products_not_purchased': len(not_purchased_products)
            }
        })
        
    except Exception as e:
        print(f"❌ Error in get_customer_detailed_comparison: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500
# اضافه کردن این کدها به فایل app.py


# فایل‌های مربوط به تورهای ویزیت
VISIT_PERIODS_FILE = 'visit_periods.xlsx'
VISIT_TOURS_FILE = 'visit_tours.xlsx'
VISIT_EXECUTIONS_FILE = 'visit_executions.xlsx'

def create_visit_files_if_not_exist():
    """ایجاد فایل‌های مربوط به تورهای ویزیت در صورت عدم وجود"""
    try:
        # فایل دوره‌های ویزیت
        if not os.path.exists(VISIT_PERIODS_FILE):
            periods_df = pd.DataFrame(columns=[
                'PeriodCode', 'PeriodName', 'StartDate', 'EndDate', 
                'TotalTours', 'CreatedDate', 'CreatedBy', 'Status'
            ])
            periods_df.to_excel(VISIT_PERIODS_FILE, sheet_name='periods', index=False)
            print("✅ فایل دوره‌های ویزیت ایجاد شد")
        
        # فایل تورهای ویزیت
        if not os.path.exists(VISIT_TOURS_FILE):
            tours_df = pd.DataFrame(columns=[
                'TourCode', 'PeriodCode', 'TourNumber', 'TourDate', 
                'BazaryabCode', 'CustomerCodes', 'PrintedDate', 
                'ReceivedDate', 'Status', 'Notes'
            ])
            tours_df.to_excel(VISIT_TOURS_FILE, sheet_name='tours', index=False)
            print("✅ فایل تورهای ویزیت ایجاد شد")
        
        # فایل اجرای تورها
        if not os.path.exists(VISIT_EXECUTIONS_FILE):
            executions_df = pd.DataFrame(columns=[
                'ExecutionCode', 'TourCode', 'CustomerCode', 'VisitDate', 
                'VisitTime', 'BazaryabCode', 'Status', 'Notes'
            ])
            executions_df.to_excel(VISIT_EXECUTIONS_FILE, sheet_name='executions', index=False)
            print("✅ فایل اجرای تورها ایجاد شد")
        
        return True
    except Exception as e:
        print(f"❌ خطا در ایجاد فایل‌های ویزیت: {e}")
        return False

def load_visit_periods():
    """بارگذاری دوره‌های ویزیت"""
    try:
        create_visit_files_if_not_exist()
        df = pd.read_excel(VISIT_PERIODS_FILE, sheet_name='periods')
        return clean_dataframe_for_json(df)
    except Exception as e:
        print(f"❌ خطا در بارگذاری دوره‌های ویزیت: {e}")
        return pd.DataFrame()

def load_visit_tours():
    """بارگذاری تورهای ویزیت"""
    try:
        create_visit_files_if_not_exist()
        df = pd.read_excel(VISIT_TOURS_FILE, sheet_name='tours')
        return clean_dataframe_for_json(df)
    except Exception as e:
        print(f"❌ خطا در بارگذاری تورهای ویزیت: {e}")
        return pd.DataFrame()

def load_visit_executions():
    """بارگذاری اجرای تورها"""
    try:
        create_visit_files_if_not_exist()
        df = pd.read_excel(VISIT_EXECUTIONS_FILE, sheet_name='executions')
        return clean_dataframe_for_json(df)
    except Exception as e:
        print(f"❌ خطا در بارگذاری اجرای تورها: {e}")
        return pd.DataFrame()

def save_visit_periods(df):
    """ذخیره دوره‌های ویزیت"""
    try:
        df.to_excel(VISIT_PERIODS_FILE, sheet_name='periods', index=False)
        return True
    except Exception as e:
        print(f"❌ خطا در ذخیره دوره‌های ویزیت: {e}")
        return False

def save_visit_tours(df):
    """ذخیره تورهای ویزیت"""
    try:
        df.to_excel(VISIT_TOURS_FILE, sheet_name='tours', index=False)
        return True
    except Exception as e:
        print(f"❌ خطا در ذخیره تورهای ویزیت: {e}")
        return False

def save_visit_executions(df):
    """ذخیره اجرای تورها"""
    try:
        df.to_excel(VISIT_EXECUTIONS_FILE, sheet_name='executions', index=False)
        return True
    except Exception as e:
        print(f"❌ خطا در ذخیره اجرای تورها: {e}")
        return False

def generate_period_code():
    """تولید کد دوره ویزیت منحصر به فرد"""
    try:
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        date_str = jalali_now.strftime('%Y%m%d')
        
        periods_df = load_visit_periods()
        if not periods_df.empty:
            today_periods = periods_df[periods_df['PeriodCode'].str.contains(f'VP-{date_str}', na=False)]
            last_number = len(today_periods) + 1
        else:
            last_number = 1
        
        return f"VP-{date_str}{last_number:03d}"
    except Exception as e:
        print(f"❌ خطا در تولید کد دوره: {e}")
        return f"VP-{datetime.now().strftime('%Y%m%d%H%M%S')}"

def generate_tour_code(period_code, tour_number):
    """تولید کد تور ویزیت"""
    try:
        return f"{period_code}-T{tour_number:02d}"
    except Exception as e:
        print(f"❌ خطا در تولید کد تور: {e}")
        return f"TOUR-{datetime.now().strftime('%Y%m%d%H%M%S')}"

# Routes مربوط به مدیریت تورهای ویزیت

@app.route('/visit_management')
def visit_management():
    """صفحه مدیریت تورهای ویزیت - فقط برای ادمین"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    create_visit_files_if_not_exist()
    return render_template('visit_management.html', user=session['user_info'])

@app.route('/create_visit_period', methods=['POST'])
def create_visit_period():
    """ایجاد دوره ویزیت جدید"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'دسترسی غیرمجاز'}), 403
    
    try:
        data = request.get_json()
        period_name = data.get('period_name', '').strip()
        start_date = data.get('start_date', '').strip()
        end_date = data.get('end_date', '').strip()
        total_tours = int(data.get('total_tours', 4))
        
        if not period_name or not start_date or not end_date:
            return jsonify({'error': 'نام دوره و تاریخ‌ها الزامی است'}), 400
        
        # تولید کد دوره
        period_code = generate_period_code()
        
        # تاریخ و ساعت فعلی
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        created_date = jalali_now.strftime('%Y/%m/%d')
        
        # بارگذاری دوره‌های موجود
        periods_df = load_visit_periods()
        
        # ایجاد رکورد جدید
        new_period = pd.DataFrame([{
            'PeriodCode': period_code,
            'PeriodName': period_name,
            'StartDate': start_date,
            'EndDate': end_date,
            'TotalTours': total_tours,
            'CreatedDate': created_date,
            'CreatedBy': session['user_info']['Codev'],
            'Status': 'فعال'
        }])
        
        # اضافه کردن به DataFrame موجود
        if periods_df.empty:
            periods_df = new_period
        else:
            periods_df = pd.concat([periods_df, new_period], ignore_index=True)
        
        # ذخیره فایل
        if save_visit_periods(periods_df):
            print(f"✅ دوره ویزیت ایجاد شد: {period_code}")
            return jsonify({
                'success': True,
                'period_code': period_code,
                'message': 'دوره ویزیت با موفقیت ایجاد شد'
            })
        else:
            return jsonify({'error': 'خطا در ذخیره دوره ویزیت'}), 500
            
    except Exception as e:
        print(f"❌ خطا در ایجاد دوره ویزیت: {e}")
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/get_visit_periods')
def get_visit_periods():
    """دریافت لیست دوره‌های ویزیت"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'دسترسی غیرمجاز'}), 403
    
    try:
        periods_df = load_visit_periods()
        
        if periods_df.empty:
            return jsonify({'periods': []})
        
        # مرتب‌سازی بر اساس تاریخ (جدیدترین اول)
        periods_df = periods_df.sort_values('CreatedDate', ascending=False)
        
        periods = []
        for _, period in periods_df.iterrows():
            periods.append({
                'period_code': period.get('PeriodCode', ''),
                'period_name': period.get('PeriodName', ''),
                'start_date': period.get('StartDate', ''),
                'end_date': period.get('EndDate', ''),
                'total_tours': int(period.get('TotalTours', 0)),
                'created_date': period.get('CreatedDate', ''),
                'status': period.get('Status', 'فعال')
            })
        
        return jsonify({'periods': periods})
        
    except Exception as e:
        print(f"❌ خطا در دریافت دوره‌های ویزیت: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/create_tours_for_period', methods=['POST'])
def create_tours_for_period():
    """ایجاد تورهای ویزیت برای یک دوره"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'دسترسی غیرمجاز'}), 403
    
    try:
        data = request.get_json()
        period_code = data.get('period_code', '').strip()
        bazaryab_code = data.get('bazaryab_code', '').strip()
        
        if not period_code or not bazaryab_code:
            return jsonify({'error': 'کد دوره و بازاریاب الزامی است'}), 400
        
        # بارگذاری اطلاعات دوره
        periods_df = load_visit_periods()
        period_info = periods_df[periods_df['PeriodCode'] == period_code]
        
        if period_info.empty:
            return jsonify({'error': 'دوره یافت نشد'}), 404
        
        period_detail = period_info.iloc[0]
        total_tours = int(period_detail['TotalTours'])
        start_date = period_detail['StartDate']
        end_date = period_detail['EndDate']
        
        # بارگذاری مشتریان این بازاریاب
        customers_df = load_customers_from_excel()
        bazaryab_customers = customers_df[customers_df['BazaryabCode'] == bazaryab_code]
        
        if bazaryab_customers.empty:
            return jsonify({'error': 'هیچ مشتری برای این بازاریاب یافت نشد'}), 404
        
        # تقسیم مشتریان به تورها (حداکثر 20 مشتری در هر تور)
        customer_codes = bazaryab_customers['CustomerCode'].tolist()
        customers_per_tour = 20
        
        # بارگذاری تورهای موجود
        tours_df = load_visit_tours()
        
        # ایجاد تورها
        new_tours = []
        for tour_num in range(1, total_tours + 1):
            # تعیین مشتریان این تور
            start_idx = (tour_num - 1) * customers_per_tour
            end_idx = start_idx + customers_per_tour
            tour_customers = customer_codes[start_idx:end_idx]
            
            if not tour_customers:  # اگر مشتری نداریم، توقف
                break
            
            tour_code = generate_tour_code(period_code, tour_num)
            
            # محاسبه تاریخ تور (توزیع در طول دوره)
            start_dt = datetime.strptime(start_date, '%Y/%m/%d') if '/' in start_date else datetime.strptime(start_date, '%Y-%m-%d')
            end_dt = datetime.strptime(end_date, '%Y/%m/%d') if '/' in end_date else datetime.strptime(end_date, '%Y-%m-%d')
            
            days_diff = (end_dt - start_dt).days
            tour_day = start_dt + timedelta(days=(days_diff * (tour_num - 1) // max(1, total_tours - 1)))
            tour_date = tour_day.strftime('%Y/%m/%d')
            
            new_tour = {
                'TourCode': tour_code,
                'PeriodCode': period_code,
                'TourNumber': tour_num,
                'TourDate': tour_date,
                'BazaryabCode': bazaryab_code,
                'CustomerCodes': ','.join(tour_customers),
                'PrintedDate': '',
                'ReceivedDate': '',
                'Status': 'تعریف شده',
                'Notes': f'تور {tour_num} از {total_tours} - {len(tour_customers)} مشتری'
            }
            
            new_tours.append(new_tour)
        
        # اضافه کردن تورهای جدید
        if new_tours:
            new_tours_df = pd.DataFrame(new_tours)
            if tours_df.empty:
                tours_df = new_tours_df
            else:
                tours_df = pd.concat([tours_df, new_tours_df], ignore_index=True)
            
            # ذخیره فایل
            if save_visit_tours(tours_df):
                print(f"✅ {len(new_tours)} تور برای دوره {period_code} ایجاد شد")
                return jsonify({
                    'success': True,
                    'created_tours': len(new_tours),
                    'message': f'{len(new_tours)} تور ویزیت با موفقیت ایجاد شد'
                })
            else:
                return jsonify({'error': 'خطا در ذخیره تورها'}), 500
        else:
            return jsonify({'error': 'هیچ توری ایجاد نشد'}), 400
            
    except Exception as e:
        print(f"❌ خطا در ایجاد تورها: {e}")
        return jsonify({'error': f'خطای سرور: {str(e)}'}), 500

@app.route('/print_tour_list/<tour_code>')
def print_tour_list(tour_code):
    """چاپ لیست مشتریان یک تور"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if session['user_info']['Typev'] != 'admin':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    try:
        # بارگذاری اطلاعات تور
        tours_df = load_visit_tours()
        tour_info = tours_df[tours_df['TourCode'] == tour_code]
        
        if tour_info.empty:
            flash('تور یافت نشد!', 'error')
            return redirect(url_for('visit_management'))
        
        tour_detail = tour_info.iloc[0].to_dict()
        
        # بارگذاری اطلاعات مشتریان
        customers_df = load_customers_from_excel()
        customer_codes = tour_detail['CustomerCodes'].split(',')
        
        tour_customers = []
        for customer_code in customer_codes:
            customer = customers_df[customers_df['CustomerCode'] == customer_code.strip()]
            if not customer.empty:
                tour_customers.append(customer.iloc[0].to_dict())
        
        # بارگذاری اطلاعات بازاریاب
        users_df = load_users_from_excel()
        bazaryab_info = users_df[users_df['Codev'] == tour_detail['BazaryabCode']]
        bazaryab_name = bazaryab_info.iloc[0]['Namev'] if not bazaryab_info.empty else 'نامشخص'
        
        # به‌روزرسانی تاریخ چاپ
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        printed_date = jalali_now.strftime('%Y/%m/%d %H:%M')
        
        tours_df.loc[tours_df['TourCode'] == tour_code, 'PrintedDate'] = printed_date
        tours_df.loc[tours_df['TourCode'] == tour_code, 'Status'] = 'چاپ شده'
        save_visit_tours(tours_df)
        
        return render_template('tour_print.html', 
                             tour=tour_detail,
                             customers=tour_customers,
                             bazaryab_name=bazaryab_name,
                             printed_date=printed_date,
                             user=session['user_info'])
        
    except Exception as e:
        print(f"❌ خطا در چاپ تور: {e}")
        flash('خطا در بارگذاری اطلاعات تور!', 'error')
        return redirect(url_for('visit_management'))    

# اضافه کردن این Route ها به فایل app.py

@app.route('/get_visit_tours')
def get_visit_tours():
    """دریافت لیست تورهای ویزیت"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'دسترسی غیرمجاز'}), 403
    
    try:
        tours_df = load_visit_tours()
        
        if tours_df.empty:
            return jsonify({'tours': []})
        
        # بارگذاری اطلاعات تکمیلی
        users_df = load_users_from_excel()
        periods_df = load_visit_periods()
        
        tours = []
        for _, tour in tours_df.iterrows():
            # نام بازاریاب
            bazaryab_info = users_df[users_df['Codev'] == tour.get('BazaryabCode', '')]
            bazaryab_name = bazaryab_info.iloc[0]['Namev'] if not bazaryab_info.empty else 'نامشخص'
            
            # نام دوره
            period_info = periods_df[periods_df['PeriodCode'] == tour.get('PeriodCode', '')]
            period_name = period_info.iloc[0]['PeriodName'] if not period_info.empty else 'نامشخص'
            
            # تعداد مشتریان
            customer_codes = tour.get('CustomerCodes', '').split(',')
            customer_count = len([c for c in customer_codes if c.strip()])
            
            tours.append({
                'tour_code': tour.get('TourCode', ''),
                'period_code': tour.get('PeriodCode', ''),
                'period_name': period_name,
                'tour_number': int(tour.get('TourNumber', 0)),
                'tour_date': tour.get('TourDate', ''),
                'bazaryab_code': tour.get('BazaryabCode', ''),
                'bazaryab_name': bazaryab_name,
                'customer_count': customer_count,
                'printed_date': tour.get('PrintedDate', ''),
                'received_date': tour.get('ReceivedDate', ''),
                'status': tour.get('Status', 'تعریف شده'),
                'notes': tour.get('Notes', '')
            })
        
        # مرتب‌سازی بر اساس تاریخ تور
        tours.sort(key=lambda x: x['tour_date'], reverse=True)
        
        return jsonify({'tours': tours})
        
    except Exception as e:
        print(f"❌ خطا در دریافت تورها: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/mark_tour_received', methods=['POST'])
def mark_tour_received():
    """علامت‌گذاری تور به عنوان تحویل داده شده"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'دسترسی غیرمجاز'}), 403
    
    try:
        data = request.get_json()
        tour_code = data.get('tour_code', '').strip()
        
        if not tour_code:
            return jsonify({'error': 'کد تور الزامی است'}), 400
        
        # بارگذاری تورها
        tours_df = load_visit_tours()
        
        # پیدا کردن تور
        tour_index = tours_df[tours_df['TourCode'] == tour_code].index
        
        if tour_index.empty:
            return jsonify({'error': 'تور یافت نشد'}), 404
        
        # به‌روزرسانی وضعیت
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        received_date = jalali_now.strftime('%Y/%m/%d %H:%M')
        
        tours_df.loc[tour_index, 'ReceivedDate'] = received_date
        tours_df.loc[tour_index, 'Status'] = 'تحویل داده شده'
        
        # ذخیره
        if save_visit_tours(tours_df):
            return jsonify({
                'success': True,
                'message': 'تور با موفقیت به عنوان تحویل داده شده ثبت شد'
            })
        else:
            return jsonify({'error': 'خطا در ذخیره اطلاعات'}), 500
            
    except Exception as e:
        print(f"❌ خطا در ثبت تحویل تور: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/get_tour_customers/<tour_code>')
def get_tour_customers(tour_code):
    """دریافت لیست مشتریان یک تور برای ثبت ویزیت"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    try:
        # بارگذاری اطلاعات تور
        tours_df = load_visit_tours()
        tour_info = tours_df[tours_df['TourCode'] == tour_code]
        
        if tour_info.empty:
            return jsonify({'error': 'تور یافت نشد'}), 404
        
        tour_detail = tour_info.iloc[0]
        
        # بررسی دسترسی (ادمین یا خود بازاریاب)
        user_code = session['user_info']['Codev']
        user_type = session['user_info']['Typev']
        
        if user_type != 'admin' and tour_detail['BazaryabCode'] != user_code:
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # بارگذاری مشتریان
        customers_df = load_customers_from_excel()
        customer_codes = tour_detail['CustomerCodes'].split(',')
        
        tour_customers = []
        for customer_code in customer_codes:
            customer_code = customer_code.strip()
            if customer_code:
                customer = customers_df[customers_df['CustomerCode'] == customer_code]
                if not customer.empty:
                    tour_customers.append({
                        'customer_code': customer_code,
                        'customer_name': customer.iloc[0]['CustomerName']
                    })
        
        # بارگذاری ویزیت‌های قبلی این تور
        executions_df = load_visit_executions()
        existing_visits = executions_df[executions_df['TourCode'] == tour_code]
        visited_customers = existing_visits['CustomerCode'].tolist() if not existing_visits.empty else []
        
        return jsonify({
            'success': True,
            'tour_code': tour_code,
            'tour_info': tour_detail.to_dict(),
            'customers': tour_customers,
            'visited_customers': visited_customers
        })
        
    except Exception as e:
        print(f"❌ خطا در دریافت مشتریان تور: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/submit_tour_execution', methods=['POST'])
def submit_tour_execution():
    """ثبت ویزیت‌های انجام شده در یک تور"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    try:
        data = request.get_json()
        tour_code = data.get('tour_code', '').strip()
        visited_customers = data.get('visited_customers', [])
        
        if not tour_code or not visited_customers:
            return jsonify({'error': 'اطلاعات ناقص است'}), 400
        
        # بررسی دسترسی به تور
        tours_df = load_visit_tours()
        tour_info = tours_df[tours_df['TourCode'] == tour_code]
        
        if tour_info.empty:
            return jsonify({'error': 'تور یافت نشد'}), 404
        
        tour_detail = tour_info.iloc[0]
        user_code = session['user_info']['Codev']
        user_type = session['user_info']['Typev']
        
        if user_type != 'admin' and tour_detail['BazaryabCode'] != user_code:
            return jsonify({'error': 'دسترسی غیرمجاز'}), 403
        
        # بارگذاری ویزیت‌های موجود
        executions_df = load_visit_executions()
        
        # حذف ویزیت‌های قبلی این تور
        executions_df = executions_df[executions_df['TourCode'] != tour_code]
        
        # اضافه کردن ویزیت‌های جدید
        new_executions = []
        now = datetime.now()
        jalali_now = jdatetime.datetime.fromgregorian(datetime=now)
        visit_date = jalali_now.strftime('%Y/%m/%d')
        visit_time = now.strftime('%H:%M')
        
        for customer_code in visited_customers:
            execution_code = f"EX-{tour_code}-{customer_code}-{now.strftime('%H%M%S')}"
            
            new_execution = {
                'ExecutionCode': execution_code,
                'TourCode': tour_code,
                'CustomerCode': customer_code,
                'VisitDate': visit_date,
                'VisitTime': visit_time,
                'BazaryabCode': tour_detail['BazaryabCode'],
                'Status': 'انجام شده',
                'Notes': f'ویزیت در تور {tour_code}'
            }
            
            new_executions.append(new_execution)
        
        # اضافه کردن به DataFrame
        if new_executions:
            new_executions_df = pd.DataFrame(new_executions)
            if executions_df.empty:
                executions_df = new_executions_df
            else:
                executions_df = pd.concat([executions_df, new_executions_df], ignore_index=True)
        
        # ذخیره
        if save_visit_executions(executions_df):
            # به‌روزرسانی وضعیت تور
            tours_df.loc[tours_df['TourCode'] == tour_code, 'Status'] = 'در حال اجرا'
            save_visit_tours(tours_df)
            
            return jsonify({
                'success': True,
                'message': f'{len(visited_customers)} ویزیت با موفقیت ثبت شد'
            })
        else:
            return jsonify({'error': 'خطا در ذخیره ویزیت‌ها'}), 500
            
    except Exception as e:
        print(f"❌ خطا در ثبت ویزیت‌ها: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/get_visit_report', methods=['POST'])
def get_visit_report():
    """تولید گزارش ویزیت‌های انجام شده"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    if session['user_info']['Typev'] != 'admin':
        return jsonify({'error': 'دسترسی غیرمجاز'}), 403
    
    try:
        data = request.get_json()
        period_code = data.get('period_code', '').strip()
        bazaryab_code = data.get('bazaryab_code', '').strip()
        
        if not period_code:
            return jsonify({'error': 'کد دوره الزامی است'}), 400
        
        # بارگذاری اطلاعات
        periods_df = load_visit_periods()
        tours_df = load_visit_tours()
        executions_df = load_visit_executions()
        customers_df = load_customers_from_excel()
        users_df = load_users_from_excel()
        
        # اطلاعات دوره
        period_info = periods_df[periods_df['PeriodCode'] == period_code]
        if period_info.empty:
            return jsonify({'error': 'دوره یافت نشد'}), 404
        
        period_detail = period_info.iloc[0]
        
        # فیلتر تورهای این دوره
        period_tours = tours_df[tours_df['PeriodCode'] == period_code]
        
        # فیلتر بر اساس بازاریاب اگر انتخاب شده
        if bazaryab_code:
            period_tours = period_tours[period_tours['BazaryabCode'] == bazaryab_code]
        
        if period_tours.empty:
            return jsonify({
                'success': True,
                'period_info': period_detail.to_dict(),
                'report_data': [],
                'summary': {
                    'total_tours': 0,
                    'total_customers': 0,
                    'visited_customers': 0,
                    'visit_percentage': 0
                }
            })
        
        # تحلیل هر تور
        report_data = []
        total_customers = 0
        total_visited = 0
        
        for _, tour in period_tours.iterrows():
            tour_code = tour['TourCode']
            bazaryab_code_tour = tour['BazaryabCode']
            
            # نام بازاریاب
            bazaryab_info = users_df[users_df['Codev'] == bazaryab_code_tour]
            bazaryab_name = bazaryab_info.iloc[0]['Namev'] if not bazaryab_info.empty else 'نامشخص'
            
            # مشتریان تور
            customer_codes = tour['CustomerCodes'].split(',')
            tour_customers = []
            
            for customer_code in customer_codes:
                customer_code = customer_code.strip()
                if customer_code:
                    customer_info = customers_df[customers_df['CustomerCode'] == customer_code]
                    if not customer_info.empty:
                        customer_name = customer_info.iloc[0]['CustomerName']
                        
                        # بررسی ویزیت
                        visit_executed = not executions_df[
                            (executions_df['TourCode'] == tour_code) & 
                            (executions_df['CustomerCode'] == customer_code)
                        ].empty
                        
                        tour_customers.append({
                            'customer_code': customer_code,
                            'customer_name': customer_name,
                            'visited': visit_executed
                        })
                        
                        total_customers += 1
                        if visit_executed:
                            total_visited += 1
            
            # آمار تور
            tour_total = len(tour_customers)
            tour_visited = len([c for c in tour_customers if c['visited']])
            tour_percentage = (tour_visited / tour_total * 100) if tour_total > 0 else 0
            
            report_data.append({
                'tour_code': tour_code,
                'tour_number': int(tour['TourNumber']),
                'tour_date': tour['TourDate'],
                'bazaryab_name': bazaryab_name,
                'customers': tour_customers,
                'total_customers': tour_total,
                'visited_customers': tour_visited,
                'visit_percentage': round(tour_percentage, 1),
                'status': tour['Status']
            })
        
        # آمار کلی
        overall_percentage = (total_visited / total_customers * 100) if total_customers > 0 else 0
        
        summary = {
            'total_tours': len(report_data),
            'total_customers': total_customers,
            'visited_customers': total_visited,
            'visit_percentage': round(overall_percentage, 1)
        }
        
        return jsonify({
            'success': True,
            'period_info': period_detail.to_dict(),
            'report_data': report_data,
            'summary': summary
        })
        
    except Exception as e:
        print(f"❌ خطا در تولید گزارش: {e}")
        return jsonify({'error': str(e)}), 500

# Route برای صفحه ثبت ویزیت توسط بازاریاب
@app.route('/my_visit_tours')
def my_visit_tours():
    """صفحه تورهای ویزیت من - برای بازاریابان"""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if session['user_info']['Typev'] != 'user':
        flash('شما اجازه دسترسی به این صفحه را ندارید!', 'error')
        return redirect(url_for('index'))
    
    return render_template('my_visit_tours.html', user=session['user_info'])

@app.route('/get_my_tours')
def get_my_tours():
    """دریافت تورهای ویزیت بازاریاب"""
    if 'user_id' not in session:
        return jsonify({'error': 'لطفاً وارد شوید'}), 401
    
    if session['user_info']['Typev'] != 'user':
        return jsonify({'error': 'دسترسی غیرمجاز'}), 403
    
    try:
        bazaryab_code = session['user_info']['Codev']
        
        # بارگذاری تورهای این بازاریاب
        tours_df = load_visit_tours()
        my_tours = tours_df[tours_df['BazaryabCode'] == bazaryab_code]
        
        if my_tours.empty:
            return jsonify({'tours': []})
        
        # بارگذاری اطلاعات تکمیلی
        periods_df = load_visit_periods()
        executions_df = load_visit_executions()
        
        tours = []
        for _, tour in my_tours.iterrows():
            tour_code = tour['TourCode']
            
            # اطلاعات دوره
            period_info = periods_df[periods_df['PeriodCode'] == tour['PeriodCode']]
            period_name = period_info.iloc[0]['PeriodName'] if not period_info.empty else 'نامشخص'
            
            # تعداد مشتریان و ویزیت‌های انجام شده
            customer_codes = tour['CustomerCodes'].split(',')
            total_customers = len([c for c in customer_codes if c.strip()])
            
            executed_visits = executions_df[executions_df['TourCode'] == tour_code]
            visited_customers = len(executed_visits)
            
            tours.append({
                'tour_code': tour_code,
                'period_name': period_name,
                'tour_number': int(tour['TourNumber']),
                'tour_date': tour['TourDate'],
                'total_customers': total_customers,
                'visited_customers': visited_customers,
                'completion_percentage': round((visited_customers / total_customers * 100), 1) if total_customers > 0 else 0,
                'status': tour['Status'],
                'printed_date': tour.get('PrintedDate', ''),
                'received_date': tour.get('ReceivedDate', '')
            })
        
        # مرتب‌سازی بر اساس تاریخ
        tours.sort(key=lambda x: x['tour_date'], reverse=True)
        
        return jsonify({'tours': tours})
        
    except Exception as e:
        print(f"❌ خطا در دریافت تورهای بازاریاب: {e}")
        return jsonify({'error': str(e)}), 500

# اضافه کردن requests به requirements اگر نداری
# pip install requests
# ===============================
# پایان کدهای آزمون
# ===============================

if __name__ == '__main__':
    #print("🚀 Starting enhanced Flask application...")
    #print("📂 Files:")
    #print(f"   Users: {USERS_FILE}")
    #print(f"   Customers: {CUSTOMERS_FILE}")
    #print(f"   Visits: {VISITS_FILE}")
    #print("🌐 URL: http://127.0.0.1:5000")
    #print("👤 Test users:")
    #print("   Admin: ahmad / 123456")
    #print("   User:  maryam / 789012")
    #print("-" * 50)
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
   
    #app.run(debug=True)