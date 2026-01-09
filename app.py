"""
Flask Application for Employee Information Management and Salary Report Export
"""
from flask import Flask, render_template, request, jsonify, send_file, session
import pandas as pd
import os
import io
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

app = Flask(__name__)
app.secret_key = 'salary_report_secret_key_2024'

# Global storage for uploaded data
data_store = {
    'thong_tin': None,
    'luong': None,
    'columns_thong_tin': [],
    'columns_luong': [],
    'employees_list': []
}

# Vietnamese to ASCII mapping for PDF
VIETNAMESE_MAP = {
    'à': 'a', 'á': 'a', 'ả': 'a', 'ã': 'a', 'ạ': 'a',
    'ă': 'a', 'ằ': 'a', 'ắ': 'a', 'ẳ': 'a', 'ẵ': 'a', 'ặ': 'a',
    'â': 'a', 'ầ': 'a', 'ấ': 'a', 'ẩ': 'a', 'ẫ': 'a', 'ậ': 'a',
    'đ': 'd',
    'è': 'e', 'é': 'e', 'ẻ': 'e', 'ẽ': 'e', 'ẹ': 'e',
    'ê': 'e', 'ề': 'e', 'ế': 'e', 'ể': 'e', 'ễ': 'e', 'ệ': 'e',
    'ì': 'i', 'í': 'i', 'ỉ': 'i', 'ĩ': 'i', 'ị': 'i',
    'ò': 'o', 'ó': 'o', 'ỏ': 'o', 'õ': 'o', 'ọ': 'o',
    'ô': 'o', 'ồ': 'o', 'ố': 'o', 'ổ': 'o', 'ỗ': 'o', 'ộ': 'o',
    'ơ': 'o', 'ờ': 'o', 'ớ': 'o', 'ở': 'o', 'ỡ': 'o', 'ợ': 'o',
    'ù': 'u', 'ú': 'u', 'ủ': 'u', 'ũ': 'u', 'ụ': 'u',
    'ư': 'u', 'ừ': 'u', 'ứ': 'u', 'ử': 'u', 'ữ': 'u', 'ự': 'u',
    'ỳ': 'y', 'ý': 'y', 'ỷ': 'y', 'ỹ': 'y', 'ỵ': 'y',
    'À': 'A', 'Á': 'A', 'Ả': 'A', 'Ã': 'A', 'Ạ': 'A',
    'Ă': 'A', 'Ằ': 'A', 'Ắ': 'A', 'Ẳ': 'A', 'Ẵ': 'A', 'Ặ': 'A',
    'Â': 'A', 'Ầ': 'A', 'Ấ': 'A', 'Ẩ': 'A', 'Ẫ': 'A', 'Ậ': 'A',
    'Đ': 'D',
    'È': 'E', 'É': 'E', 'Ẻ': 'E', 'Ẽ': 'E', 'Ẹ': 'E',
    'Ê': 'E', 'Ề': 'E', 'Ế': 'E', 'Ể': 'E', 'Ễ': 'E', 'Ệ': 'E',
    'Ì': 'I', 'Í': 'I', 'Ỉ': 'I', 'Ĩ': 'I', 'Ị': 'I',
    'Ò': 'O', 'Ó': 'O', 'Ỏ': 'O', 'Õ': 'O', 'Ọ': 'O',
    'Ô': 'O', 'Ồ': 'O', 'Ố': 'O', 'Ổ': 'O', 'Ỗ': 'O', 'Ộ': 'O',
    'Ơ': 'O', 'Ờ': 'O', 'Ớ': 'O', 'Ở': 'O', 'Ỡ': 'O', 'Ợ': 'O',
    'Ù': 'U', 'Ú': 'U', 'Ủ': 'U', 'Ũ': 'U', 'Ụ': 'U',
    'Ư': 'U', 'Ừ': 'U', 'Ứ': 'U', 'Ử': 'U', 'Ữ': 'U', 'Ự': 'U',
    'Ỳ': 'Y', 'Ý': 'Y', 'Ỷ': 'Y', 'Ỹ': 'Y', 'Ỵ': 'Y',
}


def remove_accents(text):
    """Convert Vietnamese text to ASCII for PDF compatibility"""
    result = ''
    for char in str(text):
        result += VIETNAMESE_MAP.get(char, char)
    return result


def is_valid_column(col_name):
    """Check if column name is valid"""
    col_str = str(col_name).strip()
    if col_str.startswith('Unnamed') or col_str.startswith('Col_') or col_str.startswith('_'):
        return False
    if col_str == 'nan' or col_str == 'NaN' or pd.isna(col_name) or col_str == '':
        return False
    try:
        float(col_str)
        return False
    except:
        pass
    return True


def find_column_by_keywords(df, keywords):
    """Find column in dataframe that contains any of the keywords"""
    for col in df.columns:
        col_lower = str(col).lower()
        for keyword in keywords:
            if keyword.lower() in col_lower:
                return col
    return None


def get_value_from_df(df, row_data, keywords):
    """Get value from row data matching column keywords"""
    col = find_column_by_keywords(df, keywords)
    if col and col in df.columns:
        val = row_data[col] if col in row_data.index else None
        if pd.notna(val):
            try:
                return float(val)
            except:
                return val
    return 0


def format_number(val):
    """Format number with thousand separator"""
    if val is None or val == 0:
        return ""
    try:
        num = float(val)
        if num == int(num):
            return f"{int(num):,}".replace(",", ".")
        return f"{num:,.0f}".replace(",", ".")
    except:
        return str(val)


def clean_dataframe(df, sheet_name):
    """Clean and process the dataframe from Excel"""
    if 'thông tin' in sheet_name.lower() or 'thong tin' in sheet_name.lower():
        # Find header rows
        header_row = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = str(row.values)
            if 'Họ tên' in row_str or 'STT' in row_str:
                header_row = i
                break
        
        if header_row is not None:
            # Get main headers and sub-headers
            main_headers = df.iloc[header_row].tolist()
            sub_headers = df.iloc[header_row + 1].tolist() if header_row + 1 < len(df) else [None] * len(main_headers)
            
            # Combine headers - use sub-header if main is Unnamed
            combined_headers = []
            last_valid_main = ''
            for i, (main, sub) in enumerate(zip(main_headers, sub_headers)):
                main_str = str(main).strip() if pd.notna(main) and not str(main).startswith('Unnamed') else ''
                sub_str = str(sub).strip() if pd.notna(sub) and not str(sub).startswith('Unnamed') else ''
                
                # Skip numeric headers
                try:
                    if main_str and float(main_str):
                        main_str = ''
                except:
                    pass
                try:
                    if sub_str and float(sub_str):
                        sub_str = ''
                except:
                    pass
                
                if main_str:
                    last_valid_main = main_str
                
                if sub_str and main_str:
                    combined = f"{main_str} - {sub_str}"
                elif sub_str:
                    combined = f"{last_valid_main} - {sub_str}" if last_valid_main else sub_str
                elif main_str:
                    combined = main_str
                else:
                    combined = f'_Col_{i}'
                
                combined_headers.append(combined)
            
            # Skip the number row if exists
            data_start = header_row + 2
            if data_start < len(df):
                check_row = df.iloc[data_start]
                first_vals = [str(v) for v in check_row.head(3).values if pd.notna(v)]
                if first_vals and all(v.replace('.', '').isdigit() for v in first_vals):
                    data_start += 1
            
            df = df.iloc[data_start:]
            df.columns = combined_headers
        
        # Remove columns starting with _
        cols_to_keep = [col for col in df.columns if not str(col).startswith('_')]
        df = df[cols_to_keep]
        
        # Remove completely empty rows
        df = df.dropna(how='all')
        
        # Only keep rows where STT is a valid number >= 1
        stt_col = None
        for col in df.columns:
            if 'STT' in str(col).upper():
                stt_col = col
                break
        
        if stt_col:
            df[stt_col] = pd.to_numeric(df[stt_col], errors='coerce')
            df = df[df[stt_col].notna() & (df[stt_col] >= 1)]
        
        df = df.reset_index(drop=True)
    
    elif 'lương' in sheet_name.lower() or 'luong' in sheet_name.lower():
        # Find header rows
        header_row = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = str(row.values).upper()
            if 'HỌ TÊN' in row_str or 'HO TEN' in row_str or 'NHÂN VIÊN' in row_str:
                header_row = i
                break
        
        if header_row is not None:
            main_headers = df.iloc[header_row].tolist()
            sub_headers = df.iloc[header_row + 1].tolist() if header_row + 1 < len(df) else [None] * len(main_headers)
            
            combined_headers = []
            last_valid_main = ''
            for i, (main, sub) in enumerate(zip(main_headers, sub_headers)):
                main_str = str(main).strip() if pd.notna(main) and not str(main).startswith('Unnamed') else ''
                sub_str = str(sub).strip() if pd.notna(sub) and not str(sub).startswith('Unnamed') else ''
                
                try:
                    if main_str and float(main_str):
                        main_str = ''
                except:
                    pass
                try:
                    if sub_str and float(sub_str):
                        sub_str = ''
                except:
                    pass
                
                if main_str:
                    last_valid_main = main_str
                
                if sub_str and main_str:
                    combined = f"{main_str} - {sub_str}"
                elif sub_str:
                    combined = f"{last_valid_main} - {sub_str}" if last_valid_main else sub_str
                elif main_str:
                    combined = main_str
                else:
                    combined = f'_Col_{i}'
                
                combined_headers.append(combined)
            
            data_start = header_row + 2
            if data_start < len(df):
                check_row = df.iloc[data_start]
                first_vals = [str(v) for v in check_row.head(3).values if pd.notna(v)]
                if first_vals and all(v.replace('.', '').isdigit() for v in first_vals):
                    data_start += 1
            
            df = df.iloc[data_start:]
            df.columns = combined_headers
        
        cols_to_keep = [col for col in df.columns if not str(col).startswith('_')]
        df = df[cols_to_keep]
        
        df = df.dropna(how='all')
        
        # Only keep rows where STT is valid
        stt_col = None
        for col in df.columns:
            if 'STT' in str(col).upper():
                stt_col = col
                break
        
        if stt_col:
            df[stt_col] = pd.to_numeric(df[stt_col], errors='coerce')
            df = df[df[stt_col].notna() & (df[stt_col] >= 1)]
        
        df = df.reset_index(drop=True)
    
    return df


def find_employee_name_column(df):
    """Find the column that contains employee names"""
    for col in df.columns:
        col_lower = str(col).lower()
        if 'họ tên' in col_lower or 'ho ten' in col_lower or 'tên nhân viên' in col_lower:
            return col
    return None


def get_employees_list(df):
    """Get list of employees from dataframe"""
    name_col = find_employee_name_column(df)
    if name_col is None:
        return []
    
    employees = []
    for idx, row in df.iterrows():
        name = row[name_col]
        if pd.notna(name) and str(name).strip():
            employees.append({
                'index': int(idx),
                'name': str(name).strip()
            })
    return employees


def get_salary_slip_data(employee_info, salary_data, df_info, df_luong):
    """Extract salary slip data from employee info and salary data"""
    
    # Field mappings based on user requirements:
    # Họ tên = Họ tên
    # Lương thỏa thuận = Thuế TNCN - Tổng Thu Nhập
    # Lương đóng = Lương cơ bản
    # Số tài khoản = Số tài khoản ngân hàng
    # Tên ngân hàng = Tại ngân hàng - Chi nhánh
    # Lương thực tế = Thuế TNCN - Tổng Thu Nhập (same as luong_thoa_thuan)
    # BHXH = Các khoản Người Lao Động phải nộp cho CQNN - Tổng cộng
    # Đoàn phí = Kinh phí Công Đoàn - Phí đoàn viên
    # Thuế TNCN = Thuế TNCN - Thuế TNCN phải nộp
    # Tổng Số Tiền Lương Thực Nhận = Tổng Cộng Thu Nhập - Tổng Cộng Khoản Trừ
    
    data = {
        'ho_ten': '',
        'luong_thoa_thuan': 0,  # Thuế TNCN - Tổng Thu Nhập
        'luong_dong': 0,        # Lương cơ bản
        'so_tai_khoan': '',
        'ten_ngan_hang': '',
        'luong_thuc_te': 0,     # Thuế TNCN - Tổng Thu Nhập (same as luong_thoa_thuan)
        'bhxh': 0,              # Các khoản NLĐ phải nộp - Tổng cộng
        'doan_phi': 0,          # Kinh phí Công Đoàn - Phí đoàn viên
        'thue_tncn': 0,         # Thuế TNCN - Thuế TNCN phải nộp
    }
    
    # Get employee name from info
    name_col = find_employee_name_column(df_info)
    if name_col and name_col in employee_info.index:
        val = employee_info[name_col]
        if pd.notna(val):
            data['ho_ten'] = str(val)
    
    # Get bank info from employee info
    for col in employee_info.index:
        val = employee_info[col]
        col_lower = str(col).lower()
        if pd.notna(val):
            # Số tài khoản ngân hàng
            if 'số tài khoản' in col_lower and 'ngân hàng' in col_lower:
                data['so_tai_khoan'] = str(val)
            elif 'số tài khoản' in col_lower and not data['so_tai_khoan']:
                data['so_tai_khoan'] = str(val)
            
            # Tại ngân hàng - Chi nhánh
            if 'tại ngân hàng' in col_lower or ('ngân hàng' in col_lower and 'chi nhánh' in col_lower):
                data['ten_ngan_hang'] = str(val)
            elif 'ngân hàng' in col_lower and 'số' not in col_lower and not data['ten_ngan_hang']:
                data['ten_ngan_hang'] = str(val)
    
    # Get salary data
    if salary_data is not None:
        for col in salary_data.index:
            val = salary_data[col]
            col_str = str(col)
            col_lower = col_str.lower()
            if pd.notna(val):
                try:
                    num_val = float(val)
                except:
                    num_val = 0
                
                # Lương thỏa thuận & Lương thực tế = Thuế TNCN - Tổng Thu Nhập
                # Match exactly: "Thuế TNCN - Tổng Thu Nhập" (not containing chưa, chịu, tính, bao gồm)
                if ('thuế tncn' in col_lower and 'tổng thu nhập' in col_lower and 
                    'chưa' not in col_lower and 'chịu' not in col_lower and 
                    'tính' not in col_lower and 'bao gồm' not in col_lower):
                    data['luong_thoa_thuan'] = num_val
                    data['luong_thuc_te'] = num_val
                
                # Lương đóng = Lương cơ bản
                if 'lương cơ bản' in col_lower:
                    data['luong_dong'] = num_val
                
                # BHXH = Các khoản Người Lao Động phải nộp cho CQNN - Tổng cộng
                if 'người lao động phải nộp' in col_lower and 'tổng cộng' in col_lower:
                    data['bhxh'] = num_val
                elif 'nld phải nộp' in col_lower and 'tổng cộng' in col_lower:
                    data['bhxh'] = num_val
                
                # Đoàn phí = Kinh phí Công Đoàn - Phí đoàn viên
                if 'kinh phí công đoàn' in col_lower and 'phí đoàn viên' in col_lower:
                    data['doan_phi'] = num_val
                elif 'phí đoàn viên' in col_lower and data['doan_phi'] == 0:
                    data['doan_phi'] = num_val
                
                # Thuế TNCN = Thuế TNCN - Thuế TNCN phải nộp
                # Match: column contains "Thuế TNCN phải nộp" but NOT the multi-bracket columns
                if 'thuế tncn phải nộp' in col_lower and 'tr' not in col_lower and '%' not in col_lower:
                    data['thue_tncn'] = num_val
    
    # Calculate total deductions automatically
    data['tong_khoan_tru'] = data['bhxh'] + data['doan_phi'] + data['thue_tncn']
    
    # Tổng Cộng Thu Nhập = Lương thực tế (same as Thuế TNCN - Tổng Thu Nhập)
    data['tong_thu_nhap'] = data['luong_thuc_te']
    
    # Tổng Số Tiền Lương Thực Nhận = Tổng Cộng Thu Nhập - Tổng Cộng Khoản Trừ
    data['luong_thuc_nhan'] = data['tong_thu_nhap'] - data['tong_khoan_tru']
    
    return data


@app.route('/')
def index():
    """Main page"""
    columns_info = data_store['columns_thong_tin'] if data_store['thong_tin'] is not None else []
    columns_luong = data_store['columns_luong'] if data_store['luong'] is not None else []
    has_data = data_store['thong_tin'] is not None
    employees_list = data_store['employees_list']
    
    return render_template('index.html', 
                         columns_info=columns_info,
                         columns_luong=columns_luong,
                         has_data=has_data,
                         employees_list=employees_list)


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'Không tìm thấy file'})
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'success': False, 'error': 'Chưa chọn file'})
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': 'Chỉ hỗ trợ file Excel (.xlsx, .xls)'})
    
    try:
        # Read Excel file
        xlsx = pd.ExcelFile(file)
        
        # Process sheets
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
            df_cleaned = clean_dataframe(df, sheet_name)
            
            if 'thông tin' in sheet_name.lower() or 'thong tin' in sheet_name.lower():
                data_store['thong_tin'] = df_cleaned
                data_store['columns_thong_tin'] = [col for col in df_cleaned.columns if is_valid_column(col)]
                data_store['employees_list'] = get_employees_list(df_cleaned)
            elif 'lương' in sheet_name.lower() or 'luong' in sheet_name.lower():
                data_store['luong'] = df_cleaned
                data_store['columns_luong'] = [col for col in df_cleaned.columns if is_valid_column(col)]
        
        return jsonify({
            'success': True,
            'message': f'Đã upload thành công file: {file.filename}',
            'columns_info': data_store['columns_thong_tin'],
            'columns_luong': data_store['columns_luong'],
            'employees_list': data_store['employees_list'],
            'total_employees': len(data_store['employees_list'])
        })
    
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': f'Lỗi xử lý file: {str(e)}'})


@app.route('/get_employee/<int:employee_index>')
def get_employee(employee_index):
    """Get single employee data by index"""
    if data_store['thong_tin'] is None:
        return jsonify({'success': False, 'error': 'Chưa upload dữ liệu'})
    
    df_info = data_store['thong_tin']
    df_luong = data_store['luong']
    
    if employee_index >= len(df_info):
        return jsonify({'success': False, 'error': 'Không tìm thấy nhân viên'})
    
    employee = df_info.iloc[employee_index]
    name_col = find_employee_name_column(df_info)
    
    employee_data = {
        'index': employee_index,
        'info': {}
    }
    
    # Get info data
    for col in df_info.columns:
        if is_valid_column(col):
            val = employee[col]
            if pd.notna(val):
                employee_data['info'][col] = str(val)
    
    # Get salary data if available
    if df_luong is not None and name_col:
        employee_name = employee[name_col]
        luong_name_col = find_employee_name_column(df_luong)
        if luong_name_col and pd.notna(employee_name):
            salary_match = df_luong[df_luong[luong_name_col].astype(str).str.lower().str.strip() == str(employee_name).lower().strip()]
            if len(salary_match) > 0:
                employee_data['salary'] = {}
                for col in df_luong.columns:
                    if is_valid_column(col):
                        val = salary_match.iloc[0][col]
                        if pd.notna(val):
                            employee_data['salary'][col] = str(val)
    
    return jsonify({
        'success': True,
        'result': employee_data
    })


@app.route('/search', methods=['POST'])
def search_employee():
    """Search for employee"""
    if data_store['thong_tin'] is None:
        return jsonify({'success': False, 'error': 'Chưa upload dữ liệu. Vui lòng upload file Excel trước.'})
    
    data = request.get_json()
    search_term = data.get('search_term', '').strip()
    search_fields = data.get('search_fields', [])
    
    if not search_term:
        return jsonify({'success': False, 'error': 'Vui lòng nhập từ khóa tìm kiếm'})
    
    df_info = data_store['thong_tin']
    df_luong = data_store['luong']
    
    # Search in thong_tin
    mask = pd.Series([False] * len(df_info))
    
    if not search_fields:
        # Search in all columns
        for col in df_info.columns:
            if is_valid_column(col):
                mask = mask | df_info[col].astype(str).str.lower().str.contains(search_term.lower(), na=False)
    else:
        # Search in specific fields
        for field in search_fields:
            if field in df_info.columns:
                mask = mask | df_info[field].astype(str).str.lower().str.contains(search_term.lower(), na=False)
    
    results = df_info[mask]
    
    if len(results) == 0:
        return jsonify({'success': False, 'error': 'Không tìm thấy kết quả'})
    
    # Format results
    results_list = []
    name_col = find_employee_name_column(df_info)
    
    for idx, row in results.iterrows():
        employee_data = {
            'index': int(idx),
            'info': {}
        }
        
        # Get info data
        for col in df_info.columns:
            if is_valid_column(col):
                val = row[col]
                if pd.notna(val):
                    employee_data['info'][col] = str(val)
        
        # Get salary data if available
        if df_luong is not None and name_col:
            employee_name = row[name_col]
            luong_name_col = find_employee_name_column(df_luong)
            if luong_name_col and pd.notna(employee_name):
                salary_match = df_luong[df_luong[luong_name_col].astype(str).str.lower().str.strip() == str(employee_name).lower().strip()]
                if len(salary_match) > 0:
                    employee_data['salary'] = {}
                    for col in df_luong.columns:
                        if is_valid_column(col):
                            val = salary_match.iloc[0][col]
                            if pd.notna(val):
                                employee_data['salary'][col] = str(val)
        
        results_list.append(employee_data)
    
    return jsonify({
        'success': True,
        'results': results_list,
        'count': len(results_list)
    })


@app.route('/export/excel/<int:employee_index>', methods=['GET', 'POST'])
def export_excel(employee_index):
    """Export employee salary slip to Excel with specific template"""
    if data_store['thong_tin'] is None:
        return jsonify({'success': False, 'error': 'Không có dữ liệu'})
    
    try:
        df_info = data_store['thong_tin']
        df_luong = data_store['luong']
        
        if employee_index >= len(df_info):
            return jsonify({'success': False, 'error': 'Không tìm thấy nhân viên'})
        
        # Get month and year from request
        now = datetime.now()
        month = request.args.get('month', now.month, type=int)
        year = request.args.get('year', now.year, type=int)
        
        employee = df_info.iloc[employee_index]
        name_col = find_employee_name_column(df_info)
        employee_name = employee[name_col] if name_col else f'NhanVien_{employee_index}'
        
        # Get salary data
        salary_row = None
        if df_luong is not None and name_col:
            luong_name_col = find_employee_name_column(df_luong)
            if luong_name_col:
                salary_match = df_luong[df_luong[luong_name_col].astype(str).str.lower().str.strip() == str(employee_name).lower().strip()]
                if len(salary_match) > 0:
                    salary_row = salary_match.iloc[0]
        
        # Get salary slip data
        slip_data = get_salary_slip_data(employee, salary_row, df_info, df_luong)
        
        # Create workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Phiếu Lương"
        
        # Styles
        title_font = Font(bold=True, size=16, color="FF0000")
        header_font = Font(bold=True, size=11)
        normal_font = Font(size=10)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        light_yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
        orange_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        light_blue_fill = PatternFill(start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")
        light_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        
        # Set column widths
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 25
        
        # Title with selected month and year
        ws.merge_cells('A1:E1')
        ws['A1'] = f"PHIẾU LƯƠNG THÁNG {month} NĂM {year}"
        ws['A1'].font = title_font
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Info section header
        row = 3
        
        # Left column info
        info_rows = [
            ('Họ tên:', slip_data['ho_ten'], 'Ngày công chuẩn:', ''),
            ('Lương thỏa thuận:', format_number(slip_data['luong_thoa_thuan']), 'Ngày công thực tế:', ''),
            ('% Lương Thử việc:', '', 'Nghỉ phép:', ''),
            ('Lương đóng:', format_number(slip_data['luong_dong']), 'Tổng giờ tăng ca:', ''),
            ('Số tài khoản:', slip_data['so_tai_khoan'], 'Tên ngân hàng:', slip_data['ten_ngan_hang']),
        ]
        
        for info in info_rows:
            ws[f'A{row}'] = info[0]
            ws[f'A{row}'].font = header_font
            ws[f'A{row}'].fill = yellow_fill
            ws[f'A{row}'].border = border
            
            ws[f'B{row}'] = info[1]
            ws[f'B{row}'].border = border
            
            ws[f'C{row}'] = info[2]
            ws[f'C{row}'].font = header_font
            ws[f'C{row}'].fill = yellow_fill
            ws[f'C{row}'].border = border
            
            ws.merge_cells(f'D{row}:E{row}')
            ws[f'D{row}'] = info[3]
            ws[f'D{row}'].border = border
            ws[f'E{row}'].border = border
            
            row += 1
        
        # Table header
        row += 1
        headers = ['STT', 'Các Khoản Thu Nhập', '', 'Các Khoản Trừ Vào Lương', '']
        ws[f'A{row}'] = 'STT'
        ws[f'A{row}'].font = header_font
        ws[f'A{row}'].fill = orange_fill
        ws[f'A{row}'].border = border
        ws[f'A{row}'].alignment = Alignment(horizontal='center')
        
        ws.merge_cells(f'B{row}:C{row}')
        ws[f'B{row}'] = 'Các Khoản Thu Nhập'
        ws[f'B{row}'].font = header_font
        ws[f'B{row}'].fill = orange_fill
        ws[f'B{row}'].border = border
        ws[f'B{row}'].alignment = Alignment(horizontal='center')
        ws[f'C{row}'].border = border
        
        ws.merge_cells(f'D{row}:E{row}')
        ws[f'D{row}'] = 'Các Khoản Trừ Vào Lương'
        ws[f'D{row}'].font = header_font
        ws[f'D{row}'].fill = orange_fill
        ws[f'D{row}'].border = border
        ws[f'D{row}'].alignment = Alignment(horizontal='center')
        ws[f'E{row}'].border = border
        
        # Table data
        table_data = [
            ('1', 'Lương thực tế', format_number(slip_data['luong_thuc_te']), 'BHXH', format_number(slip_data['bhxh'])),
            ('2', 'Phép năm', '', 'Đoàn phí', format_number(slip_data['doan_phi'])),
            ('3', 'Lương tăng ca', '', 'Thuế Thu Nhập Cá Nhân', format_number(slip_data['thue_tncn'])),
            ('4', 'Lương bổ sung', '', 'Tạm Ứng', ''),
            ('5', 'Giữ xe', '', 'Tiền phạt', ''),
            ('6', 'Công tác phí', '', 'Khác', ''),
        ]
        
        row += 1
        for data_row in table_data:
            ws[f'A{row}'] = data_row[0]
            ws[f'A{row}'].border = border
            ws[f'A{row}'].alignment = Alignment(horizontal='center')
            ws[f'A{row}'].fill = light_blue_fill
            
            ws[f'B{row}'] = data_row[1]
            ws[f'B{row}'].border = border
            ws[f'B{row}'].fill = light_blue_fill
            
            ws[f'C{row}'] = data_row[2]
            ws[f'C{row}'].border = border
            ws[f'C{row}'].fill = light_green_fill
            
            ws[f'D{row}'] = data_row[3]
            ws[f'D{row}'].border = border
            ws[f'D{row}'].fill = light_blue_fill
            
            ws[f'E{row}'] = data_row[4]
            ws[f'E{row}'].border = border
            ws[f'E{row}'].fill = light_green_fill
            
            row += 1
        
        # Totals row
        ws[f'A{row}'] = ''
        ws[f'A{row}'].border = border
        
        ws[f'B{row}'] = 'Tổng Cộng Thu Nhập'
        ws[f'B{row}'].font = header_font
        ws[f'B{row}'].border = border
        ws[f'B{row}'].fill = yellow_fill
        
        ws[f'C{row}'] = format_number(slip_data['tong_thu_nhap'])
        ws[f'C{row}'].border = border
        ws[f'C{row}'].fill = light_yellow_fill
        
        ws[f'D{row}'] = 'Tổng Cộng Khoản Trừ'
        ws[f'D{row}'].font = header_font
        ws[f'D{row}'].border = border
        ws[f'D{row}'].fill = yellow_fill
        
        ws[f'E{row}'] = format_number(slip_data['tong_khoan_tru'])
        ws[f'E{row}'].border = border
        ws[f'E{row}'].fill = light_yellow_fill
        
        # Net salary row - Tổng Số Tiền Lương Thực Nhận = Tổng Cộng Thu Nhập - Tổng Cộng Khoản Trừ
        row += 1
        ws.merge_cells(f'A{row}:D{row}')
        ws[f'A{row}'] = 'Tổng Số Tiền Lương Thực Nhận'
        ws[f'A{row}'].font = Font(bold=True, size=12, color="FF0000")
        ws[f'A{row}'].border = border
        ws[f'A{row}'].fill = yellow_fill
        ws[f'A{row}'].alignment = Alignment(horizontal='center')
        ws[f'B{row}'].border = border
        ws[f'C{row}'].border = border
        ws[f'D{row}'].border = border
        
        ws[f'E{row}'] = format_number(slip_data['luong_thuc_nhan'])
        ws[f'E{row}'].font = Font(bold=True, size=12, color="FF0000")
        ws[f'E{row}'].border = border
        ws[f'E{row}'].fill = light_yellow_fill
        
        # Footer note
        row += 2
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = "Anh/Chị vui lòng kiểm tra lại thông tin trên phiếu lương. Mọi thắc mắc vui lòng liên hệ Phòng HCNS trong vòng"
        ws[f'A{row}'].font = Font(size=9, italic=True)
        
        row += 1
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = "24 giờ (kể từ thời điểm nhận được thông báo này) để được giải quyết."
        ws[f'A{row}'].font = Font(size=9, italic=True)
        
        row += 1
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = "Quá thời hạn trên, thông tin trên phiếu lương sẽ được xem là chính xác và không có khiếu nại. Trân trọng cảm ơn!"
        ws[f'A{row}'].font = Font(size=9, italic=True)
        
        # Save to bytes
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Clean filename with selected month and year
        safe_name = "".join(c for c in str(employee_name) if c.isalnum() or c in (' ', '_')).strip()
        filename = f"PhieuLuong_{safe_name}_Thang{month}_{year}.xlsx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e) + '\n' + traceback.format_exc()})


@app.route('/export/pdf/<int:employee_index>', methods=['GET', 'POST'])
def export_pdf(employee_index):
    """Export employee salary slip to PDF with specific template"""
    if data_store['thong_tin'] is None:
        return jsonify({'success': False, 'error': 'Không có dữ liệu'})
    
    try:
        df_info = data_store['thong_tin']
        df_luong = data_store['luong']
        
        if employee_index >= len(df_info):
            return jsonify({'success': False, 'error': 'Không tìm thấy nhân viên'})
        
        # Get month and year from request
        now = datetime.now()
        month = request.args.get('month', now.month, type=int)
        year = request.args.get('year', now.year, type=int)
        
        employee = df_info.iloc[employee_index]
        name_col = find_employee_name_column(df_info)
        employee_name = employee[name_col] if name_col else f'NhanVien_{employee_index}'
        
        # Get salary data
        salary_row = None
        if df_luong is not None and name_col:
            luong_name_col = find_employee_name_column(df_luong)
            if luong_name_col:
                salary_match = df_luong[df_luong[luong_name_col].astype(str).str.lower().str.strip() == str(employee_name).lower().strip()]
                if len(salary_match) > 0:
                    salary_row = salary_match.iloc[0]
        
        # Get salary slip data
        slip_data = get_salary_slip_data(employee, salary_row, df_info, df_luong)
        
        # Create PDF
        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm, leftMargin=1*cm, rightMargin=1*cm)
        
        elements = []
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            alignment=1,
            spaceAfter=20,
            textColor=colors.red
        )
        
        # Title with selected month and year
        elements.append(Paragraph(f"PHIEU LUONG THANG {month} NAM {year}", title_style))
        elements.append(Spacer(1, 10))
        
        # Info table
        info_data = [
            [remove_accents('Họ tên:'), remove_accents(slip_data['ho_ten']), remove_accents('Ngày công chuẩn:'), ''],
            [remove_accents('Lương thỏa thuận:'), format_number(slip_data['luong_thoa_thuan']), remove_accents('Ngày công thực tế:'), ''],
            [remove_accents('% Lương Thử việc:'), '', remove_accents('Nghỉ phép:'), ''],
            [remove_accents('Lương đóng:'), format_number(slip_data['luong_dong']), remove_accents('Tổng giờ tăng ca:'), ''],
            [remove_accents('Số tài khoản:'), slip_data['so_tai_khoan'], remove_accents('Tên ngân hàng:'), remove_accents(slip_data['ten_ngan_hang'])],
        ]
        
        info_table = Table(info_data, colWidths=[3.5*cm, 5*cm, 4*cm, 5*cm])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.yellow),
            ('BACKGROUND', (2, 0), (2, -1), colors.yellow),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(info_table)
        elements.append(Spacer(1, 15))
        
        # Main table
        table_header = [
            ['STT', remove_accents('Các Khoản Thu Nhập'), '', remove_accents('Các Khoản Trừ Vào Lương'), '']
        ]
        
        table_data = [
            ['1', remove_accents('Lương thực tế'), format_number(slip_data['luong_thuc_te']), 'BHXH', format_number(slip_data['bhxh'])],
            ['2', remove_accents('Phép năm'), '', remove_accents('Đoàn phí'), format_number(slip_data['doan_phi'])],
            ['3', remove_accents('Lương tăng ca'), '', remove_accents('Thuế Thu Nhập Cá Nhân'), format_number(slip_data['thue_tncn'])],
            ['4', remove_accents('Lương bổ sung'), '', remove_accents('Tạm Ứng'), ''],
            ['5', remove_accents('Giữ xe'), '', remove_accents('Tiền phạt'), ''],
            ['6', remove_accents('Công tác phí'), '', remove_accents('Khác'), ''],
        ]
        
        totals_data = [
            ['', remove_accents('Tổng Cộng Thu Nhập'), format_number(slip_data['tong_thu_nhap']), 
             remove_accents('Tổng Cộng Khoản Trừ'), format_number(slip_data['tong_khoan_tru'])],
        ]
        
        full_table_data = table_header + table_data + totals_data
        
        main_table = Table(full_table_data, colWidths=[1.5*cm, 4.5*cm, 3.5*cm, 4.5*cm, 3.5*cm])
        main_table.setStyle(TableStyle([
            # Header row
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#FFC000')),
            ('SPAN', (1, 0), (2, 0)),
            ('SPAN', (3, 0), (4, 0)),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            
            # Data rows
            ('BACKGROUND', (0, 1), (0, -2), colors.HexColor('#DAEEF3')),
            ('BACKGROUND', (1, 1), (1, -2), colors.HexColor('#DAEEF3')),
            ('BACKGROUND', (2, 1), (2, -2), colors.HexColor('#E2EFDA')),
            ('BACKGROUND', (3, 1), (3, -2), colors.HexColor('#DAEEF3')),
            ('BACKGROUND', (4, 1), (4, -2), colors.HexColor('#E2EFDA')),
            
            # Totals row
            ('BACKGROUND', (1, -1), (1, -1), colors.yellow),
            ('BACKGROUND', (3, -1), (3, -1), colors.yellow),
            ('BACKGROUND', (2, -1), (2, -1), colors.HexColor('#FFFFCC')),
            ('BACKGROUND', (4, -1), (4, -1), colors.HexColor('#FFFFCC')),
            
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(main_table)
        elements.append(Spacer(1, 5))
        
        # Net salary row - Tổng Số Tiền Lương Thực Nhận = Tổng Cộng Thu Nhập - Tổng Cộng Khoản Trừ
        net_data = [[remove_accents('Tổng Số Tiền Lương Thực Nhận'), format_number(slip_data['luong_thuc_nhan'])]]
        net_table = Table(net_data, colWidths=[14*cm, 3.5*cm])
        net_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), colors.yellow),
            ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#FFFFCC')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.red),
            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(net_table)
        
        # Footer note
        elements.append(Spacer(1, 20))
        note_style = ParagraphStyle('Note', parent=styles['Normal'], fontSize=8, italic=True)
        elements.append(Paragraph(remove_accents("Anh/Chị vui lòng kiểm tra lại thông tin trên phiếu lương. Mọi thắc mắc vui lòng liên hệ Phòng HCNS trong vòng"), note_style))
        elements.append(Paragraph(remove_accents("24 giờ (kể từ thời điểm nhận được thông báo này) để được giải quyết."), note_style))
        elements.append(Paragraph(remove_accents("Quá thời hạn trên, thông tin trên phiếu lương sẽ được xem là chính xác và không có khiếu nại. Trân trọng cảm ơn!"), note_style))
        
        # Build PDF
        doc.build(elements)
        output.seek(0)
        
        # Clean filename with selected month and year
        safe_name = "".join(c for c in str(employee_name) if c.isalnum() or c in (' ', '_')).strip()
        filename = f"PhieuLuong_{safe_name}_Thang{month}_{year}.pdf"
        
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )
    
    except Exception as e:
        import traceback
        return jsonify({'success': False, 'error': str(e) + '\n' + traceback.format_exc()})


@app.route('/get_columns')
def get_columns():
    """Get available columns for search"""
    return jsonify({
        'columns_info': data_store['columns_thong_tin'],
        'columns_luong': data_store['columns_luong'],
        'employees_list': data_store['employees_list']
    })


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
