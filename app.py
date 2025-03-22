from flask import Flask, render_template, request, jsonify, make_response, send_file, abort
import pandas as pd
import os
import numpy as np
import json
import io
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Sử dụng Agg backend để tránh lỗi trên server không có GUI
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime
import base64
import traceback

# Không nhập WeasyPrint trực tiếp, mà kiểm tra nó có sẵn không
PDF_AVAILABLE = False
try:
    from weasyprint import HTML, CSS
    PDF_AVAILABLE = True
    print("WeasyPrint available - PDF export enabled")
except ImportError:
    print("WeasyPrint not available - PDF export disabled")
    print("To enable PDF export, install GTK3 and WeasyPrint following the instructions at:")
    print("https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#installation")

app = Flask(__name__)

# Data loading
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')

def load_data():
    data = {
        'avg_by_code': pd.read_csv(os.path.join(DATA_DIR, 'Average_by_Code.csv'), encoding='utf-8'),
        'avg_by_sector': pd.read_csv(os.path.join(DATA_DIR, 'Average_by_Sector.csv'), encoding='utf-8'),
        'balance_sheet': pd.read_csv(os.path.join(DATA_DIR, 'BCDKT.csv'), encoding='utf-8'),
        'fin_statements': pd.read_csv(os.path.join(DATA_DIR, 'BCTC.csv'), encoding='utf-8'),
        'income_statement': pd.read_csv(os.path.join(DATA_DIR, 'KQKD.csv'), encoding='utf-8'),
        'cash_flow': pd.read_csv(os.path.join(DATA_DIR, 'LCTT.csv'), encoding='utf-8'),
        'disclosures': pd.read_csv(os.path.join(DATA_DIR, 'TM.csv'), encoding='utf-8')
    }
   
    # Try to load company info from Excel file if available
    try:
        thongtin_path = os.path.join(DATA_DIR, 'thongtin.xlsx')
        if os.path.exists(thongtin_path):
            data['company_info'] = pd.read_excel(thongtin_path)
            print("Company info loaded from thongtin.xlsx")
    except Exception as e:
        print(f"Error loading thongtin.xlsx: {e}")
   
    return data


# Load data once when the app starts
try:
    app_data = load_data()
    print("Data loaded successfully")
except Exception as e:
    print(f"Error loading data: {e}")
    app_data = {}


@app.route('/')
def index():
    # Get list of sectors for dropdown
    sectors = app_data['avg_by_sector']['Sector'].tolist() if 'avg_by_sector' in app_data else []
   
    # Get top performing companies by ROE
    top_companies = []
    if 'avg_by_code' in app_data:
        top_companies = app_data['avg_by_code'].sort_values(by='ROE (%)', ascending=False).head(10)[['Mã', 'ROE (%)']].to_dict(orient='records')
   
    # Get average financial ratios by sector
    sector_metrics = []
    if 'avg_by_sector' in app_data:
        sector_metrics = app_data['avg_by_sector'][['Sector', 'Average ROA', 'Average ROE', 'Average ROS']].head(5).to_dict(orient='records')
   
    # General market data
    market_data = {
        'total_companies': len(app_data['avg_by_code']) if 'avg_by_code' in app_data else 0,
        'avg_market_roe': app_data['avg_by_code']['ROE (%)'].mean() if 'avg_by_code' in app_data else 0,
        'avg_market_roa': app_data['avg_by_code']['ROA (%)'].mean() if 'avg_by_code' in app_data else 0,
        'sectors_count': len(sectors)
    }
   
    # Get all company codes for search autocomplete
    all_companies = []
    if 'fin_statements' in app_data:
        all_companies = sorted(app_data['fin_statements']['Mã'].unique().tolist())
   
    return render_template('index.html',
                           sectors=sectors,
                           top_companies=top_companies,
                           sector_metrics=sector_metrics,
                           market_data=market_data,
                           all_companies=all_companies)


@app.route('/api/companies')
def get_companies():
    """API endpoint for company search suggestions"""
    search_term = request.args.get('term', '').upper()
   
    # Get list of companies from fin_statements data
    companies = []
    if 'fin_statements' in app_data:
        # Get unique company codes and corresponding names
        df = app_data['fin_statements'][['Mã', 'Tên công ty']].drop_duplicates()
        companies = [
            {'value': row['Mã'], 'label': f"{row['Mã']} - {row['Tên công ty']}"}
            for _, row in df.iterrows()
        ]
       
        # Filter based on search term
        if search_term:
            companies = [
                company for company in companies
                if search_term in company['value'] or
                search_term.lower() in company['label'].lower()
            ]
       
        # Sort alphabetically by code
        companies = sorted(companies, key=lambda x: x['value'])
   
    return jsonify(companies)


@app.route('/sector_analysis')
def sector_analysis():
    # Get the sector parameter from the query string
    selected_sector = request.args.get('sector', None)
   
    # Get all available sectors for the dropdown
    sectors = app_data['avg_by_sector']['Sector'].tolist() if 'avg_by_sector' in app_data else []
   
    # If a sector is selected, get the sector data
    sector_info = {}
    sector_metrics = {}
    trend_analysis = {}
    trend_data = {}
    top_companies = []
    sector_comparisons = {}
   
    if selected_sector:
        # Get sector metrics
        sector_row = app_data['avg_by_sector'][app_data['avg_by_sector']['Sector'] == selected_sector]
        if not sector_row.empty:
            sector_metrics = {
                'Average_ROA': sector_row['Average ROA'].values[0],
                'Average_ROE': sector_row['Average ROE'].values[0],
                'Average_ROS': sector_row['Average ROS'].values[0],
                'Average_EBITDA_Margin': sector_row['Average EBITDA Margin'].values[0],
                'Average_D_E_Ratio': sector_row['Average D/E Ratio'].values[0]
            }
       
        # Get companies in this sector for the top companies list
        companies_in_sector = []
        if 'fin_statements' in app_data:
            sector_filter = app_data['fin_statements']['Ngành ICB - cấp 1'] == selected_sector
            companies_in_sector = app_data['fin_statements'][sector_filter]['Mã'].unique()
       
        # Get company metrics for companies in this sector
        if 'avg_by_code' in app_data and len(companies_in_sector) > 0:
            sector_companies = app_data['avg_by_code'][app_data['avg_by_code']['Mã'].isin(companies_in_sector)]
            # Sort by ROE and get top 10
            top_companies_df = sector_companies.sort_values(by='ROE (%)', ascending=False).head(10)
           
            # Format for the template
            top_companies = [
                {
                    'code': row['Mã'],
                    'name': row['Mã'],  # Would ideally get company name from another dataset
                    'roe': row['ROE (%)'],
                    'roa': row['ROA (%)'],
                    'ros': row['ROS (%)'],
                    'revenue_growth': row['Revenue Growth (%)']
                }
                for _, row in top_companies_df.iterrows()
            ]
       
        # Basic sector info
        sector_info = {
            'company_count': len(companies_in_sector),
            'market_share': 5.0,  # Placeholder value
            'description': f"Thông tin về ngành {selected_sector} sẽ được hiển thị ở đây."
        }
       
        # Placeholder trend analysis
        trend_analysis = {
            'profitability_comment': f"Phân tích xu hướng khả năng sinh lời của ngành {selected_sector} trong giai đoạn 2020-2024.",
            'liquidity_comment': f"Phân tích xu hướng thanh khoản của ngành {selected_sector} trong giai đoạn 2020-2024.",
            'leverage_comment': f"Phân tích xu hướng đòn bẩy của ngành {selected_sector} trong giai đoạn 2020-2024.",
            'growth_comment': f"Phân tích xu hướng tăng trưởng của ngành {selected_sector} trong giai đoạn 2020-2024."
        }
       
        # Placeholder trend data for charts
        trend_data = {
            'profitability': json.dumps({
                'labels': ['2020', '2021', '2022', '2023', '2024'],
                'roa': [5.2, 6.1, 5.8, 6.5, 7.0],
                'roe': [12.5, 13.2, 12.8, 14.0, 15.2],
                'ros': [8.3, 9.1, 8.7, 9.5, 10.2]
            }),
            'liquidity': json.dumps({
                'labels': ['2020', '2021', '2022', '2023', '2024'],
                'current_ratio': [1.8, 1.9, 1.85, 2.0, 2.1],
                'quick_ratio': [1.2, 1.3, 1.25, 1.4, 1.5]
            }),
            'leverage': json.dumps({
                'labels': ['2020', '2021', '2022', '2023', '2024'],
                'debt_to_assets': [48, 47, 46, 45, 44],
                'debt_to_equity': [92, 89, 85, 82, 79]
            }),
            'growth': json.dumps({
                'labels': ['2020', '2021', '2022', '2023', '2024'],
                'revenue_growth': [5, 12, 8, 15, 10],
                'net_income_growth': [7, 14, 9, 18, 12]
            })
        }
       
        # Placeholder sector comparisons
        sector_comparisons = {
            'comment': f"So sánh ngành {selected_sector} với các ngành khác.",
            'data': json.dumps([
                {'name': selected_sector, 'roa': 6.5, 'roe': 15.0, 'ros': 10.0, 'ebitda_margin': 18.0, 'revenue_growth': 12.0, 'debt_to_equity': 80.0},
                {'name': 'Ngành A', 'roa': 5, 'roe': 12, 'ros': 8, 'ebitda_margin': 15, 'revenue_growth': 10, 'debt_to_equity': 75},
                {'name': 'Ngành B', 'roa': 7, 'roe': 16, 'ros': 11, 'ebitda_margin': 20, 'revenue_growth': 15, 'debt_to_equity': 85},
                {'name': 'Ngành C', 'roa': 4, 'roe': 10, 'ros': 7, 'ebitda_margin': 14, 'revenue_growth': 8, 'debt_to_equity': 70},
                {'name': 'Ngành D', 'roa': 8, 'roe': 18, 'ros': 12, 'ebitda_margin': 22, 'revenue_growth': 16, 'debt_to_equity': 90}
            ])
        }
   
    return render_template('sector_analysis.html',
                           sectors=sectors,
                           selected_sector=selected_sector,
                           sector_info=sector_info,
                           sector_metrics=sector_metrics,
                           trend_analysis=trend_analysis,
                           trend_data=trend_data,
                           top_companies=top_companies,
                           sector_comparisons=sector_comparisons)


@app.route('/company_analysis')
def company_analysis():
    # Get the company code parameter from the query string
    selected_company_code = request.args.get('code', None)
   
    # Initialize variables to pass to the template
    company_info = {}
    company_metrics = {}
    financial_time_series = {}
    ratio_time_series = {}
    ratio_analysis = {}
    balance_sheet = {}
    income_statement = {}
    financial_statements = {}
    competitor_comparison = {}
    company_ranking = {}
    company_profile = ""
   
    if selected_company_code:
        # Get company data from fin_statements
        if 'fin_statements' in app_data:
            company_data = app_data['fin_statements'][app_data['fin_statements']['Mã'] == selected_company_code]
            if not company_data.empty:
                # Get company name and basic info
                company_name = company_data['Tên công ty'].iloc[0] if 'Tên công ty' in company_data.columns else f"Công ty {selected_company_code}"
                exchange = company_data['Sàn'].iloc[0] if 'Sàn' in company_data.columns else "N/A"
                icb_level1 = company_data['Ngành ICB - cấp 1'].iloc[0] if 'Ngành ICB - cấp 1' in company_data.columns else "N/A"
                icb_level2 = company_data['Ngành ICB - cấp 2'].iloc[0] if 'Ngành ICB - cấp 2' in company_data.columns else "N/A"
                icb_level3 = company_data['Ngành ICB - cấp 3'].iloc[0] if 'Ngành ICB - cấp 3' in company_data.columns else "N/A"
               
                # Set company_info
                company_info = {
                    'name': company_name,
                    'exchange': exchange,
                    'icb_level1': icb_level1,
                    'icb_level2': icb_level2,
                    'icb_level3': icb_level3,
                    'sector': icb_level1  # For backward compatibility with existing template
                }
               
                # Get company ranking within its industry (ICB level 3)
                if 'balance_sheet' in app_data and icb_level3 != "N/A":
                    # Filter companies in the same industry
                    industry_companies = app_data['balance_sheet'][
                        app_data['balance_sheet']['Ngành ICB - cấp 3'] == icb_level3
                    ]
                   
                    # Get the latest data for each company
                    latest_data = []
                    for company in industry_companies['Mã'].unique():
                        company_bs = industry_companies[industry_companies['Mã'] == company].sort_values(by=['Năm', 'Quý'], ascending=False)
                        if not company_bs.empty:
                            company_latest = company_bs.iloc[0]
                            if 'TỔNG CỘNG TÀI SẢN' in company_latest and not pd.isna(company_latest['TỔNG CỘNG TÀI SẢN']):
                                latest_data.append({
                                    'code': company,
                                    'total_assets': company_latest['TỔNG CỘNG TÀI SẢN']
                                })
                   
                    # Sort by total assets
                    latest_data.sort(key=lambda x: x['total_assets'], reverse=True)
                   
                    # Find the rank of selected company
                    company_rank = next((i+1 for i, item in enumerate(latest_data) if item['code'] == selected_company_code), 0)
                    total_companies = len(latest_data)
                   
                    company_ranking = {
                        'rank': company_rank,
                        'total': total_companies,
                        'industry': icb_level3
                    }
               
                # Get company profile information from thongtin.xlsx
                if 'company_info' in app_data:
                    company_detail = app_data['company_info'][app_data['company_info']['Mã CK'] == selected_company_code]


                    if not company_detail.empty:
                        # Check if 'Thông tin' column exists and extract profile information
                        if 'Thông tin' in company_detail.columns:
                            company_profile = company_detail['Thông tin'].iloc[0]
                            if pd.notna(company_profile):
                                company_info['profile'] = company_profile
                       
                        # Add additional details if available
                        for col in company_detail.columns:
                            if col not in ['Mã', 'Thông tin']:
                                value = company_detail[col].iloc[0]
                                if not pd.isna(value):
                                    company_info[col.lower().replace(' ', '_')] = value
       
        # Get company metrics from avg_by_code
        if 'avg_by_code' in app_data:
            company_row = app_data['avg_by_code'][app_data['avg_by_code']['Mã'] == selected_company_code]
            if not company_row.empty:
                # Extract metrics
                company_metrics = {
                    'ROA': company_row['ROA (%)'].values[0] if 'ROA (%)' in company_row.columns else 0,
                    'ROE': company_row['ROE (%)'].values[0] if 'ROE (%)' in company_row.columns else 0,
                    'ROS': company_row['ROS (%)'].values[0] if 'ROS (%)' in company_row.columns else 0,
                    'EBITDA_Margin': company_row['EBITDA Margin (%)'].values[0] if 'EBITDA Margin (%)' in company_row.columns else 0,
                    'Current_Ratio': company_row['Current Ratio'].values[0] if 'Current Ratio' in company_row.columns else 0,
                    'Quick_Ratio': company_row['Quick Ratio'].values[0] if 'Quick Ratio' in company_row.columns else 0,
                    'Debt_to_Assets': company_row['D/A (%)'].values[0] if 'D/A (%)' in company_row.columns else 0,
                    'Debt_to_Equity': company_row['D/E (%)'].values[0] if 'D/E (%)' in company_row.columns else 0,
                    'Equity_to_Assets': company_row['E/A (%)'].values[0] if 'E/A (%)' in company_row.columns else 0,
                    'Interest_Coverage': company_row['Interest Coverage Ratio'].values[0] if 'Interest Coverage Ratio' in company_row.columns else 0,
                    'Inventory_Turnover': company_row['Inventory Turnover'].values[0] if 'Inventory Turnover' in company_row.columns else 0,
                    'Receivables_Turnover': company_row['Accounts Receivable Turnover'].values[0] if 'Accounts Receivable Turnover' in company_row.columns else 0,
                    'Asset_Turnover': company_row['Total Asset Turnover'].values[0] if 'Total Asset Turnover' in company_row.columns else 0,
                    # Add comparison with sector averages (placeholder values for now)
                    'ROA_vs_sector': 1.5,
                    'ROE_vs_sector': 2.0,
                    'ROS_vs_sector': 1.8,
                    'EBITDA_Margin_vs_sector': 1.2,
                    'Current_Ratio_vs_sector': 0.2,
                    'Quick_Ratio_vs_sector': 0.1,
                    'Debt_to_Assets_vs_sector': -2.0,
                    'Debt_to_Equity_vs_sector': -3.0,
                    'Equity_to_Assets_vs_sector': 2.5,
                    'Interest_Coverage_vs_sector': 1.0,
                    'Inventory_Turnover_vs_sector': 0.5,
                    'Receivables_Turnover_vs_sector': 0.8,
                    'Asset_Turnover_vs_sector': 0.3
                }
       
        # Use the placeholder data for other variables if real data is not available
        if not financial_time_series:
            financial_time_series = json.dumps({
                'labels': ['2020', '2021', '2022', '2023', '2024'],
                'revenue': [1000, 1200, 1350, 1500, 1700],
                'net_profit': [100, 130, 150, 180, 210],
                'gross_profit': [400, 450, 500, 550, 620],
                'total_assets': [2000, 2200, 2500, 2800, 3100],
                'total_liabilities': [800, 850, 900, 950, 1000],
                'equity': [1200, 1350, 1600, 1850, 2100],
                'operating_cash_flow': [150, 170, 200, 220, 250],
                'investing_cash_flow': [-120, -150, -180, -200, -230],
                'financing_cash_flow': [-30, -20, -10, -15, -25]
            })
       
        # Placeholder for ratio time series
        ratio_time_series = json.dumps({
            'labels': ['2020', '2021', '2022', '2023', '2024'],
            'roa': [5.0, 5.5, 6.0, 6.5, 7.0],
            'roe': [12.0, 13.0, 14.0, 15.0, 16.0],
            'ros': [9.0, 10.0, 11.0, 12.0, 13.0],
            'ebitda_margin': [15.0, 16.0, 17.0, 18.0, 19.0],
            'current_ratio': [1.5, 1.6, 1.7, 1.8, 1.9],
            'quick_ratio': [1.1, 1.2, 1.3, 1.4, 1.5],
            'debt_to_assets': [40, 39, 38, 37, 36],
            'debt_to_equity': [65, 63, 61, 59, 57],
            'equity_to_assets': [60, 61, 62, 63, 64],
            'interest_coverage': [8, 9, 10, 11, 12],
            'inventory_turnover': [5, 5.2, 5.4, 5.6, 5.8],
            'receivables_turnover': [7, 7.3, 7.6, 7.9, 8.2],
            'asset_turnover': [0.8, 0.85, 0.9, 0.95, 1.0]
        })
       
        # Placeholder for ratio analysis
        ratio_analysis = {
            'profitability_comment': f"Công ty {selected_company_code} có khả năng sinh lời tốt hơn trung bình ngành.",
            'liquidity_comment': f"Thanh khoản của công ty {selected_company_code} ở mức khá, cao hơn trung bình ngành.",
            'leverage_comment': f"Cơ cấu vốn của công ty {selected_company_code} có tỷ lệ nợ thấp hơn trung bình ngành.",
            'efficiency_comment': f"Hiệu quả sử dụng tài sản của công ty {selected_company_code} tốt hơn trung bình ngành."
        }
       
        # Placeholder for balance sheet
        balance_sheet = {
            'assets': json.dumps({
                'labels': ['Tiền và tương đương tiền', 'Đầu tư tài chính ngắn hạn', 'Các khoản phải thu', 'Hàng tồn kho', 'Tài sản ngắn hạn khác', 'Tài sản dài hạn'],
                'values': [10, 15, 20, 25, 5, 25]
            }),
            'capital': json.dumps({
                'labels': ['Vốn chủ sở hữu', 'Nợ ngắn hạn', 'Nợ dài hạn'],
                'values': [60, 30, 10]
            })
        }
       
        # Placeholder for income statement
        income_statement = {
            'revenue': 1700000000000,  # 1.7 tỷ VNĐ
            'cogs': 1080000000000,  # 1.08 tỷ VNĐ
            'cogs_pct': 63.5,
            'gross_profit': 620000000000,  # 620 tỷ VNĐ
            'gross_margin': 36.5,
            'selling_expenses': 170000000000,  # 170 tỷ VNĐ
            'selling_expenses_pct': 10.0,
            'admin_expenses': 85000000000,  # 85 tỷ VNĐ
            'admin_expenses_pct': 5.0,
            'financial_revenue': 25000000000,  # 25 tỷ VNĐ
            'financial_revenue_pct': 1.5,
            'financial_expenses': 30000000000,  # 30 tỷ VNĐ
            'financial_expenses_pct': 1.8,
            'profit_before_tax': 360000000000,  # 360 tỷ VNĐ
            'profit_before_tax_margin': 21.2,
            'income_tax': 72000000000,  # 72 tỷ VNĐ
            'income_tax_pct': 4.2,
            'net_profit': 288000000000,  # 288 tỷ VNĐ
            'net_profit_margin': 17.0,
            'structure': json.dumps({
                'revenue': 1700,
                'expenses': 1412,
                'profit': 288
            })
        }
       
        # Placeholder for financial statements year
        financial_statements = {
            'year': 2024
        }
       
        # Placeholder for competitor comparison
        competitor_comparison = {
            'comment': f"So sánh {selected_company_code} với các đối thủ cạnh tranh trong cùng ngành.",
            'data': json.dumps([
                {'code': selected_company_code, 'roe': 16.0, 'roa': 7.0, 'ros': 13.0, 'ebitda_margin': 19.0, 'debt_to_equity': 57.0, 'revenue_growth': 13.0},
                {'code': 'VNM', 'roe': 14.5, 'roa': 6.5, 'ros': 12.0, 'ebitda_margin': 18.0, 'debt_to_equity': 55.0, 'revenue_growth': 12.0},
                {'code': 'MSN', 'roe': 15.0, 'roa': 6.8, 'ros': 12.5, 'ebitda_margin': 18.5, 'debt_to_equity': 56.0, 'revenue_growth': 12.5},
                {'code': 'FPT', 'roe': 17.0, 'roa': 7.5, 'ros': 14.0, 'ebitda_margin': 20.0, 'debt_to_equity': 58.0, 'revenue_growth': 14.0}
            ])
        }
   
    # Get all company codes for search autocomplete
    all_companies = []
    if 'fin_statements' in app_data:
        all_companies = sorted(app_data['fin_statements']['Mã'].unique().tolist())

    # Chuẩn bị dữ liệu doanh thu và lợi nhuận theo thời gian
    financial_revenue_series = json.dumps({'labels': [], 'revenue': []})
    financial_profit_series = json.dumps({'labels': [], 'profit': []})

    if selected_company_code and 'income_statement' in app_data:
        # Lọc dữ liệu của công ty được chọn
        company_income_data = app_data['income_statement'][app_data['income_statement']['Mã'] == selected_company_code]
        
        # Debug information
        print(f"Rows found for company {selected_company_code}: {len(company_income_data)}")
        if not company_income_data.empty:
            print(f"Columns available: {company_income_data.columns.tolist()}")
        
        # Sắp xếp theo năm và quý để tạo chuỗi thời gian
        if not company_income_data.empty:
            company_income_data = company_income_data.sort_values(by=['Năm', 'Quý'])
            
            # Check for columns individually to handle each chart separately
            if 'Doanh thu thuần' in company_income_data.columns:
                # Tạo nhãn thời gian (Năm-Quý)
                time_labels = [f"{int(row['Năm'])}-Q{int(row['Quý'])}" for _, row in company_income_data.iterrows()]
                
                # Trích xuất dữ liệu doanh thu (chuyển đổi sang tỷ VNĐ để hiển thị)
                revenue_data = [float(row['Doanh thu thuần'])/1e9 if pd.notna(row['Doanh thu thuần']) else 0 for _, row in company_income_data.iterrows()]
                
                # Tạo đối tượng JSON cho dữ liệu doanh thu
                financial_revenue_series = json.dumps({
                    'labels': time_labels,
                    'revenue': revenue_data
                })
                print(f"Revenue data prepared with {len(time_labels)} data points")
            else:
                print("Column 'Doanh thu thuần' not found in the data")
            
            # Try both column names for profit - sometimes columns might have different names or variations
            profit_column = None
            if 'Lợi nhuận sau thuế thu nhập doanh nghiệp' in company_income_data.columns:
                profit_column = 'Lợi nhuận sau thuế thu nhập doanh nghiệp'
            elif 'Lợi nhuận sau thuế thu nhập doanh nghiệp.1' in company_income_data.columns:
                profit_column = 'Lợi nhuận sau thuế thu nhập doanh nghiệp.1'
            
            if profit_column:
                # Tạo nhãn thời gian (Năm-Quý)
                time_labels = [f"{int(row['Năm'])}-Q{int(row['Quý'])}" for _, row in company_income_data.iterrows()]
                
                # Trích xuất dữ liệu lợi nhuận (chuyển đổi sang tỷ VNĐ để hiển thị)
                profit_data = [float(row[profit_column])/1e9 if pd.notna(row[profit_column]) else 0 for _, row in company_income_data.iterrows()]
                
                # Tạo đối tượng JSON cho dữ liệu lợi nhuận
                financial_profit_series = json.dumps({
                    'labels': time_labels,
                    'profit': profit_data
                })
                print(f"Profit data prepared with {len(time_labels)} data points using column {profit_column}")
            else:
                print("No profit column found in the data")
    
    
    return render_template('company_analysis.html',
                           selected_company_code=selected_company_code,
                           company_info=company_info,
                           company_metrics=company_metrics,
                           financial_time_series=financial_time_series,
                           ratio_time_series=ratio_time_series,
                           ratio_analysis=ratio_analysis,
                           balance_sheet=balance_sheet,
                           income_statement=income_statement,
                           financial_statements=financial_statements,
                           competitor_comparison=competitor_comparison,
                           company_ranking=company_ranking,
                           company_profile=company_profile,
                           all_companies=all_companies,
                           financial_revenue_series=financial_revenue_series,
                           financial_profit_series=financial_profit_series)


@app.route('/comparison')
def comparison():
    # Get comparison parameters
    comparison_type = request.args.get('type', None)
    company1 = request.args.get('company1', None)
    company2 = request.args.get('company2', None)
    company3 = request.args.get('company3', None)
    sector1 = request.args.get('sector1', None)
    sector2 = request.args.get('sector2', None)
    sector3 = request.args.get('sector3', None)
    company = request.args.get('company', None)
    sector = request.args.get('sector', None)
   
    # Get list of sectors for dropdowns
    sectors = app_data['avg_by_sector']['Sector'].tolist() if 'avg_by_sector' in app_data else []
   
    # Initialize comparison results
    comparison_results = {}
   
    if comparison_type:
        # Placeholder comparison results
        comparison_results = {
            'overview_comment': "Phân tích tổng quan về các chỉ số tài chính.",
            'overview': [
                {'name': 'Đối tượng 1', 'roe': 15.0, 'roa': 7.0, 'ros': 12.0, 'current_ratio': 1.8, 'debt_to_equity': 60.0},
                {'name': 'Đối tượng 2', 'roe': 14.0, 'roa': 6.5, 'ros': 11.0, 'current_ratio': 1.7, 'debt_to_equity': 58.0},
            ],
            'profitability_comment': "Phân tích khả năng sinh lời của các đối tượng so sánh.",
            'leverage_comment': "Phân tích cơ cấu vốn và thanh khoản của các đối tượng so sánh.",
            'growth_comment': "Phân tích tăng trưởng của các đối tượng so sánh.",
            'valuation_comment': "Phân tích định giá của các đối tượng so sánh.",
            'valuation': [
                {'name': 'Đối tượng 1', 'pe': 15.5, 'pb': 2.3, 'ps': 1.8, 'ev_ebitda': 8.2, 'dividend_yield': 3.5},
                {'name': 'Đối tượng 2', 'pe': 14.8, 'pb': 2.1, 'ps': 1.6, 'ev_ebitda': 7.8, 'dividend_yield': 3.2},
            ],
            'time_series': json.dumps({
                'labels': ['2020', '2021', '2022', '2023', '2024'],
                'entities': [
                    {
                        'name': 'Đối tượng 1',
                        'roa': [5.5, 6.0, 6.5, 7.0, 7.5],
                        'roe': [13.0, 14.0, 14.5, 15.0, 15.5],
                        'ros': [10.0, 10.5, 11.0, 11.5, 12.0],
                        'ebitda_margin': [17.0, 17.5, 18.0, 18.5, 19.0],
                        'debt_to_assets': [40, 39, 38, 37, 36],
                        'debt_to_equity': [65, 63, 62, 60, 58],
                        'interest_coverage': [7, 8, 9, 10, 11],
                        'current_ratio': [1.6, 1.7, 1.75, 1.8, 1.85],
                        'revenue_growth': [8, 10, 12, 11, 9],
                        'net_income_growth': [10, 12, 15, 13, 11],
                        'assets_growth': [7, 9, 10, 8, 7],
                        'equity_growth': [9, 11, 12, 10, 9],
                        'pe': [14, 15, 15.5, 16, 15.5],
                        'pb': [2.0, 2.1, 2.2, 2.3, 2.2],
                        'ps': [1.6, 1.7, 1.8, 1.9, 1.8],
                        'ev_ebitda': [7.5, 7.8, 8.0, 8.2, 8.0]
                    },
                    {
                        'name': 'Đối tượng 2',
                        'roa': [5.0, 5.5, 6.0, 6.5, 7.0],
                        'roe': [12.0, 12.5, 13.0, 13.5, 14.0],
                        'ros': [9.0, 9.5, 10.0, 10.5, 11.0],
                        'ebitda_margin': [16.0, 16.5, 17.0, 17.5, 18.0],
                        'debt_to_assets': [42, 41, 40, 39, 38],
                        'debt_to_equity': [68, 66, 64, 60, 58],
                        'interest_coverage': [6, 7, 8, 9, 10],
                        'current_ratio': [1.5, 1.6, 1.65, 1.7, 1.75],
                        'revenue_growth': [7, 9, 11, 10, 8],
                        'net_income_growth': [9, 11, 13, 12, 10],
                        'assets_growth': [6, 8, 9, 7, 6],
                        'equity_growth': [8, 10, 11, 9, 8],
                        'pe': [13, 14, 14.5, 15, 14.8],
                        'pb': [1.8, 1.9, 2.0, 2.1, 2.0],
                        'ps': [1.4, 1.5, 1.6, 1.7, 1.6],
                        'ev_ebitda': [7.0, 7.3, 7.5, 7.8, 7.5]
                    }
                ]
            }),
            'historical': json.dumps({
                'labels': ['2020', '2021', '2022', '2023', '2024'],
                'entities': [
                    {
                        'name': 'Đối tượng 1',
                        'revenue': [1000, 1100, 1230, 1365, 1485],
                        'net_profit': [120, 140, 165, 190, 210],
                        'total_assets': [2000, 2180, 2375, 2565, 2745],
                        'equity': [1200, 1320, 1450, 1595, 1740],
                        'operating_cash_flow': [150, 170, 195, 215, 235],
                        'free_cash_flow': [100, 115, 135, 150, 165]
                    },
                    {
                        'name': 'Đối tượng 2',
                        'revenue': [900, 980, 1080, 1185, 1280],
                        'net_profit': [100, 115, 135, 155, 170],
                        'total_assets': [1800, 1950, 2125, 2295, 2460],
                        'equity': [1000, 1090, 1200, 1310, 1430],
                        'operating_cash_flow': [120, 135, 155, 170, 185],
                        'free_cash_flow': [80, 90, 105, 115, 125]
                    }
                ]
            }),
            'strengths_weaknesses': [
                {
                    'name': 'Đối tượng 1',
                    'strengths': ['Biên lợi nhuận cao', 'Tăng trưởng doanh thu ổn định', 'Cơ cấu vốn tốt'],
                    'weaknesses': ['Chi phí vận hành cao', 'Doanh thu từ thị trường quốc tế thấp']
                },
                {
                    'name': 'Đối tượng 2',
                    'strengths': ['Quản lý chi phí hiệu quả', 'Khả năng thanh toán tốt', 'Đa dạng sản phẩm'],
                    'weaknesses': ['Tốc độ tăng trưởng chậm', 'Tỷ suất sinh lời thấp hơn trung bình ngành']
                }
            ],
            'conclusion': "Kết luận tổng thể về so sánh các đối tượng. Dựa trên các chỉ số tài chính và xu hướng phát triển, Đối tượng 1 có kết quả kinh doanh tốt hơn về mặt khả năng sinh lời, trong khi Đối tượng 2 có lợi thế về tính thanh khoản và cơ cấu vốn."
        }
       
        # Adjust entity names based on comparison type
        if comparison_type == 'companies':
            # Companies comparison
            if company1:
                comparison_results['overview'][0]['name'] = company1
                comparison_results['valuation'][0]['name'] = company1
                comparison_results['time_series'] = json.loads(comparison_results['time_series'])
                comparison_results['time_series']['entities'][0]['name'] = company1
                comparison_results['time_series'] = json.dumps(comparison_results['time_series'])
               
                comparison_results['historical'] = json.loads(comparison_results['historical'])
                comparison_results['historical']['entities'][0]['name'] = company1
                comparison_results['historical'] = json.dumps(comparison_results['historical'])
               
                comparison_results['strengths_weaknesses'][0]['name'] = company1
           
            if company2:
                comparison_results['overview'][1]['name'] = company2
                comparison_results['valuation'][1]['name'] = company2
                comparison_results['time_series'] = json.loads(comparison_results['time_series'])
                comparison_results['time_series']['entities'][1]['name'] = company2
                comparison_results['time_series'] = json.dumps(comparison_results['time_series'])
               
                comparison_results['historical'] = json.loads(comparison_results['historical'])
                comparison_results['historical']['entities'][1]['name'] = company2
                comparison_results['historical'] = json.dumps(comparison_results['historical'])
               
                comparison_results['strengths_weaknesses'][1]['name'] = company2
       
        elif comparison_type == 'sectors':
            # Sectors comparison
            if sector1:
                comparison_results['overview'][0]['name'] = sector1
                comparison_results['valuation'][0]['name'] = sector1
                comparison_results['time_series'] = json.loads(comparison_results['time_series'])
                comparison_results['time_series']['entities'][0]['name'] = sector1
                comparison_results['time_series'] = json.dumps(comparison_results['time_series'])
               
                comparison_results['historical'] = json.loads(comparison_results['historical'])
                comparison_results['historical']['entities'][0]['name'] = sector1
                comparison_results['historical'] = json.dumps(comparison_results['historical'])
               
                comparison_results['strengths_weaknesses'][0]['name'] = sector1
           
            if sector2:
                comparison_results['overview'][1]['name'] = sector2
                comparison_results['valuation'][1]['name'] = sector2
                comparison_results['time_series'] = json.loads(comparison_results['time_series'])
                comparison_results['time_series']['entities'][1]['name'] = sector2
                comparison_results['time_series'] = json.dumps(comparison_results['time_series'])
               
                comparison_results['historical'] = json.loads(comparison_results['historical'])
                comparison_results['historical']['entities'][1]['name'] = sector2
                comparison_results['historical'] = json.dumps(comparison_results['historical'])
               
                comparison_results['strengths_weaknesses'][1]['name'] = sector2
       
        elif comparison_type == 'company_with_sector':
            # Company with sector comparison
            if company:
                comparison_results['overview'][0]['name'] = company
                comparison_results['valuation'][0]['name'] = company
                comparison_results['time_series'] = json.loads(comparison_results['time_series'])
                comparison_results['time_series']['entities'][0]['name'] = company
                comparison_results['time_series'] = json.dumps(comparison_results['time_series'])
               
                comparison_results['historical'] = json.loads(comparison_results['historical'])
                comparison_results['historical']['entities'][0]['name'] = company
                comparison_results['historical'] = json.dumps(comparison_results['historical'])
               
                comparison_results['strengths_weaknesses'][0]['name'] = company
           
            if sector:
                comparison_results['overview'][1]['name'] = sector
                comparison_results['valuation'][1]['name'] = sector
                comparison_results['time_series'] = json.loads(comparison_results['time_series'])
                comparison_results['time_series']['entities'][1]['name'] = sector
                comparison_results['time_series'] = json.dumps(comparison_results['time_series'])
               
                comparison_results['historical'] = json.loads(comparison_results['historical'])
                comparison_results['historical']['entities'][1]['name'] = sector
                comparison_results['historical'] = json.dumps(comparison_results['historical'])
               
                comparison_results['strengths_weaknesses'][1]['name'] = sector
   
    # Get all company codes for search autocomplete
    all_companies = []
    if 'fin_statements' in app_data:
        all_companies = sorted(app_data['fin_statements']['Mã'].unique().tolist())
   
    return render_template('comparison.html',
                           sectors=sectors,
                           comparison_type=comparison_type,
                           company1=company1,
                           company2=company2,
                           company3=company3,
                           sector1=sector1,
                           sector2=sector2,
                           sector3=sector3,
                           company=company,
                           sector=sector,
                           comparison_results=comparison_results,
                           selected_sector=sector,  # For the company_with_sector form
                           all_companies=all_companies)


@app.route('/api/sector_data/<sector>')
def sector_data(sector):
    # API endpoint for fetching sector data
    if 'avg_by_sector' in app_data:
        sector_data = app_data['avg_by_sector'][app_data['avg_by_sector']['Sector'] == sector]
        if not sector_data.empty:
            return jsonify(sector_data.iloc[0].to_dict())
    return jsonify({})


@app.route('/api/company_data/<code>')
def company_data(code):
    # API endpoint for fetching company data
    if 'avg_by_code' in app_data:
        company_data = app_data['avg_by_code'][app_data['avg_by_code']['Mã'] == code]
        if not company_data.empty:
            return jsonify(company_data.iloc[0].to_dict())
    return jsonify({})


# Thêm route mới để hiển thị trang xuất báo cáo
@app.route('/export_report')
def export_report_page():
    try:
        # Lấy danh sách công ty từ dữ liệu
        companies = []
        if 'fin_statements' in app_data:
            # Lấy mã công ty và tên công ty không trùng lặp
            df = app_data['fin_statements'][['Mã', 'Tên công ty']].drop_duplicates()
            companies = [
                {'code': row['Mã'], 'name': row['Tên công ty']}
                for _, row in df.iterrows()
            ]
            # Sắp xếp theo mã công ty
            companies = sorted(companies, key=lambda x: x['code'])
        
        return render_template('export_report.html', companies=companies, pdf_available=PDF_AVAILABLE)
    except Exception as e:
        print(f"Error in export_report_page: {e}")
        traceback.print_exc()
        return "Đã xảy ra lỗi khi tải trang xuất báo cáo. Vui lòng thử lại sau.", 500

# Thêm route để xử lý việc xuất báo cáo
@app.route('/generate_report/<company_code>')
def generate_report(company_code):
    try:
        # Kiểm tra tham số công ty
        if not company_code:
            return "Không tìm thấy mã công ty", 400
        
        # Lấy dữ liệu công ty
        company_data = get_company_report_data(company_code)
        
        # Nếu không tìm thấy dữ liệu công ty
        if not company_data:
            return "Không tìm thấy dữ liệu công ty", 404
        
        # Thêm thông tin về khả năng xuất PDF
        company_data['pdf_available'] = PDF_AVAILABLE
        
        # Render template báo cáo với dữ liệu
        rendered_template = render_template('report_template.html', **company_data)
        
        # Xử lý các thông số request
        output_format = request.args.get('format', 'html')
        
        if output_format == 'pdf' and PDF_AVAILABLE:
            try:
                # Tạo PDF từ HTML
                pdf = HTML(string=rendered_template).write_pdf()
                
                # Tạo response với file PDF
                response = make_response(pdf)
                response.headers['Content-Type'] = 'application/pdf'
                response.headers['Content-Disposition'] = f'attachment; filename=Bao_Cao_{company_code}.pdf'
                
                return response
            except Exception as e:
                print(f"Error generating PDF: {e}")
                traceback.print_exc()
                return "Đã xảy ra lỗi khi tạo PDF. Vui lòng thử lại hoặc chọn định dạng HTML.", 500
        elif output_format == 'pdf' and not PDF_AVAILABLE:
            # Thông báo người dùng rằng PDF không khả dụng
            return "Xuất file PDF không khả dụng do thiếu thư viện WeasyPrint. Vui lòng cài đặt WeasyPrint theo hướng dẫn hoặc sử dụng định dạng HTML.", 501
        else:
            # Trả về HTML
            return rendered_template
    except Exception as e:
        # Thay vì trả về chuỗi chung chung, ta log hoặc trả về hẳn thông tin lỗi
        print("========== LỖI CHI TIẾT ==========")
        traceback.print_exc()
        print("=================================")

        # Hoặc tạm thời trả exception ra cho dễ debug (chỉ dùng khi DEV)
        return f"Lỗi cụ thể: {repr(e)}", 500

# Hàm lấy dữ liệu cho báo cáo
def get_company_report_data(company_code):
    result = {
        'company_code': company_code,
        'report_date': datetime.now().strftime('%d/%m/%Y'),
    }
    
    try:
        # Kiểm tra dữ liệu có tồn tại
        if 'fin_statements' not in app_data or 'balance_sheet' not in app_data or 'income_statement' not in app_data:
            print("Missing required data sets")
            return None
        
        # Lấy thông tin cơ bản về công ty
        company_info = app_data['fin_statements'][app_data['fin_statements']['Mã'] == company_code]
        if company_info.empty:
            print(f"No company info found for {company_code}")
            return None
        
        # Lấy thông tin tên công ty
        result['company_name'] = company_info['Tên công ty'].iloc[0]
        result['exchange'] = company_info['Sàn'].iloc[0] if 'Sàn' in company_info.columns else "N/A"
        
        # Lấy thông tin ngành
        result['industry_level1'] = company_info['Ngành ICB - cấp 1'].iloc[0] if 'Ngành ICB - cấp 1' in company_info.columns else "N/A"
        result['industry_level2'] = company_info['Ngành ICB - cấp 2'].iloc[0] if 'Ngành ICB - cấp 2' in company_info.columns else "N/A"
        result['industry_level3'] = company_info['Ngành ICB - cấp 3'].iloc[0] if 'Ngành ICB - cấp 3' in company_info.columns else "N/A"
        
        # Lấy dữ liệu bảng cân đối kế toán
        balance_sheet = app_data['balance_sheet'][app_data['balance_sheet']['Mã'] == company_code].sort_values(by=['Năm', 'Quý'])
        
        # Lấy dữ liệu báo cáo kết quả kinh doanh
        income_statement = app_data['income_statement'][app_data['income_statement']['Mã'] == company_code].sort_values(by=['Năm', 'Quý'])
        
        # Lấy dữ liệu lưu chuyển tiền tệ
        cash_flow = app_data['cash_flow'][app_data['cash_flow']['Mã'] == company_code].sort_values(by=['Năm', 'Quý']) if 'cash_flow' in app_data else pd.DataFrame()
        
        # Lấy dữ liệu trung bình ngành
        if 'avg_by_sector' in app_data and result['industry_level1'] != "N/A":
            sector_avg = app_data['avg_by_sector'][app_data['avg_by_sector']['Sector'] == result['industry_level1']]
            if not sector_avg.empty:
                result['sector_avg'] = sector_avg.iloc[0].to_dict()
        
        # Lấy dữ liệu chỉ số tài chính của công ty
        if 'avg_by_code' in app_data:
            company_metrics = app_data['avg_by_code'][app_data['avg_by_code']['Mã'] == company_code]
            if not company_metrics.empty:
                result['company_metrics'] = company_metrics.iloc[0].to_dict()
        
        # Chuẩn bị dữ liệu báo cáo tài chính theo năm
        years = []
        
        # Lấy các năm từ dữ liệu bảng cân đối kế toán
        if not balance_sheet.empty:
            years.extend(balance_sheet['Năm'].unique())
        
        # Lấy các năm từ dữ liệu kết quả kinh doanh
        if not income_statement.empty:
            years.extend(income_statement['Năm'].unique())
        
        # Loại bỏ trùng lặp và sắp xếp
        years = sorted(list(set(years)))
        
        financial_data = {}
        
        for year in years[-5:]:  # Lấy 5 năm gần nhất nếu có
            year_data = {
                'balance_sheet': {},
                'income_statement': {},
                'cash_flow': {}
            }
            
            # Lấy dữ liệu bảng cân đối kế toán của năm cuối cùng
            year_bs = balance_sheet[balance_sheet['Năm'] == year].sort_values(by='Quý', ascending=False)
            if not year_bs.empty:
                last_quarter_bs = year_bs.iloc[0]
                year_data['balance_sheet'] = {
                    'total_assets': last_quarter_bs['TỔNG CỘNG TÀI SẢN'] if 'TỔNG CỘNG TÀI SẢN' in last_quarter_bs and pd.notna(last_quarter_bs['TỔNG CỘNG TÀI SẢN']) else 0,
                    'current_assets': last_quarter_bs['TÀI SẢN NGẮN HẠN'] if 'TÀI SẢN NGẮN HẠN' in last_quarter_bs and pd.notna(last_quarter_bs['TÀI SẢN NGẮN HẠN']) else 0,
                    'fixed_assets': last_quarter_bs['TÀI SẢN DÀI HẠN'] if 'TÀI SẢN DÀI HẠN' in last_quarter_bs and pd.notna(last_quarter_bs['TÀI SẢN DÀI HẠN']) else 0,
                    'liabilities': last_quarter_bs['NỢ PHẢI TRẢ'] if 'NỢ PHẢI TRẢ' in last_quarter_bs and pd.notna(last_quarter_bs['NỢ PHẢI TRẢ']) else 0,
                    'equity': last_quarter_bs['VỐN CHỦ SỞ HỮU'] if 'VỐN CHỦ SỞ HỮU' in last_quarter_bs and pd.notna(last_quarter_bs['VỐN CHỦ SỞ HỮU']) else 0,
                }
            
            # Lấy dữ liệu báo cáo kết quả kinh doanh của năm
            year_is = income_statement[income_statement['Năm'] == year].sort_values(by='Quý', ascending=False)
            if not year_is.empty:
                last_quarter_is = year_is.iloc[0]
                year_data['income_statement'] = {
                    'revenue': last_quarter_is['Doanh thu thuần'] if 'Doanh thu thuần' in last_quarter_is and pd.notna(last_quarter_is['Doanh thu thuần']) else 0,
                    'gross_profit': last_quarter_is['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ'] if 'Lợi nhuận gộp về bán hàng và cung cấp dịch vụ' in last_quarter_is and pd.notna(last_quarter_is['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ']) else 0,
                    'operating_profit': last_quarter_is['Lợi nhuận thuần từ hoạt động kinh doanh'] if 'Lợi nhuận thuần từ hoạt động kinh doanh' in last_quarter_is and pd.notna(last_quarter_is['Lợi nhuận thuần từ hoạt động kinh doanh']) else 0,
                    'profit_before_tax': last_quarter_is['Tổng lợi nhuận kế toán trước thuế'] if 'Tổng lợi nhuận kế toán trước thuế' in last_quarter_is and pd.notna(last_quarter_is['Tổng lợi nhuận kế toán trước thuế']) else 0,
                    'net_profit': last_quarter_is['Lợi nhuận sau thuế thu nhập doanh nghiệp'] if 'Lợi nhuận sau thuế thu nhập doanh nghiệp' in last_quarter_is and pd.notna(last_quarter_is['Lợi nhuận sau thuế thu nhập doanh nghiệp']) else 0,
                }
            
            # Lấy dữ liệu báo cáo lưu chuyển tiền tệ của năm
            if not cash_flow.empty:
                year_cf = cash_flow[cash_flow['Năm'] == year].sort_values(by='Quý', ascending=False)
                if not year_cf.empty:
                    last_quarter_cf = year_cf.iloc[0]
                    year_data['cash_flow'] = {
                        'operating_cash_flow': last_quarter_cf['Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh (TT)'] if 'Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh (TT)' in last_quarter_cf and pd.notna(last_quarter_cf['Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh (TT)']) else 0,
                        'investing_cash_flow': last_quarter_cf['Lưu chuyển tiền tệ ròng từ hoạt động đầu tư (TT)'] if 'Lưu chuyển tiền tệ ròng từ hoạt động đầu tư (TT)' in last_quarter_cf and pd.notna(last_quarter_cf['Lưu chuyển tiền tệ ròng từ hoạt động đầu tư (TT)']) else 0,
                        'financing_cash_flow': last_quarter_cf['Lưu chuyển tiền tệ từ hoạt động tài chính (TT)'] if 'Lưu chuyển tiền tệ từ hoạt động tài chính (TT)' in last_quarter_cf and pd.notna(last_quarter_cf['Lưu chuyển tiền tệ từ hoạt động tài chính (TT)']) else 0,
                    }
            
            financial_data[str(int(year))] = year_data
        
        result['financial_data'] = financial_data
        result['years'] = [str(int(year)) for year in years[-5:]]
        
        # Tính toán chỉ số tài chính theo năm
        financial_ratios = {}
        
        # Trích xuất dữ liệu cho năm gần nhất để đảm bảo chúng ta có dữ liệu
        if not years:
            print(f"No financial years data found for {company_code}")
            return result
            
        for year in years[-5:]:
            year_bs = balance_sheet[balance_sheet['Năm'] == year].sort_values(by='Quý', ascending=False)
            year_is = income_statement[income_statement['Năm'] == year].sort_values(by='Quý', ascending=False)
            
            if not year_bs.empty and not year_is.empty:
                bs = year_bs.iloc[0]
                is_data = year_is.iloc[0]
                
                ratios = {}
                
                # ROA
                try:
                    if 'TỔNG CỘNG TÀI SẢN' in bs and pd.notna(bs['TỔNG CỘNG TÀI SẢN']) and bs['TỔNG CỘNG TÀI SẢN'] != 0 and 'Lợi nhuận sau thuế thu nhập doanh nghiệp' in is_data and pd.notna(is_data['Lợi nhuận sau thuế thu nhập doanh nghiệp']):
                        ratios['ROA'] = (is_data['Lợi nhuận sau thuế thu nhập doanh nghiệp'] / bs['TỔNG CỘNG TÀI SẢN']) * 100
                    else:
                        ratios['ROA'] = 0
                except Exception as e:
                    print(f"Error calculating ROA for {company_code}, year {year}: {e}")
                    ratios['ROA'] = 0
                
                # ROE
                try:
                    if 'VỐN CHỦ SỞ HỮU' in bs and pd.notna(bs['VỐN CHỦ SỞ HỮU']) and bs['VỐN CHỦ SỞ HỮU'] != 0 and 'Lợi nhuận sau thuế thu nhập doanh nghiệp' in is_data and pd.notna(is_data['Lợi nhuận sau thuế thu nhập doanh nghiệp']):
                        ratios['ROE'] = (is_data['Lợi nhuận sau thuế thu nhập doanh nghiệp'] / bs['VỐN CHỦ SỞ HỮU']) * 100
                    else:
                        ratios['ROE'] = 0
                except Exception as e:
                    print(f"Error calculating ROE for {company_code}, year {year}: {e}")
                    ratios['ROE'] = 0
                
                # ROS
                try:
                    if 'Doanh thu thuần' in is_data and pd.notna(is_data['Doanh thu thuần']) and is_data['Doanh thu thuần'] != 0 and 'Lợi nhuận sau thuế thu nhập doanh nghiệp' in is_data and pd.notna(is_data['Lợi nhuận sau thuế thu nhập doanh nghiệp']):
                        ratios['ROS'] = (is_data['Lợi nhuận sau thuế thu nhập doanh nghiệp'] / is_data['Doanh thu thuần']) * 100
                    else:
                        ratios['ROS'] = 0
                except Exception as e:
                    print(f"Error calculating ROS for {company_code}, year {year}: {e}")
                    ratios['ROS'] = 0
                
                # Gross Profit Margin
                try:
                    if 'Doanh thu thuần' in is_data and pd.notna(is_data['Doanh thu thuần']) and is_data['Doanh thu thuần'] != 0 and 'Lợi nhuận gộp về bán hàng và cung cấp dịch vụ' in is_data and pd.notna(is_data['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ']):
                        ratios['Gross_Profit_Margin'] = (is_data['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ'] / is_data['Doanh thu thuần']) * 100
                    else:
                        ratios['Gross_Profit_Margin'] = 0
                except Exception as e:
                    print(f"Error calculating Gross Profit Margin for {company_code}, year {year}: {e}")
                    ratios['Gross_Profit_Margin'] = 0
                
                # Current Ratio
                try:
                    if 'Nợ ngắn hạn' in bs and pd.notna(bs['Nợ ngắn hạn']) and bs['Nợ ngắn hạn'] != 0 and 'TÀI SẢN NGẮN HẠN' in bs and pd.notna(bs['TÀI SẢN NGẮN HẠN']):
                        ratios['Current_Ratio'] = bs['TÀI SẢN NGẮN HẠN'] / bs['Nợ ngắn hạn']
                    else:
                        ratios['Current_Ratio'] = 0
                except Exception as e:
                    print(f"Error calculating Current Ratio for {company_code}, year {year}: {e}")
                    ratios['Current_Ratio'] = 0
                
                # Debt to Equity
                try:
                    if 'VỐN CHỦ SỞ HỮU' in bs and pd.notna(bs['VỐN CHỦ SỞ HỮU']) and bs['VỐN CHỦ SỞ HỮU'] != 0 and 'NỢ PHẢI TRẢ' in bs and pd.notna(bs['NỢ PHẢI TRẢ']):
                        ratios['Debt_to_Equity'] = (bs['NỢ PHẢI TRẢ'] / bs['VỐN CHỦ SỞ HỮU']) * 100
                    else:
                        ratios['Debt_to_Equity'] = 0
                except Exception as e:
                    print(f"Error calculating Debt to Equity for {company_code}, year {year}: {e}")
                    ratios['Debt_to_Equity'] = 0
                
                financial_ratios[str(int(year))] = ratios
        
        result['financial_ratios'] = financial_ratios
        
        # Tạo biểu đồ
        try:
            prepare_financial_charts(result)
        except Exception as e:
            print(f"Error preparing financial charts for {company_code}: {e}")
            traceback.print_exc()
            # Khởi tạo các trường biểu đồ rỗng để template vẫn hoạt động
            result['revenue_profit_chart'] = None
            result['ratios_chart'] = None
            result['balance_sheet_chart'] = None
            result['comparison_chart'] = None
        
        return result
    
    except Exception as e:
        print(f"Error in get_company_report_data for {company_code}: {e}")
        traceback.print_exc()
        return None

# Hàm tạo biểu đồ cho báo cáo
def prepare_financial_charts(data):
    years = data.get('years', [])
    if not years:
        raise ValueError("No years data available for charts")
    
    # Thiết lập style cho biểu đồ
    plt.style.use('seaborn-v0_8-whitegrid')
    
    # Cấu hình font
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans']
    
    # Đặt chất lượng biểu đồ
    plt.rcParams['figure.dpi'] = 100
    plt.rcParams['savefig.dpi'] = 100
    
    # Dữ liệu doanh thu và lợi nhuận
    try:
        # Dữ liệu doanh thu
        revenue_data = [data.get('financial_data', {}).get(year, {}).get('income_statement', {}).get('revenue', 0)/1e9 for year in years]
        
        # Dữ liệu lợi nhuận
        profit_data = [data.get('financial_data', {}).get(year, {}).get('income_statement', {}).get('net_profit', 0)/1e9 for year in years]
        
        # Tạo biểu đồ doanh thu và lợi nhuận
        plt.figure(figsize=(10, 6))
        width = 0.35
        x = np.arange(len(years))
        
        plt.bar(x - width/2, revenue_data, width, label='Doanh thu thuần', color='#3498db')
        plt.bar(x + width/2, profit_data, width, label='Lợi nhuận sau thuế', color='#2ecc71')
        
        plt.xlabel('Năm', fontsize=12)
        plt.ylabel('Tỷ VNĐ', fontsize=12)
        plt.title('Doanh thu và Lợi nhuận qua các năm', fontsize=14, fontweight='bold')
        plt.xticks(x, years, fontsize=10)
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        
        # Thêm số liệu lên biểu đồ
        for i, v in enumerate(revenue_data):
            plt.text(i - width/2, v + 0.1, f'{v:.1f}', ha='center', fontsize=9)
        
        for i, v in enumerate(profit_data):
            plt.text(i + width/2, v + 0.1, f'{v:.1f}', ha='center', fontsize=9)
        
        # Lưu biểu đồ vào buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        
        # Chuyển đổi hình ảnh sang base64 để hiển thị trong HTML
        data['revenue_profit_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
    except Exception as e:
        print(f"Error creating revenue/profit chart: {e}")
        traceback.print_exc()
        data['revenue_profit_chart'] = None
    
    # Tạo biểu đồ ROA, ROE, ROS
    try:
        plt.figure(figsize=(10, 6))
        
        roa_data = [data.get('financial_ratios', {}).get(year, {}).get('ROA', 0) for year in years]
        roe_data = [data.get('financial_ratios', {}).get(year, {}).get('ROE', 0) for year in years]
        ros_data = [data.get('financial_ratios', {}).get(year, {}).get('ROS', 0) for year in years]
        
        plt.plot(years, roa_data, marker='o', label='ROA (%)', linewidth=2, color='#3498db')
        plt.plot(years, roe_data, marker='s', label='ROE (%)', linewidth=2, color='#e74c3c')
        plt.plot(years, ros_data, marker='^', label='ROS (%)', linewidth=2, color='#2ecc71')
        
        plt.xlabel('Năm', fontsize=12)
        plt.ylabel('Phần trăm (%)', fontsize=12)
        plt.title('Chỉ số ROA, ROE, ROS qua các năm', fontsize=14, fontweight='bold')
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        
        # Thêm số liệu lên biểu đồ
        for i, v in enumerate(roa_data):
            plt.text(i, v, f'{v:.1f}%', ha='center', va='bottom', fontsize=9)
        
        for i, v in enumerate(roe_data):
            plt.text(i, v, f'{v:.1f}%', ha='center', va='bottom', fontsize=9)
        
        for i, v in enumerate(ros_data):
            plt.text(i, v, f'{v:.1f}%', ha='center', va='top', fontsize=9)
        
        # Lưu biểu đồ vào buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        
        # Chuyển đổi hình ảnh sang base64 để hiển thị trong HTML
        data['ratios_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
    except Exception as e:
        print(f"Error creating ratios chart: {e}")
        traceback.print_exc()
        data['ratios_chart'] = None
    
    # Tạo biểu đồ cơ cấu tài sản, nợ, vốn
    try:
        plt.figure(figsize=(10, 6))
        
        assets_data = [data.get('financial_data', {}).get(year, {}).get('balance_sheet', {}).get('total_assets', 0)/1e9 for year in years]
        liabilities_data = [data.get('financial_data', {}).get(year, {}).get('balance_sheet', {}).get('liabilities', 0)/1e9 for year in years]
        equity_data = [data.get('financial_data', {}).get(year, {}).get('balance_sheet', {}).get('equity', 0)/1e9 for year in years]
        
        x = np.arange(len(years))
        width = 0.25
        
        plt.bar(x - width, assets_data, width, label='Tổng tài sản', color='#3498db')
        plt.bar(x, liabilities_data, width, label='Nợ phải trả', color='#e74c3c')
        plt.bar(x + width, equity_data, width, label='Vốn chủ sở hữu', color='#2ecc71')
        
        plt.xlabel('Năm', fontsize=12)
        plt.ylabel('Tỷ VNĐ', fontsize=12)
        plt.title('Cơ cấu tài sản, nợ, vốn qua các năm', fontsize=14, fontweight='bold')
        plt.xticks(x, years, fontsize=10)
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        
        # Thêm số liệu lên biểu đồ
        for i, v in enumerate(assets_data):
            if v > 0:
                plt.text(i - width, v, f'{v:.1f}', ha='center', va='bottom', fontsize=9)
        
        for i, v in enumerate(liabilities_data):
            if v > 0:
                plt.text(i, v, f'{v:.1f}', ha='center', va='bottom', fontsize=9)
        
        for i, v in enumerate(equity_data):
            if v > 0:
                plt.text(i + width, v, f'{v:.1f}', ha='center', va='bottom', fontsize=9)
        
        # Lưu biểu đồ vào buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        
        # Chuyển đổi hình ảnh sang base64 để hiển thị trong HTML
        data['balance_sheet_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
    except Exception as e:
        print(f"Error creating balance sheet chart: {e}")
        traceback.print_exc()
        data['balance_sheet_chart'] = None
    
    # Tạo biểu đồ so sánh với trung bình ngành
    try:
        if 'company_metrics' in data and 'sector_avg' in data:
            plt.figure(figsize=(10, 6))
            
            metrics = ['ROA (%)', 'ROE (%)', 'ROS (%)']
            company_values = [
                data['company_metrics'].get('ROA (%)', 0),
                data['company_metrics'].get('ROE (%)', 0),
                data['company_metrics'].get('ROS (%)', 0)
            ]
            
            sector_values = [
                data['sector_avg'].get('Average ROA', 0),
                data['sector_avg'].get('Average ROE', 0),
                data['sector_avg'].get('Average ROS', 0)
            ]
            
            x = np.arange(len(metrics))
            width = 0.35
            
            plt.bar(x - width/2, company_values, width, label=data['company_code'], color='#3498db')
            plt.bar(x + width/2, sector_values, width, label='Trung bình ngành', color='#e74c3c')
            
            plt.xlabel('Chỉ số', fontsize=12)
            plt.ylabel('Phần trăm (%)', fontsize=12)
            plt.title('So sánh với trung bình ngành', fontsize=14, fontweight='bold')
            plt.xticks(x, metrics, fontsize=10)
            plt.legend(fontsize=10)
            plt.grid(True, linestyle='--', alpha=0.7)
            
            # Thêm số liệu lên biểu đồ
            for i, v in enumerate(company_values):
                plt.text(i - width/2, v + 0.3, f'{v:.1f}%', ha='center', fontsize=9)
            
            for i, v in enumerate(sector_values):
                plt.text(i + width/2, v + 0.3, f'{v:.1f}%', ha='center', fontsize=9)
            
            # Lưu biểu đồ vào buffer
            buf = io.BytesIO()
            plt.savefig(buf, format='png', bbox_inches='tight')
            buf.seek(0)
            
            # Chuyển đổi hình ảnh sang base64 để hiển thị trong HTML
            data['comparison_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
            plt.close()
    except Exception as e:
        print(f"Error creating comparison chart: {e}")
        traceback.print_exc()
        data['comparison_chart'] = None


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)