# Báo Cáo Phân Tích Tài Chính Chuyên Nghiệp

Ứng dụng web phân tích tài chính các doanh nghiệp niêm yết trên thị trường chứng khoán Việt Nam trong giai đoạn 2020-2024.

## Tổng Quan

Đây là đồ án cuối kỳ phát triển ứng dụng web phân tích tài chính chuyên nghiệp, cho phép người dùng xem và phân tích các chỉ số tài chính của các doanh nghiệp niêm yết trên các sàn chứng khoán tại Việt Nam.

### Tính Năng Chính

- **Tổng quan thị trường**: Hiển thị các số liệu thống kê tổng quan về thị trường
- **Phân tích ngành**: Phân tích các chỉ số tài chính theo từng ngành
- **Phân tích công ty**: Phân tích chi tiết tình hình tài chính của từng công ty
- **So sánh**: So sánh các chỉ số tài chính giữa các công ty hoặc các ngành

## Cài Đặt

### Yêu Cầu

- Python 3.7+
- Flask
- Pandas
- NumPy
- Chart.js

### Cài Đặt Môi Trường

1. Clone repository

```
git clone <repository-url>
cd financial_analysis_app
```

2. Tạo và kích hoạt môi trường ảo

```
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate  # Windows
```

3. Cài đặt các gói phụ thuộc

```
pip install -r requirements.txt
```

4. Khởi chạy ứng dụng

```
python app.py
```

5. Mở trình duyệt và truy cập `http://localhost:5000`

## Cấu Trúc Thư Mục

```
financial_analysis_app/
├── app.py                 # Main Flask application
├── config.py              # Configuration settings
├── models/                # Database models
│   └── __init__.py
├── static/                # Static files (CSS, JS, images)
│   ├── css/
│   │   ├── bootstrap.min.css
│   │   └── style.css
│   ├── js/
│   │   ├── bootstrap.min.js
│   │   ├── chart.js
│   │   └── main.js
│   └── img/
├── templates/             # HTML templates
│   ├── base.html
│   ├── index.html
│   ├── sector_analysis.html
│   ├── company_analysis.html
│   └── comparison.html
├── utils/                 # Utility functions
│   ├── __init__.py
│   ├── data_loader.py
│   └── financial_calculations.py
└── data/                  # CSV data files
    ├── Average_by_Code.csv
    ├── Average_by_Sector.csv
    ├── BCDKT.csv
    ├── BCTC.csv
    ├── KQKD.csv
    ├── LCTT.csv
    └── TM.csv
    thongtin.xlsx

## Nguồn Dữ Liệu

Dữ liệu được lấy từ báo cáo tài chính, báo cáo lưu chuyển tiền tệ, kết quả hoạt động kinh doanh của các doanh nghiệp đã niêm yết trên các sàn chứng khoán tại Việt Nam trong giai đoạn 2020-2024.

## Phát Triển

Dự án này được phát triển như một đồ án cuối kỳ của nhóm sinh viên. Mọi đóng góp đều được hoan nghênh.

## Giấy Phép

Dự án này được phân phối dưới giấy phép MIT.

## Liên Hệ

Nếu có bất kỳ câu hỏi nào, vui lòng liên hệ qua email: [email@example.com]
```
