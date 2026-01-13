# 📊 Quản lý Nhân sự & Xuất Phiếu Lương

A Flask-based web application for managing employee information and generating salary slips with full Vietnamese language support.

![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)
![Flask](https://img.shields.io/badge/Flask-2.3+-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## ✨ Features

- **📤 Excel Upload**: Import employee data from Excel files (`.xlsx`, `.xls`) with automatic sheet detection
- **🔍 Employee Search**: Search employees by name or any field in the data
- **📄 Salary Slip Generation**: Generate professional salary slips in both **PDF** and **Excel** formats
- **🇻🇳 Vietnamese Support**: Full Vietnamese character support in PDF documents using custom fonts
- **📧 Email Integration**: Send salary slips directly to employees via SMTP (supports Gmail)
- **📦 Bulk Operations**: Download all salary slips as ZIP or send emails to all employees at once
- **📊 Email Status Tracking**: Track which employees have received their salary slips
- **⚙️ Configurable SMTP**: Runtime configuration for email server settings

## 🖼️ Screenshots

The application provides:

- Clean upload interface for Excel files
- Employee selection dropdown with email status indicators
- Individual employee detail view with export options
- Bulk operations panel for mass exports/emails
- Email configuration panel with Gmail App Password support

## 🚀 Getting Started

### Prerequisites

- Python 3.11 or higher
- pip (Python package manager)

### Installation

1. **Clone the repository**

   ```bash
   git clone <repository-url>
   cd excel_processing
   ```

2. **Create a virtual environment** (recommended)

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**

   ```bash
   python app.py
   ```

5. **Open your browser** and navigate to `http://localhost:5001`

## 📁 Project Structure

```
excel_processing/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── Procfile              # Heroku deployment config
├── render.yaml           # Render deployment config
├── runtime.txt           # Python version specification
├── fonts/
│   └── ArialUnicode.ttf  # Vietnamese font for PDF generation
├── static/
│   └── style.css         # Application styles
├── templates/
│   └── index.html        # Main HTML template
└── exports/              # Directory for exported files
```

## 📋 Excel File Format

The application expects an Excel file with two sheets:

### Sheet 1: "Thông tin" (Employee Information)

Contains employee personal information including:

- STT (Employee number)
- Họ tên (Full name)
- Email
- Bank account information
- Other personal details

### Sheet 2: "Lương" (Salary)

Contains salary information including:

- Lương thỏa thuận (Agreed salary)
- Lương cơ bản (Basic salary)
- BHXH (Social insurance)
- Thuế TNCN (Personal income tax)
- Other salary components

## 📧 Email Configuration

### Using Gmail

1. **Enable 2-Factor Authentication** on your Google account
2. **Generate an App Password**:
   - Go to [Google Account Security](https://myaccount.google.com/security)
   - Select "2-Step Verification" → "App passwords"
   - Generate a new app password for "Mail"
3. **Configure in the app**:
   - SMTP Server: `smtp.gmail.com`
   - Port: `587`
   - Email: Your Gmail address
   - Password: The generated App Password (not your Gmail password)

### Environment Variables (Optional)

You can also configure email via environment variables:

```bash
export SMTP_SERVER=smtp.gmail.com
export SMTP_PORT=587
export SENDER_EMAIL=your-email@gmail.com
export SENDER_PASSWORD=your-app-password
```

## 🌐 API Endpoints

| Endpoint             | Method | Description                       |
| -------------------- | ------ | --------------------------------- |
| `/`                  | GET    | Main application page             |
| `/upload`            | POST   | Upload Excel file                 |
| `/search`            | POST   | Search for employees              |
| `/get_employee/<id>` | GET    | Get single employee details       |
| `/export/excel/<id>` | GET    | Download Excel salary slip        |
| `/export/pdf/<id>`   | GET    | Download PDF salary slip          |
| `/export/bulk`       | POST   | Download all slips as ZIP         |
| `/send_email/<id>`   | POST   | Send email to single employee     |
| `/send_email_bulk`   | POST   | Send emails to multiple employees |
| `/email_status`      | GET    | Get email sending status          |
| `/configure_email`   | POST   | Update email configuration        |
| `/get_columns`       | GET    | Get available data columns        |

## 🚢 Deployment

### Deploy to Render

The project includes a `render.yaml` configuration file for easy deployment:

1. Connect your repository to [Render](https://render.com)
2. Create a new Web Service
3. Render will automatically detect the configuration
4. Set environment variables for email configuration

### Deploy to Heroku

```bash
heroku create your-app-name
git push heroku main
heroku config:set SENDER_EMAIL=your-email@gmail.com
heroku config:set SENDER_PASSWORD=your-app-password
```

## 🔧 Configuration Options

| Variable          | Default        | Description                        |
| ----------------- | -------------- | ---------------------------------- |
| `PORT`            | 5001           | Application port                   |
| `FLASK_DEBUG`     | True           | Debug mode                         |
| `SMTP_SERVER`     | smtp.gmail.com | SMTP server address                |
| `SMTP_PORT`       | 587            | SMTP server port                   |
| `SENDER_EMAIL`    | -              | Sender email address               |
| `SENDER_PASSWORD` | -              | Sender email password/app password |

## 📦 Dependencies

- **Flask** - Web framework
- **pandas** - Data manipulation and Excel reading
- **openpyxl** - Excel file creation
- **reportlab** - PDF generation
- **gunicorn** - Production WSGI server
- **flask-mail** - Email support

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Vietnamese font support via Arial Unicode
- ReportLab for excellent PDF generation capabilities
- Flask community for the amazing web framework
