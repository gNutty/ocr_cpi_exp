# AI OCR & Document Editor - Installation Guide

A cross-platform OCR application with Thai language support, powered by Typhoon AI.

## 🚀 Quick Start

### Option 1: Using Docker (Recommended for Deployment)

```bash
# Clone the repository
git clone https://github.com/gNutty/view_ocr.git
cd view_ocr

# Copy and configure environment
cp env.example .env
# Edit .env with your API key

# Run with Docker Compose
docker-compose up -d

# Access at http://localhost:8501
```

### Option 2: Local Installation

#### Prerequisites
- Python 3.9 or higher
- pip (Python package manager)

#### Step 1: Clone and Setup

```bash
# Clone the repository
git clone https://github.com/gNutty/view_ocr.git
cd view_ocr

# Copy configuration files
cp env.example .env
cp config.json.example config.json
```

#### Step 2: Install Python Dependencies

**Windows:**
```cmd
pip install -r requirements.txt
```

**Linux/macOS:**
```bash
chmod +x install.sh
./install.sh
```

Or manually:
```bash
pip install -r requirements.txt
```

#### Step 3: Install System Dependencies

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-tha
```

**CentOS/RHEL/Fedora:**
```bash
# CentOS/RHEL
sudo yum install poppler-utils tesseract tesseract-langpack-tha

# Fedora
sudo dnf install poppler-utils tesseract tesseract-langpack-tha
```

**macOS (with Homebrew):**
```bash
brew install poppler tesseract
```

**Windows:**
1. **Poppler:**
   - Download from: https://github.com/oschwartz10612/poppler-windows/releases/
   - Extract to: `C:\poppler\`
   - Add `C:\poppler\Library\bin` to PATH or set in config

2. **Tesseract OCR:**
   - Download from: https://github.com/UB-Mannheim/tesseract/wiki
   - Install to: `C:\Program Files\Tesseract-OCR\`
   - **Important:** Select Thai language pack during installation

#### Step 4: Configure API Key

**Using Environment Variable (Recommended):**
```bash
# Linux/macOS
export TYPHOON_API_KEY="your_api_key_here"

# Windows (Command Prompt)
set TYPHOON_API_KEY=your_api_key_here

# Windows (PowerShell)
$env:TYPHOON_API_KEY="your_api_key_here"
```

**Using config.json:**
```json
{
  "API_KEY": "your_api_key_here",
  "POPPLER_PATH": null
}
```

#### Step 5: Run the Application

**Windows:**
```cmd
runapp.bat
```

**Linux/macOS:**
```bash
chmod +x run.sh
./run.sh
```

**Or directly with Python:**
```bash
streamlit run app.py
```

Access the application at: http://localhost:8501

---

## 🔧 Configuration Options

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `TYPHOON_API_KEY` | API key for Typhoon OCR | (required) |
| `OLLAMA_API_URL` | Ollama API endpoint | `http://localhost:11434/api/generate` |
| `OCR_MODEL_NAME` | OCR model for local processing | `scb10x/typhoon-ocr1.5-3b:latest` |
| `POPPLER_PATH` | Path to Poppler binaries | Auto-detected |
| `TESSERACT_PATH` | Path to Tesseract executable | Auto-detected |
| `STREAMLIT_SERVER_PORT` | Streamlit server port | `8501` |

### Config File (config.json)

```json
{
  "API_KEY": "your_api_key_here",
  "POPPLER_PATH": null
}
```

---

## 📦 Optional: Local OCR with Ollama

For offline/private OCR processing using local AI models:

### Install Ollama

**Linux:**
```bash
curl -fsSL https://ollama.ai/install.sh | sh
```

**macOS:**
```bash
brew install ollama
```

**Windows:**
Download from: https://ollama.ai/

### Download OCR Model

```bash
ollama pull scb10x/typhoon-ocr1.5-3b:latest
```

### Start Ollama Server

```bash
ollama serve
```

---

## 🐳 Docker Deployment

### Build and Run

```bash
# Build image
docker build -t ocr-app .

# Run container
docker run -d \
  -p 8501:8501 \
  -e TYPHOON_API_KEY="your_api_key" \
  -v $(pwd)/source:/app/source \
  -v $(pwd)/output:/app/output \
  --name ocr-app \
  ocr-app
```

### Docker Compose

```bash
# Start all services
docker-compose up -d

# View logs
docker-compose logs -f

# Stop services
docker-compose down
```

---

## ☁️ GitHub Actions / CI-CD

The repository includes GitHub Actions workflows for:

1. **Testing** (`.github/workflows/python-app.yml`):
   - Runs on: Ubuntu, Windows, macOS
   - Python versions: 3.9, 3.10, 3.11, 3.12
   - Includes linting with flake8

2. **Deployment** (`.github/workflows/deploy-streamlit.yml`):
   - Creates deployment artifacts
   - Can be extended for cloud deployment

### Streamlit Cloud Deployment

1. Push code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io/)
3. Connect your GitHub repository
4. Set secrets in Streamlit Cloud dashboard:
   - `TYPHOON_API_KEY`

---

## 🔍 Troubleshooting

### Python Issues

**pip not working:**
```bash
python -m pip install --upgrade pip
```

**Module not found:**
```bash
pip install -r requirements.txt
```

### Poppler Issues

**pdf2image error on Linux:**
```bash
sudo apt-get install poppler-utils
```

**pdf2image error on Windows:**
1. Verify Poppler is installed at `C:\poppler\Library\bin`
2. Update `POPPLER_PATH` in config.json:
```json
{
  "POPPLER_PATH": "C:\\poppler\\Library\\bin"
}
```

### Tesseract Issues

**Tesseract not found:**
1. Verify installation path
2. Add to system PATH or set `TESSERACT_PATH` environment variable

**Thai language not working:**
- Ensure Thai language pack is installed
- Linux: `tesseract-ocr-tha`
- Windows: Select during installation

### Ollama Issues

**Connection refused:**
```bash
# Check if Ollama is running
curl http://localhost:11434/api/tags

# Start Ollama
ollama serve
```

**Model not found:**
```bash
ollama pull scb10x/typhoon-ocr1.5-3b:latest
```

### Docker Issues

**"[error] Error running 'docker info'" in VS Code/Cursor Output:**

ข้อผิดพลาดนี้มาจาก Docker extension ของ VS Code/Cursor ที่พยายามตรวจสอบสถานะ Docker โดยอัตโนมัติ โดยปกติไม่จำเป็นต้องแก้ไขถ้า:

1. **Docker ทำงานปกติ** - ตรวจสอบด้วย:
   ```bash
   # Windows (PowerShell)
   docker --version
   docker info
   ```
   
   ถ้าแสดงข้อมูล Docker แสดงว่าทำงานปกติ ✅

2. **ถ้าไม่ได้ใช้ Docker** - สามารถข้าม error นี้ได้ ไม่กระทบการใช้งานแอปพลิเคชัน

**ถ้าต้องการแก้ไข (ถ้า Docker ไม่ทำงาน):**

**Windows:**
1. เปิด **Docker Desktop** จาก Start Menu
2. รอจน Docker Desktop ทำงานเต็มที่ (ไอคอนวาฬจะไม่หมุน)
3. ถ้า Docker Desktop ไม่เริ่มทำงาน:
   - ตรวจสอบว่า WSL 2 ใช้งานได้: `wsl --status`
   - อัปเดต Docker Desktop เป็นเวอร์ชันล่าสุด
   - Restart เครื่องคอมพิวเตอร์

**ตรวจสอบ Docker:**
```bash
# ตรวจสอบ Docker version
docker --version

# ตรวจสอบ Docker daemon
docker info

# ทดสอบ Docker
docker run hello-world
```

**ถ้ายังมีปัญหา:**
- ตรวจสอบว่า Docker Desktop ทำงานอยู่ (ดูจาก System Tray)
- ตรวจสอบ WSL 2 backend: Docker Desktop → Settings → General → Use the WSL 2 based engine
- Restart Docker Desktop

**หมายเหตุ:** Error นี้เป็นเพียง warning จาก Docker extension และไม่กระทบการทำงานของแอปพลิเคชัน ถ้า Docker ทำงานปกติสามารถข้ามได้

---

## 📋 System Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| Python | 3.9 | 3.11+ |
| RAM | 4 GB | 8 GB+ |
| Disk Space | 1 GB | 5 GB (with Ollama) |
| OS | Windows 10, Ubuntu 20.04, macOS 12 | Latest LTS |

---

## 🤝 Support

For issues and questions:
- Create an issue on GitHub
- Include error messages and system info

**Getting System Info:**
```bash
python --version
pip list | grep -E "streamlit|pandas|pdf2image"
tesseract --version
```
