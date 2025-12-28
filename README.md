# 📄 AI OCR & Document Editor

[![Python Application](https://github.com/gNutty/view_ocr/actions/workflows/python-app.yml/badge.svg)](https://github.com/gNutty/view_ocr/actions/workflows/python-app.yml)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.28+-FF4B4B.svg)](https://streamlit.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful cross-platform OCR application with Thai language support, built with Streamlit and powered by Typhoon AI.

## ✨ Features

- 🔍 **PDF OCR Processing** - Extract text from PDF invoices and documents
- 🇹🇭 **Thai Language Support** - Full support for Thai text recognition
- 📊 **Excel Export** - Export extracted data to Excel with vendor mapping
- 🖥️ **Web Interface** - Modern Streamlit-based user interface
- 🔄 **Dual OCR Modes**:
  - **API Mode**: Use Typhoon OCR cloud API
  - **Local Mode**: Use Ollama with local AI models
- 🐳 **Docker Ready** - Easy deployment with Docker and Docker Compose
- ⚡ **Cross-Platform** - Works on Windows, Linux, and macOS

## 📸 Screenshots

| OCR Processing | Document Editor |
|----------------|-----------------|
| Process PDF invoices | Edit and verify extracted data |

## 🚀 Quick Start

### Option 1: Docker (Recommended)

```bash
git clone https://github.com/gNutty/view_ocr.git
cd view_ocr
docker-compose up -d
```

Access at: http://localhost:8501

### Option 2: Local Installation

```bash
# Clone repository
git clone https://github.com/gNutty/view_ocr.git
cd view_ocr

# Install dependencies
pip install -r requirements.txt

# Configure API key
cp config.json.example config.json
# Edit config.json with your API key

# Run application
streamlit run app.py
```

See [INSTALLATION.md](INSTALLATION.md) for detailed instructions.

## 📋 Requirements

- Python 3.9+
- Poppler (for PDF processing)
- Tesseract OCR (optional, for text positioning)
- Ollama (optional, for local OCR)

## 🔧 Configuration

### Environment Variables

```bash
export TYPHOON_API_KEY="your_api_key"
export OLLAMA_API_URL="http://localhost:11434/api/generate"
```

### Config File

```json
{
  "API_KEY": "your_api_key",
  "POPPLER_PATH": null
}
```

## 📁 Project Structure

```
view_ocr/
├── app.py                  # Main Streamlit application
├── Extract_Inv.py          # API-based OCR processing
├── Extract_Inv_local.py    # Local OCR processing (Ollama)
├── Vendor_branch.xlsx      # Vendor master data
├── config.json             # Application configuration
├── requirements.txt        # Python dependencies
├── packages.txt            # System packages (Streamlit Cloud)
├── Dockerfile              # Docker build file
├── docker-compose.yml      # Docker Compose configuration
├── run.sh                  # Linux/macOS run script
├── runapp.bat              # Windows run script
└── INSTALLATION.md         # Installation guide
```

## 🐳 Docker Deployment

```bash
# Build and run
docker-compose up -d

# With Ollama for local OCR
docker-compose --profile ollama up -d

# View logs
docker-compose logs -f ocr-app
```

## ☁️ Cloud Deployment

### Streamlit Cloud

1. Fork this repository
2. Connect to [Streamlit Cloud](https://share.streamlit.io/)
3. Add secrets in dashboard:
   - `TYPHOON_API_KEY`

### GitHub Actions

CI/CD workflows are included for:
- Multi-platform testing (Ubuntu, Windows, macOS)
- Python 3.9-3.12 compatibility
- Code linting with flake8

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [Typhoon AI](https://opentyphoon.ai/) - OCR API
- [Streamlit](https://streamlit.io/) - Web framework
- [Ollama](https://ollama.ai/) - Local AI runtime
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) - Open source OCR engine

## 📧 Contact

For questions and support, please open an issue on GitHub.

---

Made with ❤️ by the OCR Team


