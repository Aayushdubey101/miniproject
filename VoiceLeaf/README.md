# VoiceLeaf 🎧📄

Transform your PDF documents into audio with just a few clicks. VoiceLeaf uses advanced Text-to-Speech technology to convert any PDF file into high-quality audio, making your documents accessible anytime, anywhere.

## ✨ Features

- **PDF to Audio Conversion** - Convert any text-based PDF into natural-sounding audio
- **Multiple Voice Options** - Choose from various voices and accents
- **Adjustable Speech Rate** - Control the speed of narration to match your preference
- **Batch Processing** - Convert multiple PDFs at once
- **OCR Support** - Extract text from scanned PDFs and images
- **Audio Format Options** - Export to MP3, WAV, or other popular formats
- **Bookmark Support** - Resume listening from where you left off
- **Chapter Detection** - Automatically identifies document sections for easy navigation

## 🚀 Getting Started

### Prerequisites

- Python 3.8 or higher
- pip package manager
- Internet connection (for cloud TTS services)

### Installation

1. Clone the repository
```bash
git clone https://github.com/Aayushdubey101/miniproject.git
cd miniproject/VoiceLeaf
```

2. Install dependencies
```bash
pip install -r requirements.txt
```

3. Configure API keys (if using cloud TTS)
```bash
cp .env.example .env
# Edit .env and add your API keys
```

### Quick Start
```bash
# Basic usage
python voiceleaf.py input.pdf

# With custom voice and speed
python voiceleaf.py input.pdf --voice en-US-Neural2-A --speed 1.2

# Batch conversion
python voiceleaf.py folder/*.pdf --output audio_files/
```

## 📖 Usage

### Command Line Interface
```bash
voiceleaf [OPTIONS] PDF_FILE
```

**Options:**
- `-o, --output` - Specify output audio file path
- `-v, --voice` - Select voice (default: en-US-Standard-A)
- `-s, --speed` - Speech rate (0.5 to 2.0, default: 1.0)
- `-f, --format` - Audio format (mp3, wav, ogg)
- `--ocr` - Enable OCR for scanned PDFs
- `--chapters` - Split audio by chapters

### Web Interface

Launch the web application:
```bash
python app.py
```

Navigate to `http://localhost:5000` in your browser.

### API Usage
```python
from voiceleaf import PDFConverter

converter = PDFConverter(voice='en-US-Neural2-A', speed=1.0)
converter.convert('document.pdf', output='output.mp3')
```

## 🛠️ Configuration

Create a `config.yaml` file to customize default settings:
```yaml
tts:
  provider: google  # google, amazon, azure, local
  voice: en-US-Neural2-A
  speed: 1.0
  pitch: 0

audio:
  format: mp3
  bitrate: 192k
  
processing:
  chunk_size: 5000
  enable_ocr: false
  detect_chapters: true
```

## 📦 Supported Formats

### Input
- PDF (text-based and scanned)
- Multi-page documents
- Password-protected PDFs (with password)

### Output
- MP3 (recommended)
- WAV
- OGG
- FLAC

## 🎯 Use Cases

- **Accessibility** - Make documents accessible for visually impaired users
- **Multitasking** - Listen to documents while commuting or exercising
- **Learning** - Improve retention by listening to educational materials
- **Productivity** - Process information faster with audio playback
- **Language Learning** - Hear proper pronunciation of foreign language texts

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Text-to-Speech engines: Google Cloud TTS, Amazon Polly, Azure Cognitive Services
- PDF processing libraries
- Open-source community contributors

## 📧 Contact

Aayush Dubey - [@Aayushdubey101](https://github.com/Aayushdubey101)

Project Link: [https://github.com/Aayushdubey101/miniproject/tree/main/VoiceLeaf](https://github.com/Aayushdubey101/miniproject/tree/main/VoiceLeaf)

## 🐛 Known Issues

- Very large PDFs (>500 pages) may take significant time to process
- Complex mathematical formulas may not read naturally
- Some PDF formatting may not be preserved in audio

## 🗺️ Roadmap

- [ ] Mobile app support (iOS/Android)
- [ ] Real-time streaming conversion
- [ ] Custom voice training
- [ ] Multi-language support enhancement
- [ ] Cloud storage integration
- [ ] Podcast-style formatting
- [ ] Audio enhancement filters

## 💡 Tips

- Use headings in your PDFs for better chapter detection
- For scanned documents, ensure high image quality for better OCR results
- Experiment with different voices to find your preferred narration style
- Use slower speeds for technical or complex content

## ⭐ Show Your Support

If you find VoiceLeaf helpful, please consider giving it a star on GitHub!

---

Made with ❤️ by Aayush Dubey