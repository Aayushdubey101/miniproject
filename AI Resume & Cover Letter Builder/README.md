# AI Resume & Cover Letter Builder

An AI-powered application built with **Streamlit** that helps users generate professional resumes and cover letters using the **Groq API** for intelligent content generation.

---

## 🚀 Features

- **AI-Powered Content Generation**  
  Uses Groq’s LLaMA-based models to generate high-quality, tailored resumes and cover letters based on user input.

- **User-Friendly Interface**  
  Built with Streamlit to provide a clean, simple, and interactive web-based experience.

- **PDF Export**  
  Automatically generates and downloads professional PDF documents using **FPDF**.

- **Customizable Resume Sections**  
  Supports multiple resume sections including:
  - Career Objective  
  - Education  
  - Projects  
  - Skills  
  - Experience  
  - Extracurricular Activities  

- **Multiple Resume Styles**  
  Choose from different resume formats based on your target role:
  - Standard  
  - MNC  
  - Startup  
  - Government  
  - Academic  

---

## 🛠 Installation

### 1. Clone the Repository
```bash
git clone <repository-url>
cd <repository-directory>

## Installation

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd <repository-directory>
   ```

2. **Set up Python environment**:
   - Ensure you have Python 3.8+ installed.
   - Use the specified Python version from `.python-version` if using pyenv.

3. **Install dependencies**:
   - Install dependencies using uv (recommended):
     ```bash
     uv sync
     ```
   - Or using pip:
     ```bash
     pip install -r requirements.txt
     ```

4. **Set up Groq API**:
   - Obtain an API key from [Groq](https://groq.com/).
   - Set the API key in your environment:
     ```bash
     export GROQ_API_KEY="your-api-key-here"
     ```

## Usage

1. **Run the application**:
   ```bash
   streamlit run app.py
   ```

2. **Access the app**:
   - Open your browser and navigate to the local Streamlit URL (usually `http://localhost:8501`).

3. **Generate Resume**:
   - Fill in your personal details, education, skills, projects, and other information.
   - Select a resume style.
   - Click "🚀 Generate Resume" to create and download your resume PDF.

4. **Generate Cover Letter**:
   - Provide the necessary details for the cover letter.
   - Click "✉ Generate Cover Letter" to create and download your cover letter PDF.

## Project Structure

- `app.py`: Main Streamlit application for the AI resume and cover letter builder.
- `main.py`: Simple hello world script (for testing purposes).
- `utils.py`: Utility functions for PDF generation and text cleaning.
- `pyproject.toml`: Project configuration and dependencies.
- `uv.lock`: Lock file for dependency management.
- `.gitignore`: Git ignore rules.
- `.python-version`: Specified Python version for the project.

## Dependencies

- **Streamlit**: For building the web interface.
- **Groq**: For AI-powered content generation.
- **FPDF**: For PDF document creation.
- **Unicodedata**: For text processing and cleaning.

## Troubleshooting

- **API Key Issues**: Ensure your Groq API key is correctly set in the environment variables.
- **Dependency Errors**: Try reinstalling dependencies with `uv sync`.
- **PDF Generation Errors**: Check that FPDF is properly installed and there are no Unicode issues in the input text.
- **Streamlit Not Running**: Verify that Streamlit is installed and you're using the correct command to run the app.

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Make your changes and test thoroughly.
4. Submit a pull request with a clear description of your changes.



## Disclaimer

This tool uses AI to generate content. Always review and customize the generated resumes and cover letters to ensure they accurately represent your qualifications and experiences.
