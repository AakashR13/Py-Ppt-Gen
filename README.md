# Py-Ppt-Gen

A Python-based PowerPoint presentation generator that uses LLM (Large Language Model) to automatically create presentations from text descriptions. Built with Streamlit for an interactive web interface and python-pptx for presentation generation.

![Language Distribution](https://img.shields.io/badge/Python-100%25-blue)

## Overview

Py-Ppt-Gen is an automated presentation generation tool that leverages LLM capabilities to create PowerPoint presentations from natural language descriptions. The tool provides a user-friendly Streamlit interface for generating presentations with customizable content and design templates.

## Features

- 🤖 **LLM-Powered Generation** - Uses LLM Foundry to generate presentation content
- 🎨 **Template Support** - Multiple presentation templates (Light modernist, Prismatic design)
- 📊 **Streamlit Interface** - Interactive web UI for easy presentation creation
- 📝 **Text-Based Input** - Generate presentations from simple text descriptions
- 🔄 **Batch Generation** - Support for generating multiple presentations

## Installation

### Prerequisites

- Python 3.7+
- LLM Foundry API token
- Required Python packages (install via pip)

### Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/AakashR13/Py-Ppt-Gen.git
   cd Py-Ppt-Gen
   ```

2. **Install dependencies:**
   ```bash
   pip install streamlit python-pptx
   ```

3. **Configure environment:**
   Create a `.env` file in the project root:
   ```env
   LLMF_TOK=<Your_LLM_Foundry_Token>
   ```

## Usage

### Running the Application

Start the Streamlit application:

```bash
streamlit run ./query-ppt.py
```

The application will open in your default web browser, typically at `http://localhost:8501`.

### Generating Presentations

1. Enter your presentation description in the text input field
2. Select a template (if available)
3. Click "Generate Presentation"
4. Download the generated `.pptx` file

## Project Structure

```
Py-Ppt-Gen/
├── query-ppt.py              # Streamlit application entry point
├── py-pptx-gen.py             # Core presentation generation logic
├── presentation_description.txt  # Example input description
├── .env                       # Environment variables (not in repo)
├── .gitignore                 # Git ignore file
└── README.md                  # This file
```

## Language Distribution

- **Python**: 100% - Complete implementation in Python

## Components

- **query-ppt.py** - Streamlit web interface for user interaction
- **py-pptx-gen.py** - Core logic for generating PowerPoint presentations using python-pptx
- **Templates** - Pre-designed presentation templates (Light modernist, Prismatic design)

## Configuration

The application requires an LLM Foundry API token to function. Add your token to the `.env` file:

```env
LLMF_TOK=your_token_here
```


## References

- [Streamlit Documentation](https://docs.streamlit.io/)

## License

This project is licensed under the **MIT License**.

See the [LICENSE](https://github.com/AakashR13/Py-Ppt-Gen/blob/main/LICENSE) file in the repository for full license details.

- [python-pptx Documentation](https://python-pptx.readthedocs.io/)

