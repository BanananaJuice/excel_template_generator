# Excel Template Inserter

A modern Python application for Excel file manipulation with a clean, modern UI.

## Features

- Modern web interface built with Streamlit
- Excel file reading and creation
- Template management
- Easy-to-use interface

## Setup

1. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: .\venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
streamlit run src/main.py
```

## Project Structure

```
├── src/
│   ├── main.py              # Main application entry point
│   ├── utils/               # Utility functions
│   │   ├── excel_ops.py     # Excel operations
│   │   └── config.py        # Configuration
│   └── frontend/            # Frontend components
│       └── components.py    # UI components
├── requirements.txt         # Project dependencies
└── README.md               # Project documentation
``` 