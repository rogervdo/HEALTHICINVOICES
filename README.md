# Hanova Excel Viewer

A Streamlit application to view and interact with Excel files from the `hanovaexcel` folder.

## Features

- 📊 Display multiple Excel files in separate tabs
- 📄 Support for multi-sheet Excel files with sheet selection
- 📥 Download functionality for each Excel file
- 📱 Responsive design with wide layout
- ⚡ Fast data loading and display

## Setup

1. Install the required dependencies:

```bash
pip install -r requirements.txt
```

2. Run the Streamlit app:

```bash
streamlit run app.py
```

3. Open your browser and navigate to the displayed URL (usually `http://localhost:8501`)

## Excel Files

The app will automatically detect and display all Excel files (`.xlsx` and `.xls`) in the `hanovaexcel` folder:

- **una sola factura.xlsx**
- **varias facturas.xlsx**
- **varios clientes.xlsx**

Each file will be displayed in its own tab with full data viewing capabilities.
