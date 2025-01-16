import pandas as pd

def read_excel_links(file_path):
    """Reads a list of URLs from an Excel file."""
    df = pd.read_excel(file_path)
    return df['Links'].tolist()  # Assuming your Excel has a column named 'Links'
