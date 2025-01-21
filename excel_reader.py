import pandas as pd
import os

def read_excel_links(file_path):
    """Reads a list of URLs from an Excel file."""
    df = pd.read_excel(file_path)
    return df['Links'].tolist()  # Assuming your Excel has a column named 'Links'

def store_incorrect_bleeds(folder_path, file_name, incorrect_bleeds):
    try:
        # Ensure the folder exists
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        file_path = os.path.join(folder_path, file_name)

        # Create a DataFrame from the incorrect bleeds
        df = pd.DataFrame(incorrect_bleeds)

        if os.path.exists(file_path):
            # Append to the existing file
            existing_data = pd.read_csv(file_path)
            updated_data = pd.concat([existing_data, df], ignore_index=True)
            updated_data.to_csv(file_path, index=False)
        else:
            # Create a new file and save the data
            df.to_csv(file_path, index=False)

    except Exception as e:
        print(f"ERROR: Could not write to file '{file_name}': {e}")