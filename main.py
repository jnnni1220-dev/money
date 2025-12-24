import pandas as pd
import msoffcrypto

def read_excel_file(file_path, password):
    try:
        # Decrypt the file
        with open(file_path, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            # Decrypt directly to the output file
            decrypted_path = file_path.replace('.xlsx', '_decrypted.xlsx')
            with open(decrypted_path, 'wb') as out_file:
                office_file.decrypt(out_file)
        
        print(f"Decrypted file saved as: {decrypted_path}")
        
        # Now read the decrypted file for display (optional)
        df = pd.read_excel(decrypted_path, engine='openpyxl')
        print("Excel file read successfully!")
        print(df.head())  # Print the first few rows
        return df
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    file_path = "template.xlsx"
    password = "togemoney"
    read_excel_file(file_path, password)