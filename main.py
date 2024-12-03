import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QVBoxLayout, QWidget
import os
import pandas as pd
import json

class DocumentClassifier(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Document Classifier")
        self.setGeometry(100, 100, 400, 200)

        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Create buttonhttps://github.com/Sagar-Agicha/Tracker_MOM.git
        self.upload_button = QPushButton("Upload Document", self)
        self.upload_button.clicked.connect(self.upload_document)
        layout.addWidget(self.upload_button)

        # Create label for displaying results
        self.result_label = QLabel("No document uploaded yet", self)
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label)

    def upload_document(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Document", "", "All Files (*.*)")
        
        if file_path:
            _, file_extension = os.path.splitext(file_path)
            file_extension = file_extension[1:].lower()  # Remove the dot and convert to lowercase
            
            # Handle Excel files
            if file_extension in ['xlsx', 'xls', 'csv']:
                # Read Excel file
                df = pd.read_excel(file_path)
                
                # Check if any row contains any of the required headers
                required_headers = ['ID', 'Discussion', 'ETA', 'Status', 'Owner', 'Remarks', 'Name', 'City']
                header_row = -1
                row_index = 0
                
                while row_index < len(df):
                    # Check current row
                    if row_index == 0:
                        current_row = df.columns.str.lower()
                        found_headers = [header for header in required_headers if header.lower() in current_row]
                        if found_headers:
                            # Print column mapping for first row
                            for col_num, col_name in enumerate(df.columns):
                                print(f"Column {col_num + 1} -> {col_name}")
                    else:
                        current_row = df.iloc[row_index-1].astype(str).str.lower()
                        found_headers = [header for header in required_headers if header.lower() in current_row.values]
                        if found_headers:
                            # Print column mapping for other rows
                            row_values = df.iloc[row_index-1]
                            for col_num, col_value in enumerate(row_values):
                                print(f"Column {col_num + 1} -> {col_value}")
                    
                    if found_headers:  # If any headers were found
                        header_row = row_index
                        print(f"\nHeaders found in row {header_row+1}!")
                        required_headers = [header for header in required_headers if header not in found_headers]
                        print("Remaining required headers: ", required_headers)
                        print("Remaining required headers count: ", len(required_headers))
                        
                        if len(required_headers) < 3:
                            # After finding headers, set up dataframe and print all owner-eta mappings
                            if header_row > 0:
                                df.columns = df.iloc[header_row-1]
                                df = df.iloc[header_row:].reset_index(drop=True)
                            
                            # Find owner and eta columns
                            owner_col = None
                            eta_col = None
                            for col in df.columns:
                                col_lower = str(col).lower()
                                if 'owner' in col_lower:
                                    owner_col = col
                                elif 'eta' in col_lower:
                                    eta_col = col
                            
                            # Print all owner-eta mappings if both columns found
                            if owner_col is not None and eta_col is not None:
                                print("\nOwner --> ETA mappings --> Task:")
                                for idx in range(len(df)):
                                    owner = df.loc[idx, owner_col]
                                    eta = df.loc[idx, eta_col]
                                    task = df.loc[idx, 'DISCUSSION']
                                    print(f"{owner} --> {eta} --> {task}")
                                    file_json = open('datas/Data.json', 'r')
                                    data = json.load(file_json)
                                    for i in range(len(data)):
                                        if data[i]['nickname'] == owner:
                                            task_num = data[i]['content'][0]['task_cnt'] - 1
                                            file_json_write = open('datas/Data.json', 'w')
                                            data[i]['content'].append({'Task' + str(task_num+1): [{'discussion': task, 'eta': eta}]})
                                            data[i]['content'][0]['task_cnt'] += 1
                                            json.dump(data, file_json_write, indent=4)
                                            file_json_write.close()
                                            file_json.close()
                                            break
                                break
                        
                    row_index += 1

            # Continue with existing file copying logic
            upload_dir = os.path.join('uploads', f'documents_{file_extension}')
            os.makedirs(upload_dir, exist_ok=True)
            
            # Copy file to the appropriate directory
            file_name = os.path.basename(file_path)
            destination = os.path.join(upload_dir, file_name)
            
            try:
                # Ensure the uploads directory exists
                if not os.path.exists('uploads'):
                    os.makedirs('uploads')
                
                # Copy the file
                import shutil
                shutil.copy2(file_path, destination)
                
                # Update the label with file information
                self.result_label.setText(f"Document uploaded successfully!\nFile extension: {file_extension}\nSaved to: {upload_dir}")
            except Exception as e:
                self.result_label.setText(f"Error saving file: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = DocumentClassifier()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
