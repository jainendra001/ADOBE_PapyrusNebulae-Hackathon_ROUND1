# Adobe_PDF_Extract_API_Project
## This project is made as a submission to ongoing Adobe PapyrusNebulae Hackathon 2023 .
*This repository contain files, there relevance in the project* 
* Adobe_papyrus.py is the main python file that contain the core code for conversion of PDF's to JSON file using the Adobe PDF Extract API and also contains the code for conversion these PDF's to a single CSV file.
* final_add_jupiter.ipynb is the jupiter file that contains the code for concatenation of multiple CSV files into one.
* merged_file.csv contains all the details of the invoices as PDF file.
* **NOTE :**In my approach I have appended certain details directly as ADOBE is a Product Based Company so it have to make products specific to certain Company so the Primary Details can be fixed.
* pdfservices-api-credentials.json it contains the credential of mine Adobe PDF Extract API.
* private.key is also with the Adobe PDF Extract API details.
* work.py contains the code that convert multiple PDFs to multiple CSV file.
* *ZIP* this folder contains all the Zipped files.
* *TestDataSet* this folder contains all the Test PDF given by ADOBE for the Hackathon.
* *final* this folder contains all the CSV files for the individual PDFs with relavant name to access them easily.
* *NOTE :* In my approach I have made multiple CSV files so that the company can easily access single CSV file when they needed. 
### My approach for this TASK :
1.The code begins by importing the required libraries and modules:\
    <sub>
    
    import os
    import zipfile
    import json
    import csv
    import logging
    from adobe.pdfservices.operation.auth.credentials import Credentials
    from adobe.pdfservices.operation.execution_context import ExecutionContext
    from adobe.pdfservices.operation.pdf_operation_factory import PDFOperationFactory
    from adobe.pdfservices.operation.pdf_operations import ExtractPDFOperation,ExtractPDFOption

    
2.The *conversion_from_pdf_to_multi_csv()* function is defined:\
    <sub>
    
    def conversion_from_pdf_to_multi_csv():
    
   
    
3.The *source_folder* variable is set to the directory path where the PDF files are located.\
4.The *zip_file* variable specifies the name and path of the ZIP file that will be created to store the extracted data.\
5.The *destination_folder* variable is set to the directory path where the ZIP file will be saved. \
6.The *output_folder* variable specifies the directory path where the final CSV files will be stored. 

7.The *ExtractPDFOperation* class is used to create a new operation instance for extracting text from a PDF. 
     <sub>
     
     extract_pdf_operation = PDFOperationFactory.create_extract_pdf_operation()

 


8.The code iterates over each file in the *source_folder*:
    <sub>
    
        
    for file_name in os.listdir(source_folder):
        if file_name.endswith(".pdf"):
           (Process each Pdf)



9.The current PDF file is set as the input for the extraction operation:
    <sub>
    
    
    input_pdf_path = os.path.join(source_folder, file_name)
    extract_pdf_operation.set_input_file(input_pdf_path)


10.The *ExtractPDFOptions* class is used to configure the extraction options. In this case, only text extraction is enabled:
    <sub>
    
    extract_options = ExtractPDFOptions.Builder().add_elements_to_extract(ExtractElementType.TEXT).build()



11.The extraction options are set for the operation:
    <sub>

    extract_pdf_operation.set_options(extract_options)


12.The result of the extraction operation is saved as a ZIP file:\
    <sub>
    
    result = extract_pdf_operation.execute(execution_context)
    result.save_as(zip_file)


13.The ZIP file is opened and the "structuredData.json" file containing the extracted data is read:
    <sub>
        
    with zipfile.ZipFile(zip_file, "r") as zip_ref:
        zip_ref.extract("structuredData.json", destination_folder)

    
14.The JSON data is loaded into a Python dictionary:
    <sub>
        
    with open(os.path.join(destination_folder, "structuredData.json"), "r") as json_file:
        json_data = json.load(json_file)


    
15.The code defines an extracted_data dictionary that will store the extracted data for each field:
    <sub>
    
     extracted_data = {
                "Bussiness__City": [],
                "Bussiness__Country": [],
                "Bussiness__Description": [],
                "Bussiness__Name": [],
                "Bussiness__StreetAddress": [],
                "Bussiness__Zipcode": [],
                "Customer_Address_line1": [],
                "Customer_Address_line2": [],
                "Customer__Email": [],
                "Customer__Name": [],
                "Customer__PhoneNumber": [],
                "Invoice_BillDetails_Name": [],
                "Invoice_BillDetails_Quantity": [],
                "Invoice_BillDetails_Rate": [],
                "Invoice__Description": [],
                "Invoice__DueDate": [],
                "Invoice__IssueDate": [],
                "Invoice__Number": [],
                "Invoice__Tax": []
                }
                
                
16.The code searches for specific elements in the extracted data that correspond to different fields and extracts their values. \This part of the code will depend on the structure of the JSON data and the specific fields we want to extract.\

17.The extracted values are appended to the corresponding lists in the *extracted_data* dictionary.\
18.Finally, the extracted data is written to a CSV file in the *output_folder* using the *csv.writer* module. Each field's data will be written in a separate CSV file:
    <sub>
    
    for field, data in extracted_data.items():
        csv_file_path = os.path.join(output_folder, f"{field}.csv")
        with open(csv_file_path, "w", newline="") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow([field])
            writer.writerows(data)

            
19.The code handles exceptions that may occur during the execution of the operation and logs any encountered exceptions:
    <sub>
    
    except Exception as e:
        logging.error(f"An error occurred while processing file {file_name}: {str(e)}")

        
        
20.The *conversion_from_multi_csv_to_one()* function is defined, this function reads multiple CSV files from a specified directory and concatenates them into a 
    single CSV file.\
21.*To use this function:*
  * Set the file_path variable to the directory path where the CSV files are located.
  * List all the files in the file_path directory.
  * Sort the file list based on the lexiographically as per name.
  * Create an empty DataFrame to store the concatenated data.
  * Iterate over each file in the sorted file list:
    If the file is a CSV file, read it using pd.read_csv().
    Concatenate the data from the current file with the existing data in df_concat using 
    pd.concat().
  * The concatenated DataFrame is saved to a new CSV file named "merged_file.csv" in the 
    current working directory.
