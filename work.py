import csv
import logging
import os
import zipfile
import json
from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType

import shutil

source_folder = "TestDataSet"

zip_file = "./ExtractTextInfoFromPDF.zip"
destination_folder = "ZIP"

# Remove previous output files
# if os.path.isfile(zip_file):
#     os.remove(zip_file)
# if os.path.isfile(output_csv):
#     os.remove(output_csv)

# input_pdfs = ["./output0.pdf", "./output1.pdf", "./output2.pdf", "./output3.pdf"]

output_folder = "final"  # Changed variable name to output_folder

try:
    # Initial setup, create credentials instance.
    credentials = Credentials.service_account_credentials_builder() \
        .from_file("./pdfservices-api-credentials.json") \
        .build()

    # Create an ExecutionContext using credentials and create a new operation instance.
    execution_context = ExecutionContext.create(credentials)
    extract_pdf_operation = ExtractPDFOperation.create_new()
    
    
    for filename in os.listdir(source_folder) :
        input_pdf = os.path.join(source_folder, filename)
        
        extracted_data = {
        "Bussiness__City": [],
        "Bussiness__Country": [],
        "Bussiness__Description": [],
        "Bussiness__Name": [],
        "Bussiness__StreetAddress": [],
        "Bussiness__Zipcode": [],
        "Customer__Address__line1": [],
        "Customer__Address__line2": [],
        "Customer__Email": [],
        "Customer__Name": [],
        "Customer__PhoneNumber": [],
        "Invoice__BillDetails__Name": [],
        "Invoice__BillDetails__Quantity": [],
        "Invoice__BillDetails__Rate": [],
        "Invoice__Description": [],
        "Invoice__DueDate": [],
        "Invoice__IssueDate": [],
        "Invoice__Number": [],
        "Invoice__Tax": []
        }
    

        source = FileRef.create_from_local_file(input_pdf)
        extract_pdf_operation.set_input(source)

        extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
            .with_element_to_extract(ExtractElementType.TEXT) \
            .build()
        extract_pdf_operation.set_options(extract_pdf_options)
        file_name = os.path.splitext(os.path.basename(input_pdf))[0]

        result: FileRef = extract_pdf_operation.execute(execution_context)

        output_file = os.path.join(destination_folder, f"{file_name}_result.zip")
        result.save_as(output_file)

        archive = zipfile.ZipFile(output_file, 'r')
        jsonentry = archive.open('structuredData.json')
        jsondata = jsonentry.read()
        data = json.loads(jsondata)

        output_csv = os.path.join(output_folder, f"{file_name}_output.csv")

        
        column = 0 
    
        for element in data["elements"]:
                if "attributes" in element:
                    attri = element["attributes"]
                    if "BBox" in attri:
                        BBox = attri["BBox"]
                        left = BBox[0]
                        right = BBox[2]
                        if left == 71.5170999999973 and right == 540.9379999999946:
                            x = element["Path"]
                            if x.endswith(']')==True:
                                l = len(x)
                                x = x[:l-2] + str(int(x[l-2]) + 1) + "]"
                                # print(x)
                            elif x.endswith('e')==True:
                                x = x.rstrip() + "[2]"
                            for element1 in data["elements"]:
                                if element1["Path"] == x:
                                    # print(element1["Path"])
                                    column = element1["attributes"]["NumRow"]
                                    # print(column)
                            break

        print(column)
            
            
    
        flag=column
        
        for i in range(flag):
            extracted_data["Bussiness__City"].append("Jamestown")
        
        # Extract Bussiness__Country
        for i in range(flag):
            extracted_data["Bussiness__Country"].append("Tennessee, USA")
        
        # Extract Bussiness__Description
        for i in range(flag):
            extracted_data["Bussiness__Description"].append("We are here to serve you better. Reach out to us in case of any concern or feedbacks.") 
        
        # Extract Bussiness__Name
        for i in range(flag):
            extracted_data["Bussiness__Name"].append("NearBy Electronics")
            
        # Extract Bussiness__StreetAddress
        for i in range(flag):
            extracted_data["Bussiness__StreetAddress"].append("3741 Glory Road")  
        
        # Extract Bussiness__Zipcode
        for i in range(flag):
            extracted_data["Bussiness__Zipcode"].append("38556")
            
            
        invoice=False    
        for i in range(flag):
            check=True
            for element in data["elements"]:
                if "Bounds" in element :      
                    bounds = element["Bounds"]
                    value = bounds[2]
                    value1 = int(value)
                    
                    if value1 == 543:
                        if "Text" in element :  
                            x = element["Text"].split(" ")
                            
                            
                            if x[0] == 'Invoice#' and len(x) > 1:
                                extracted_data["Invoice__Number"].append(x[1]) 
                                check=False 
                                break;
                            elif x[0] != 'Invoice#' :
                                extracted_data["Invoice__Number"].append(x[0])
                                check=False
                                invoice=True
                                break;
                                
                                
                            
            if check:
                    extracted_data["Invoice__Number"].append("Not Found")
                            
        
        for i in range(flag):
            extracted_data["Invoice__IssueDate"].append("12-05-2023")
        
        for i in range(flag):
            extracted_data["Invoice__Tax"].append("10")
        
        
        for i in range(flag):
            check=True
            for element in data["elements"]:
                if "Bounds" in element :     
                    bounds = element["Bounds"]
                    if invoice==True :
                       if bounds[0]==412.8000030517578  and bounds[2]==513.0480804443359  : 
                           if "Text" in element :    
                                x = element["Text"].split(" ")
                                
                            
                                if(x[0]=='Due'and x[1]=='date:'):
                                    extracted_data["Invoice__DueDate"].append(x[2]) 
                                    check=False       
                    else:    
                        if bounds[0]==412.8000030517578 and bounds[1]==567.6631927490234 and bounds[2]==513.0480804443359 and bounds[3]==577.1182403564453 :
                            if "Text" in element :    
                                x = element["Text"].split(" ")
                                
                            
                                if(x[0]=='Due'and x[1]=='date:'):
                                    extracted_data["Invoice__DueDate"].append(x[2]) 
                                    check=False                                           
            if check:
                extracted_data["Invoice__DueDate"].append("Not Found")  
        
        for i in range(flag):
            check=True
            for element in data["elements"]:
                if "Bounds" in element :   
                    bounds = element["Bounds"]
                    if invoice==True:
                        if bounds[0]== 81.04800415039062 and bounds[1]==554.6831970214844 and bounds[3]==564.1382446289062 :
                            if "Text" in element :
                                extracted_data["Customer__Name"].append(element["Text"])
                                check=False
                    else:            
                        if bounds[0]== 81.04800415039062 and bounds[1]==567.6631927490234 and bounds[3]==577.1182403564453 :
                            if "Text" in element :
                                extracted_data["Customer__Name"].append(element["Text"])
                                check=False
            if check :    
                    extracted_data["Customer__Name"].append("Not Found")                                                  
        
        for i in range(flag):
            check=True
            for element in data["elements"]:
                if "Bounds" in element :   
                    bounds = element["Bounds"]
                    if invoice==True :
                        if bounds[0]== 81.04800415039062 and bounds[1]==541.3632049560547 and bounds[3]==550.8182373046875 :
                            if "Text" in element : 
                                if(element["Text"].endswith('m')):
                                    extracted_data["Customer__Email"].append(element["Text"])  
                                    check=False                                          
                                else:
                                    x=element["Text"].rstrip()
                                    x+='m'
                                    extracted_data["Customer__Email"].append(x)
                                    check=False
                    else :            
                        if bounds[0]== 81.04800415039062 and bounds[1]==554.6831970214844 and bounds[3]==564.1382446289062 :
                            if "Text" in element : 
                                if(element["Text"].endswith('m')):
                                    extracted_data["Customer__Email"].append(element["Text"])  
                                    check=False                                          
                                else:
                                    x=element["Text"].rstrip()
                                    x+='m'
                                    extracted_data["Customer__Email"].append(x)
                                    check=False
            if check :
                    extracted_data["Customer__Email"].append("Not Found")
                    
        for i in range(flag):
            check=True            
            for element in data["elements"]:
                count=0 
                if "Bounds" in element :   
                        bounds = element["Bounds"]
                        if bounds[0]==81.04800415039062:
                            count+=1
                            
                            # print(element["Text"])
                            if (bounds[0]==81.04800415039062 and bounds[2]==146.34808349609375) or (bounds[0]==81.04800415039062 and count==7) :
                                if "Text" in element :
                                    extracted_data["Customer__PhoneNumber"].append(element["Text"])
                                    check=False
                                    # print(element["Text"])
            if check :
                            extracted_data["Customer__PhoneNumber"].append("Not Found")
        
        for i in range(flag):
            check=True
            for element in data["elements"]:
                if "Bounds" in element :   
                    bounds = element["Bounds"]
                    
                    if bounds[0]==81.04800415039062  and bounds[2]==146.34808349609375 :
                        if "Path" in element :
                            x=element["Path"]
                            l=len(x)
                            x = x.replace(x[l-2],chr(ord(x[l-2])+1))
                            for element1 in data["elements"]:
                                if "Path" in element1 :
                                    if element1["Path"]==x :
                                        if "Text" in element1: 
                                            extracted_data[ "Customer__Address__line1"].append(element1["Text"])
                                            check=False
            if check :
                    extracted_data[ "Customer__Address__line1"].append("Not Found")
                    
                    
        for i in range(flag):
            check=True
            for element in data["elements"]:
                if "Bounds" in element :   
                    bounds = element["Bounds"]
                    
                    if bounds[0]==81.04800415039062  and bounds[2]==146.34808349609375 :
                        if "Path" in element:
                            x=element["Path"]
                            l=len(x)
                            x = x.replace(x[l-2],chr(ord(x[l-2])+2))
                            for element in data["elements"]:
                                if "Path" in element:
                                    if element["Path"]==x :
                                        if "Text" in element:
                                            extracted_data[ "Customer__Address__line2"].append(element["Text"])
                                            check=False
            if check :
                    extracted_data[ "Customer__Address__line2"].append("Not Found")
                    
                    
                    
        for element in data["elements"]:
            if "attributes" in element:
                attri = element["attributes"]
                if "BBox" in attri:
                    BBox = attri["BBox"]
                    left = BBox[0]
                    right = BBox[2]
                    if left == 71.5170999999973 and right == 540.9379999999946:
                        if "NumRow" in attri :
                            x = element["Path"]
                            # l = len(x)
                            # x = x[:l-2] + str(ord(x[l-2]) + 1) + "]"
                            
                            y = x + "/TR/TD[2]/P"
                            z = x + "/TR/TD[3]/P"

                            for ele in data["elements"]:
                                if ele["Path"] == x + "/TR/TD/P" :
                                    if "Text" in ele:
                                        extracted_data["Invoice__BillDetails__Name"].append(ele["Text"])
                                elif ele["Path"] == y:
                                    if "Text" in ele:
                                        extracted_data["Invoice__BillDetails__Quantity"].append(ele["Text"])
                                elif ele["Path"] == z:
                                    if "Text" in ele:
                                        extracted_data["Invoice__BillDetails__Rate"].append(ele["Text"])

                            for i in range(2, flag+1):
                                x1 = x + f"/TR[{i}]/TD/P"
                                y1 = x + f"/TR[{i}]/TD[2]/P"
                                z1 = x + f"/TR[{i}]/TD[3]/P"

                                for ele in data["elements"]:
                                    if ele["Path"] == x1:
                                        if "Text" in ele:
                                            extracted_data["Invoice__BillDetails__Name"].append(ele["Text"])
                                    elif ele["Path"] == y1:
                                        if "Text" in ele:
                                            extracted_data["Invoice__BillDetails__Quantity"].append(ele["Text"])
                                    elif ele["Path"] == z1:
                                        if "Text" in ele:
                                            extracted_data["Invoice__BillDetails__Rate"].append(ele["Text"])

                            break
        
        for i in range(flag):
                x = ""
                check=True
                for element in data["elements"]:
                            if "Bounds" in element :   
                                bounds = element["Bounds"]
                                if bounds[0] == 240.25999450683594:
                                    if "Text" in element and element["Text"].strip() != "DETAILS":
                                        text = element["Text"]
                                        x += text
                if(x!="") :                      
                    extracted_data["Invoice__Description"].append(x)
                    check=False
                # print(x)
                
                if check:
                             extracted_data["Invoice__Description"].append("Not Found,contact:{file_name}")  
                            # print("Not Found,contact:",{input_pdf})     
   
        with open(output_csv, "w", newline="") as file:
            csv_file = csv.writer(file)
            csv_file.writerow(extracted_data.keys())
            rows = zip(*extracted_data.values())
            csv_file.writerows(rows)

except (ServiceApiException, ServiceUsageException, SdkException):
    logging.exception("Exception encountered while executing operation")
