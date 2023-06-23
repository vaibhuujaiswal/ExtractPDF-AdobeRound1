import logging
import os.path
import os
import zipfile
import json
import csv

from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation


def dataRetriever():
    print("Function DataRetriever call recieved")
    testFolderName = "TestDataSet"
    directory = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))) + "/"+ testFolderName
    pdf_test_folder = os.listdir(directory)
    if '.DS_Store' in pdf_test_folder:
        pdf_test_folder.remove('.DS_Store')
    for pdf in pdf_test_folder:
        print(pdf)
        pdfDataExtractor(pdf[:-4]+".zip","/" + testFolderName + "/" + pdf)
    unzipDataFiles()
    directory=os.path.dirname(directory)+"/jsonDataSet"
    json_file=os.listdir(directory)
    if '.DS_Store' in json_file:
        json_file.remove('.DS_Store')
    data=[]
    for path_json in json_file:
        try:
            destinationDirectory = directory+ "/"+ path_json + "/"+ "structuredData.json"
            data.extend(extractJsonDataforOutput(destinationDirectory))
            print(path_json +" Extracted Successfully!")
        except:
            print("Extracting Json file error!")
            continue

    grandparent_directory = os.path.abspath(os.path.join(os.getcwd(), "..", ".."))
    final_csv_file_path = os.path.join(grandparent_directory, "ExtractedData.csv")
    with open(final_csv_file_path,"w",newline='') as file:
        writer = csv.writer(file)
        dataheader=["Bussiness__City","Bussiness__Country",	"Bussiness__Description",	"Bussiness__Name",	"Bussiness__StreetAddress",	"Bussiness__Zipcode",	"Customer__Address__line1",	"Customer__Address__line2",	"Customer__Email",	"Customer__Name",	"Customer__PhoneNumber",	"Invoice__BillDetails__Name",	"Invoice__BillDetails__Quantity",	"Invoice__BillDetails__Rate",	"Invoice__Description",	"Invoice__DueDate",	"Invoice__IssueDate",	"Invoice__Number",	"Invoice__Tax"]
        data.insert(0,dataheader)
        writer.writerows(data)

    print("Function ends here, go check out your output in the main file")

#function directly taken from sample code provided by API service adobe
def pdfDataExtractor(zipFile,dataSet):
    logging.basicConfig(level=os.environ.get("LOGLEVEL", "INFO"))
    try:
    # get base path.
        base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

        # Initial setup, create credentials instance.
        credentials = Credentials.service_account_credentials_builder() \
            .from_file(base_path + "/pdfservices-api-credentials.json") \
            .build()

        # Create an ExecutionContext using credentials and create a new operation instance.
        execution_context = ExecutionContext.create(credentials)
        extract_pdf_operation = ExtractPDFOperation.create_new()

        # Set operation input from a source file.
        source = FileRef.create_from_local_file(base_path + dataSet)
        extract_pdf_operation.set_input(source)

        # Build ExtractPDF options and set them into the operation
        extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder().with_element_to_extract(ExtractElementType.TEXT).build()
        extract_pdf_operation.set_options(extract_pdf_options)

        # Execute the operation.
        result: FileRef = extract_pdf_operation.execute(execution_context)

        # Save the result to the specified location : i.e
        result.save_as(base_path + "/zippedDataSet/"+zipFile)
    except (ServiceApiException, ServiceUsageException, SdkException):
        logging.exception("Exception encountered while executing operation")

#Using the zip folder, we will extract json data by unzipping files and then run it with our main program
def unzipDataFiles():
    '''
    UnzipDataFile() takes function from zipped folder file where we have stored the zippedDataset
    and stores the structured json file in each folder representing the pdf of the dataset
    '''
    directory = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))+"/zippedDataSet"
    zip_folder = os.listdir(directory)
    #remove  hidden ds value from our zip folder list
    if '.DS_Store' in zip_folder:
        zip_folder.remove('.DS_Store') 
    
    for zip in zip_folder:
        zip_path = directory+ "/" +zip
        destination_name = os.path.dirname(directory)+"/jsonDataSet/"+zip[0:-4]
        # print(destination_name)
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(destination_name)

def extractJsonDataforOutput(json_file_path):
    '''
   extractJsonDataforOutput() function inputs a path to the structured json file for each and everyinput and 
   extracts tables based on the requirement asked in the csv file
    
    '''
    with open(json_file_path, 'r') as file:
    # Load the JSON data
        data = json.load(file)
    #count total element points in the data
    elementPoints = len(data["elements"])
    IDTextList = {}
    elementID = 0
    for element in data["elements"]:
        elementID  = elementID + 1
        try:
            IDTextList[elementID]= element["Text"]
        except KeyError:
            continue
    
    #bussiness name extacted by iterating and fin and find text with size greater than 20.
    Bussiness__Name= ""
    business_name_id = 0 #will be used in other function to extract business related information likes its information
    print("business name extracted")
    for iterator in IDTextList:
        if (data["elements"][iterator-1]["TextSize"] > 20):
            Bussiness__Name += IDTextList[iterator].strip() 
            business_name_id = iterator
            break
    
    #extract business description
    print("business description extracted")
    Bussiness__Description = "".join(IDTextList[business_name_id+1].strip())

    #using the business name id, we can find the sut
    Complete_Customer__Details = ""
    Invoice__Description= ""
    Invoice__DueDate= ""
    customerInformationFlag = False
    inventoryCatalgoueb = 10
    for iterator in IDTextList:
        if iterator > (business_name_id+1):
            if data["elements"][iterator - 1]["Font"]["name"].endswith("BoldMT") and IDTextList[iterator].strip() == "ITEM":
                inventoryCatalgoueb = iterator
                break
            elif data["elements"][iterator - 1]["Bounds"][0] < 83 and data["elements"][iterator - 1]["Bounds"][0] > 79: #review address limits
                Complete_Customer__Details += IDTextList[iterator]
                # print(Complete_Customer__Details)
            elif data["elements"][iterator - 1]["Bounds"][2] > 500:
                Invoice__DueDate += IDTextList[iterator]
            else:
                Invoice__Description += IDTextList[iterator]

    #We have split the invoice description to iterate over the word and remover words like payement, or
    #the dollar sign indicating. We have created filtered description list which stores words that do not contain
    #any words like details, payment or dollar


    Invoice__Description = Invoice__Description.split()
    filtered_description = []
    for word in Invoice__Description:
        if word.strip() != "DETAILS" and "PAYMENT" not in word and "$" not in word:
            filtered_description.append(word)

    if filtered_description[0] == "PAYMENT":
        filtered_description = filtered_description[1:]

    if "$" in filtered_description[-1]:
        filtered_description = filtered_description[:-1]

    Invoice__Description = " ".join(filtered_description)
    Invoice__DueDate = Invoice__DueDate.split("Due date:")[-1].strip()

    #using the complete customer details we will extract information like customer phone number,
    #name, email and adress
    #Used a flag (phone_number_found) to ensure that only the first phone number matching the specified criteria is assigned to Customer__PhoneNumber.
    # Created a new list called updated_details to store the filtered words when iterating over Complete_Customer__Details.

    Customer__PhoneNumber = ""
    Customer__Name = ""
    Customer__Email = ""
    Customer__Address__line1 = ""
    Customer__Address__line2 = ""

    Complete_Customer__Details = Complete_Customer__Details.split()  # converts into a list of words
    Complete_Customer__Details.remove('BILL')
    Complete_Customer__Details.remove('TO')

    Customer__First__Name = Complete_Customer__Details.pop(0)
    Customer__Last__Name = Complete_Customer__Details.pop(0)
    Customer__Name = Customer__First__Name + " " + Customer__Last__Name

    phone_number_found = False
    updated_details = []

    for word in Complete_Customer__Details:
        if '-' in word and len(word) == 12 and len(word.split('-')) == 3 and not phone_number_found:
            Customer__PhoneNumber = word
            phone_number_found = True
        else:
            updated_details.append(word)

    Complete_Customer__Details = updated_details

    if '@' in Complete_Customer__Details[0]:
        if '.com' in Complete_Customer__Details[0]:
            Customer__Email = Complete_Customer__Details.pop(0)
        else:
            Customer__Email = Complete_Customer__Details.pop(0) + Complete_Customer__Details.pop(0)

    Customer__Address__line1 = " ".join(Complete_Customer__Details[:3])
    Customer__Address__line2 = " ".join(Complete_Customer__Details[3:])

    Complete_invoice__Details = ""
    Complete_bussiness_Details = ""
    Invoice__IssueDate = ""
    Invoice__Number = ""
    Bussiness__Zipcode = ""
    Bussiness__Country = ""
    Bussiness__City = ""
    Bussiness__StreetAddress = ""

    for iterator in IDTextList:
        if iterator < business_name_id:
            if data["elements"][iterator - 1]["Bounds"][2] > 540:
                Complete_invoice__Details += IDTextList[iterator]
            elif data["elements"][iterator - 1]["Bounds"][0] < 78:
                Complete_bussiness_Details += IDTextList[iterator]


    post_issue_invoice_details = Complete_invoice__Details.split("Issue date")
    Invoice__IssueDate = post_issue_invoice_details[1].strip()
    Invoice__Number = post_issue_invoice_details[0].split("Invoice#")[-1].strip()

    # Extract Business details
    business_details = Complete_bussiness_Details.split()

    Bussiness__Zipcode = business_details[-1]
    # print(Bussiness__Zipcode)
    Bussiness__Country = business_details[-3] + " " + business_details[-2]
    # print(Bussiness__Country)
    Bussiness__City = business_details[-4].split(",")[0]
    # print(Bussiness__City)

    #Street Address is extracted
    Bussiness__StreetAddress = " ".join(business_details[len(IDTextList[business_name_id].split()):-4])
    Bussiness__StreetAddress = Bussiness__StreetAddress.split(",")[:-1]
    if len(Bussiness__StreetAddress) > 0:
        Bussiness__StreetAddress = Bussiness__StreetAddress[0]
    else:
        Bussiness__StreetAddress = ""

    keys=list(IDTextList.keys())
    Invoice__Tax = 10
    invoice_tax_id = ""
    for index in range(len(keys)):
        if "Total Due" in IDTextList[keys[index]].strip():
            if index > 0:
                previous_text = IDTextList[keys[index-1]].strip()
                try:
                    Invoice__Tax = int(previous_text)
                    invoice_tax_id = keys[index-1]
                except:
                    try:
                        Invoice__Tax = int(previous_text.split()[-1])
                        invoice_tax_id = keys[index-1]
                    except:
                        pass #keep original tax value
            break
    
    Invoice__BillDetails__Name=[]
    Invoice__BillDetails__Quantity=[]
    Invoice__BillDetails__Rate=[]

    ##Extract invoice bill details

    first = None
    last = None
    Invoice__BillDetails__Name = []
    Invoice__BillDetails__Quantity = []
    Invoice__BillDetails__Rate = []

    for index in range(len(list(IDTextList))):
        if list(IDTextList)[index] == inventoryCatalgoueb:
            first = list(IDTextList)[index + 4]
        elif list(IDTextList)[index] == invoice_tax_id:
            last = list(IDTextList)[index - 7]
            break

    if first is not None and last is not None:
        for iterator in range(first, last + 1, 8):
            Invoice__BillDetails__Name.append(IDTextList[iterator].strip())
            Invoice__BillDetails__Quantity.append(int(IDTextList[iterator + 2]))
            Invoice__BillDetails__Rate.append(int(IDTextList[iterator + 4]))

    #Return the list of data
    listForCSVOutput=[]
    for rows in range(len(Invoice__BillDetails__Name)):
        listForCSVOutput.append([Bussiness__City,Bussiness__Country,Bussiness__Description,Bussiness__Name,Bussiness__StreetAddress,Bussiness__Zipcode,Customer__Address__line1,Customer__Address__line2,Customer__Email,Customer__Name,Customer__PhoneNumber,Invoice__BillDetails__Name[rows],Invoice__BillDetails__Quantity[rows],Invoice__BillDetails__Rate[rows],Invoice__Description,Invoice__DueDate,Invoice__IssueDate,Invoice__Number,Invoice__Tax])
    return listForCSVOutput 


##call our main function to run, can add extra step to include a new folder
dataRetriever()