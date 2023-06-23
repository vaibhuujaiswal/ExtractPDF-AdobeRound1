
# Adobe Document Cloud Hackathon - Round 1 submission

The repository contains source code and final csv output for test dataset for the first round of adobe document cloud hackathon submission by Vaibhav Jaiswal.



## Repository content

1) *src* : src file contains the extractPDF folder
2) *extractPDF* : folder contains the run.py file that will extract information from pdf and store it in a csv file.
3) *TestDataSet*: folder contains the TestDataSet pdf files (0 -99) that have to be processed using the extractPDF Adobe API and outputed in form of a csv file.
4) *zippedDataSet* : contains the zipped json output files that have been outputed by the Adobe extractPDF service 
5) *jsonDataSet* : Json data set has folders correlated to each pdf with the "structured.json" file inside of it.
6) *InvoicesData* : contains sample pdf file and respective output (given by adobe)
7) *ExtractedData.csv* : Contains the extracted compiled data from 100 pdfs given the TestDataSet
8)*pdfservices-api-credentials.json* : Contains pdfservices credentials. Must always be in the folder outside of src.
9) *private.key* : Contains "My - Vaibhav's" private key to access *pdfservices*
10) requirements.txt : fulfills dependency
11) Note : Language used - python


## How to Run the program?
1) Store all PDFs whos content you want to extract and store in csv inside the TestDataSet. Please do not store them inside a folder, but rather as pdf inside TestDataSet. 
2) Make sure to keep *pdfservices-api-credentials.json* and *private.key* outside the src folder. (Basically as the grandparent directory of run.py file)
3) Change directory to run.py (present inside src/extractPDF). Fulfill all dependency (view documentation of adobe Extract API).
4) use command ```python3 ./run.py``` to run the python file.


## Flow of the program

The program will call retrivedata() function that will loop through all pdfs in TestDataSet file and pass it through the extract-pdf-api services. It will store the zipped json file in the zippedDataSet folder. We will then unzip the files and store them in jsonDataSet.
We will then use our extractJsonDataforOutput() function to extract data based on properties of each "element" in json file like font, size, text etc. We will store each variable corresponding to what was asked in the csv file.
We will append all data in a list and finally store it as a csv file in the main directory.

## Result
ExtractedData.csv contains the result for the given testset.


## Extract PDF API (Adobe)

https://developer.adobe.com/document-services/docs/overview/pdf-extract-api/

## Contact

Contact me at vaibhav20547@iiitd.ac.in for any doubts!


