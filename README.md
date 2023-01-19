# ExcelVBA
Projects using Excel and  VBA macros to automate repetitive tasks such as data entry, data manipulation, and report generation, which can save time and reduce the risk of manual errors.

Some examples are:

- **SplitReport():** The macro creates a new sheet for each distinct value in column A of sheet1. Then it will copy the entire information of every distinct value to the sheet with the same name as the distinct value.

  *A use case:* GA reports with hundreds of rows where information must be reviewed and validated by distinct stakeholders. It is very useful to have a tool to expedite the process.

- CreateAndEmailPDF(): This macro first looks at the values (stakeholders’ names) in Sheet1 column A, finds the sheet with the exact value, creates a PDF file, and saves it to the exact location the Excel file is saved in. Then it emails the PDF files created to recipients listed (stakeholders’ emails) in Sheet1's column B. It uses the same assumption that there is a sheet in the workbook with the same name as the value in column A of Sheet1.

  *A use case:* Email customized reports to different stake holders, like expenditures to their incumbents.

- Compare(): This macro will compare the contents of each cell in the sheet "Sheet1" with the corresponding cell in the sheet "Sheet1" of the workbook "Previous" which must be in the same directory. If the contents of a cell in "Compare" do not match the contents of the corresponding cell in "Previous", the cell will be highlighted in yellow in the "Compare" file.

  *A use case:* Validate the information bulk loaded into a system is exact to the source file. As well as when information goes through interfaces from one system to another there is need to audit for accuracy.
  
  
  
  
