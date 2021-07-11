# Bulk Insertion Of Excel Data To MS Access Database
##### _A process automation for bulk insertion of data from excel to MS Access using Blueprism v.6.10_

This process automation has been developed using Bluerism RPA tool with following Business Objects
- Utility - Environment
- Utility - General
- Data - OLEDB
- Utility - File Management
- MS Access VBO (Custom Object)

#### Processing Steps :

- The input data is in form of an excel file having approximately 30 columns and 2000 rows whereas the output file is MS Access database file with the same table structure.
- BOT first validates if both the files are present or not. If ot, then it will throw a business exception.
- Upon successful validation, BOT will read the excel data using OLEDB driver.
- Once, the data has been read BOT will first delete the existing records and then will insert all the records read in the prior step.
- BOT finishes the overall execution in matter of seconds!

#### DLL's Used :

- Microsoft.Office.Interop.Access.dll
- Microsoft.Office.Interop.Access.Dao.dll

---
**NOTE**

__Ensure that the DLL listed in the DLL's folder are present in the Blueprism installation folder.__

---
#### Custom Objects Created :

#### 1) MS Access VBO
__The runmode of this business object is "background"__

#### 1.1) Delete All Records
__This action has been created to delete all the records from a given table in the supplied MS Access database file.__
|Parameter|Direction|Data Type|Description|
|--- |--- |--- |--- |
|Database File Path|In|Text|The MS access database file path from where all of the table data needs to be deleted.|
|Table Name|In|Text|The table name in the MS Access database file where deletion operation will take place.|
|Message|Out|Text|The execution message for the action.|
|Success|Out|Flag|The flag value indication the status of the action.|


#### 1.2) Insert All Records
__This action has been created to insert all the records of a given datatable in the supplied table of the MS Access database file.__

|Parameter|Direction|Data Type|Description|
|--- |--- |--- |--- |
|Database File Path|In|Text|The MS access database file path records need to be inserted.|
|Table Name|In|Text|The table name in the MS Access database file where insertion operation will take place.|
|Input Data|In|Collection|The datatable which needs to be inserted.|
|Message|Out|Text|The execution message for the action.|
|Success|Out|Flag|The flag value indication the status of the action.|
