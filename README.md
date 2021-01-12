# UiPath-Snippets
Snippet Repository (One-Stop Hack Portal)

## Convert_DataTable_X_HTML

> Convert Datatable Object to HTML Table to be used for Sending of Emails.

Important Variables
1. in_dictionaryTableFlag (Determine if key pair value or row X column format)

Custom Variables for Normal Datatable Conversion
1. fontType (Font Family e.g. Arial)
2. tableHeaderBgColor
3. tableHeaderFontColor
4. colWidth (Width of Each Column)
5. tableWidth (Set to 100 if you want do not have dyanmic scaling of column)

```
When in_dictionaryTableFlag = True, means to ignore the header column and assume the headers are the first column.
When in_dictionaryTableFlag = False. means to assume that the datatable columns are the html table columns.
```

## ToggleSheetProtection

> Toggle the protection status of excel sheets

Important Variables
1. in_pwStr (Password to Protect File)
2. in_ExcelPath (Path to Excel File)
3. in_Sheet (Array of Sheet Names to Toggle Protection)

*Can be improved to include password protection properties such as "Content Protection".*

## Extract_HTML_Table

> Extract HTML Tables from .MSG file

Important Variables
1. in_MSGPath (Path to .MSG File)
2. isCustomHeader (True means that the first row of the file can be considered as the datatable header column)

*Can be improved to take emails from outlook directly or HTML Text*

## SaveMailFormat

> Save Outlook Mail as a .msg file extension.

Important Variables
1. in_Path (Full Path to Store .MSG fle)
2. in_Root_Folder (First-Level Folder normally it is "Inbox")
3. in_Child_Folder (Second-Level Folder)
4. in_Robot_Email (Email of Account to Access)

*Can be improved to take in series of email and path names, can be changed to multi-level folder structure*

## InvokeReplywithOFT

> Reply ontop of a .msg file

Important Variables
1. in_MsgFilePath (Path to MSG File to serve as the source to reply)
2. in_Attachments (Attachment to put in body)
3. in_ReplyBody (Additional Text to reply)
4. in_ToList (Additional Recipents)
5. in_CcList (Additional Recipents in the loop)
6. in_Subject (Empty, Text if want to overwrite current subject)

## ExcelNumberToLetter

> Convert a datatable index to Alphabet for Writing Range purposes.

Important Variables
1. colNum (Index f Datatable Column)

## Initialize_Outlook

> Simple Workflow to Open Outlook if closed.

Important Variables
1. vOutlookPath (Path to Outlook.exe)

* Type "where /r %systemdrive% outlook.exe" (without the double quotes) in cmd to get full path to outlook

## Index_Finder

> Search For a Dynamic Column in a Datatable by using keywords. (Works well when Column Name and Index not fixed)

Important Variables
1. in_KeywordList (Array of Keywords to used)
2. In_PureDataTable (Datatable of Column Headers to search for)

* Currently does not accept regex definers such as "()"

## Check_Sharepoint

> Checks whether a successful connection can be made to SharePoint

Important Variables
1. sharepointPath (Path to any file in SharePoint)

## Copy Directory and Delete All Files

> Copy all files from Source to Destination whilst deleting Source Folder (Optional)

Important Variables
1. in_SourcePath (Path to Copy File in Folder From)
2. in_DestinationPath (Path to Copy File to Folder To)
3. DeleteSourceFolders (Boolean to determine if source folder should be deleted)

* Can be enhanced to only delete sub folders and ignore source folder

## ErrorHandler

> Generic Way to Handle Error (Send Email w Screenshot)

Important Variables
1. in_ProcessName (Name of Process which threw an error)
2. in_Exception 
3. in_Config (Uses Value from Config File to reduce Copy Paste) - Looks for Template & Email Account

* Can be enhanced to cater for your own needs, this is only to serve as a starter template

## Cell Merge

> Merges Cell in a Excel File

Important Variables
1. in_ExcelPath (Path to Excel File)
2. in_Sheet (Sheet Name where merging occurs)
3. in_RangeSeries (Array of Cell Range to merge e.g. {"A1:A2", "B3:C3"} )

## Excel_Borders

> Border Line Styles for All Cells in a Range

Important Variables
1. in_ExcelPath (Path to Excel File)
2. in_Sheet (Sheet Name where Bordering Style occurs)
3. in_RangeSeries (Array of Cell Range to have Border Styles e.g. {"A1:A2", "B3:C3"} )

* Can be enhanced to suit your needs (E.g. Border outisde only or Thicker Border w Colors)

## CopyWorksheetToWorkbook

> Copies Excel Worksheet from Workbook to Another Workbook (Work with Macro Files)

Important Variables
1. in_ExcelPath (Path to Excel File)
2. in_Sheet (Sheet Name to Copy From)
3. in_Destination (Full Path to New Excel File to be created)

* It will also paste all used cells as values instead of formula/ references.

## ConfigReader

> Read Excel Sheet for Dictionary Values while ensuring no overwriting of values

Important Variables
1. in_TargetObject (Target Name to retrieve from Windows Credential Manager)

* Can be editted to not use from Credential Manager instead

Prone to Breaking - In-Pro Summary, EXCO Report, In-Pro Report

# References
1. https://download.uipath.com/UiPathStudio.msi
2. https://download.uipath.com/versions/18.4.6/UiPathPlatformInstaller.exe
3. https://download.uipath.com/versions/18.4.6/UiPathStudio.msi
