# UiPath-Snippets
Snippet Repository (One-Stop Hack Portal)

## Convert_DataTable_X_HTML

> Convert Datatable Object to HTML Table to be used for Sending of Emails.

Important Variables
1. in_dictionaryTableFlag (Determine if key pair value or row X column format)

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

