# Macro-Enhanced-Excels

<br />
<div align="center">
  <a href="https://learn.microsoft.com/en-us/office/vba/api/overview/excel">
    <img src="readmeDocs/ExcelBanner.png" alt="Logo">
  </a>

  <h3 align="center">Excel Macro Automation</h3>
</div>
 
 ---
 
<p align="center">
• <a href="#-introduction">Introduction</a> • <a href="#-macros-Catalogue">Macros Catalogue</a> • <a href="#-getting-started">Getting Started</a> • 
 </p>
 
---

## 💡 Introduction

During the course of my internship at British America Tobacco, I picked up VBA to automate some tasks that involved Excel. Little did I know despite it being known as "outdated", it has some incredible uses which I would soon discover and make some applications with them.
I call it: Macro Enhanced Excels

<table>
<tr>
<th width="350rem" align="center"> </th> <th align="center"> Function </th> </th> <th align="center"> Subs </th>
</tr>
<tr></tr>
<tr>
<td> Sample Code
<td>
	
```vb
Function FuncNameHere(arg1 As String) As String
	MsgBox "You passed in a string argument: " & arg1
	' Returning argument passed in to caller
	SubNameHere = arg1 
End Function
```

</td>
<td>
	
```vb
Sub SubNameHere(arg1 As String, arg2 As Integer)
	MsgBox "You passed in a string argument: " & arg1 
	MsgBox "You passed in an integer argument: " & arg2 
End Sub
```

</td>
</tr>
  
<tr>
<td> Arguments </td>
<td align="center">✔️</td>
<td align="center">✔️</td>
</tr>
<tr>
<td> Return Values </td>
<td align="center">✔️</td>
<td align="center">❌</td>
</tr>
<tr>
<td> Used as formulas </td>
<td align="center">❌</td>
<td align="center">❌</td>
</tr>
<tr>
<td> Direct execution by user </td>
<td align="center">❌</td>
<td align="center">✔️</td>
</tr>
</table>


## 📜 Excel Catalogue 

|       🤖   Excel            |  ⚙️ Functionalities                    | 
| :--------------------------: || :------------------------------------ | 
| [Search Keys in Directory](https://github.com/LimJiaEarn/ExcelMacroAutomations)  · Search in your file system an Excel file that contain a certain cell value <br> · Formatted search result with better insights/filters                 | 


## 🤸 Getting Started

1. **Find & Download your required excel files**
   You can find the available excel macros [here](macros-Catalogue).


2. **Open the excel macros** 
   > **Note**
   > You need to have organisation's permission to run macro-enabled excels or you need to [enable it manually](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)


## 🧬 Building Blocks of my Macros 
Here are the individual sub components I have developed which are the building blocks for the broader macro features within the Excel files.

<details>
<summary> 1) Handles exception of deleting a given table name in active sheet. Returns True if deleted, False if table not found </summary>
	
```vb
Function DeleteTable(tableName As String) As Boolean
    On Error Resume Next
	Dim tbl As ListObject
	Set tbl = ActiveSheet.ListObjects(tableName)
    On Error GoTo 0
    If Not tbl Is Nothing Then
	tbl.Delete
	DeleteTable = True
	Exit Function
    End If
    DeleteTable = False
End Function
```
	
</details>

<details>
<summary> 2) Reads the data in a given CSV filepath and return the data in an Array Object </summary>
	
```vb
Function ReadCSV(filePath As String) As Object
    Dim searchKeyList As Object
    Set searchKeyList = CreateObject("System.Collections.ArrayList")

    Dim keyString As String ' Stores entire csv file as a string to be processed
    Open filePath For Input As #1
    	keyString = Input$(LOF(1), #1) 
    Close #1

    Dim searchKeys() As String ' Array to store each value in the csv string 
    searchKeys = Split(keyString, ",")

    ' Filter out whitespace / newline characters / empty values
    Dim i As Long
    For i = LBound(searchKeys) To UBound(searchKeys)
	Dim key As String
	key = Trim(Replace(searchKeys(i), vbNewLine, ""))

	If Len(key) > 0 Then
	    searchKeyList.Add key
	End If
    Next i

    Set ReadCSV = searchKeyList 
End Function
```
	
</details>

<details>
<summary> 3) Goes through a column of cells containing file paths and hyperlink them </summary>
	
```vb
Sub HyperlinkFilePaths()
    Dim FilePathRange As Range
    Dim cell As Range
    Dim Hyperlink As Hyperlink

    ' Define the range of cells containing file paths
    Set FilePathRange = Range("B2:B3307")

    ' Loop through each cell in the range
    For Each cell In FilePathRange
	' Create a hyperlink for each non-empty cell
	If Len(cell.Value) > 0 Then
	    Set Hyperlink = ActiveSheet.Hyperlinks.Add(Anchor:=cell, Address:=cell.Value, TextToDisplay:=cell.Value)
	    ' Customize the formatting of the hyperlink
	    Hyperlink.Range.Font.Color = RGB(48, 105, 248) ' Blue color
	End If
    Next cell
End Sub
```
	
</details>

<details>
<summary> 4) Given the top left cell of a data range, auto detect the whole data range and table it </summary>
	
```vb
Sub CreateTableFromTopLeftCell(topLeftCell As String) 

    ' Extract the columns and rows of the top left cell
    Dim topLeftRow As Long
    Dim topLeftColumn As Long
    topLeftRow = Range(topLeftCell).Row
    topLeftColumn = Range(topLeftCell).Column

    ' Extract the columns and rows of the bottom right cell
    Dim lastRightColumn As Long
    Dim lastRightRow As Long
    lastRightColumn = ActiveSheet.Cells(topLeftRow, ActiveSheet.Columns.Count).End(xlToLeft).Column
    lastRightRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, lastRightColumn).End(xlUp).Row

    ' Finalise the range of the table
    Dim tableRange As String
    tableRange = ActiveSheet.Cells(topLeftRow, topLeftColumn).Address & ":" & ActiveSheet.Cells(lastRightRow, lastRightColumn).Address

    ' Create the table from the range configured
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(tableRange), , xlYes).Name = "CreatedTableFromMacro"

End Sub
```
	
</details>
    
<details>
<summary> 5) Display the time taken for your code/function to run </summary>
	
```vb
Sub TimerFunction() ' Main function for you to append to your code
    Dim startTime As Double
    Dim endTime As Double
	
    startTime = Timer
    ' Your algorithm/entry to function call
    endTime = Timer

    Dim timeTaken As String
    timeTaken = timerBuilder(endTime-startTime) 
    Debug.Print timeTaken ' display the time taken of your code

End Sub

Function timerBuilder(timeTaken As Double) As String ' The function that builds the display message
    
    Dim timeString As String
    Dim timeUnit As Integer

    ' Calculating hrs taken (if any)
    timeUnit = Int(timeTaken / 3600)
    If timeUnit > 0 Then
        timeString = timeUnit & " hrs"
    End If
    
    ' Calculating mins taken (if any)
    timeUnit = Int((timeTaken Mod 3600) / 60)
    If timeUnit > 0 Then
        If timeString <> "" Then
            timeString = timeString & ", "
        End If
        timeString = timeString & timeUnit & " mins"
    End If

    ' Calculating secs taken (if any)
    timeUnit = Int(timeTaken Mod 60)
    If timeUnit > 0 Then
        If timeString <> "" Then
            timeString = timeString & ", "
        End If
        timeString = timeString & timeUnit & " secs"
    End If

    Set timerBuilder = timeString

End Function
```
	
</details>
	
<details>
<summary> 6) Get current excel file name </summary>
	
```vb
Function CurrentfileName() As String
    Dim fileFullName As String
    Dim fileName As String
     
    ' Extract the file name from the file full path
    fileFullName = ThisWorkbook.FullName
    fileName = Mid(fileFullName, InStrRev(fileFullName, "/") + 1)
    
    ' Return the file name without file extension
    CurrentfileName = Left(fileName, InStrRev(fileName, ".") - 1)

End Function
```
	
</details>
	
<details>
<summary> 7) Prints all formulas in a given row (can be adjusted to column too!) in immediate window </summary>
	
```vb
Sub PrintFormulasInGivenRow(rowNum As Long)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastCol As Long
    lastCol = ws.Cells(rowNum, Columns.Count).End(xlToLeft).Column
    
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol))
    
    Dim i As Integer
    i = 1
    
    Dim cell As Range
    For Each cell In rng
        If cell.HasFormula Then
            Debug.Print "' " & Replace(cell.Address, "$", "") & ":"
            Debug.Print "Formula: " & """" & cell.Formula & """" & vbNewLine
            
        End If
        i = i + 1
    Next cell
End Sub
```
	
</details>

<details>
<summary> 8) Opens file explorer to get user to select an excel file. Returns the file path of excel file  </summary>

```vb
Function getFilePath(boxMsg As String) As String

    Dim filePath As Variant
    filePath = Application.GetOpenFilename(Title:=boxMsg, FileFilter:="Excel Files (*.xlsx; *.xls; *.xlsm), *.xlsx; *.xls")
    If filePath <> False Then
        getFilePath = filePath
    Else
        getFilePath = "CANCELLED"
    End If

End Function
```

</details>

## ⚠️ Possible Limitations

Macros that searches in folder and file paths are only tested in windows 11. Mac compatability is not tested.



## 🌟 Spread the joy!
**Share** these macros with your colleagues or friends on social media.

<a href="https://www.reddit.com" target="_blank">
 <img src="https://img.shields.io/twitter/url?label=Reddit&logo=Reddit&style=social&url=https://www.reddit.com/" alt="Share on Reddit"/></a>&nbsp;
<a href="https://www.linkedin.com" target="_blank">
 <img src="https://img.shields.io/twitter/url?label=LinkedIn&logo=LinkedIn&style=social&url=https://www.linkedin.com" alt="Share on LinkedIn"/></a>&nbsp;
<a href="https://twitter.com" target="_blank">
 <img src="https://img.shields.io/twitter/url?label=Twitter&logo=Twitter&style=social&url=https://twitter.com" alt="Shared on Twitter"/></a>&nbsp;
<a href="https://www.facebook.com" target="_blank">
 <img src="https://img.shields.io/twitter/url?label=Facebook&logo=Facebook&style=social&url=https://www.facebook.com" alt="Share on Facebook"/></a>&nbsp;
<a href="https://t.me/share" target="_blank">
 <img src="https://img.shields.io/twitter/url?label=Telegram&logo=Telegram&style=social&url=https://t.me/share" alt="Share on Telegram"/></a>&nbsp;
<a href="https://wa.me" target="_blank">
 <img src="https://img.shields.io/twitter/url?label=Whatsapp&logo=Whatsapp&style=social&url=https://wa.me" alt="Share on Whatsapp"/></a>&nbsp;

