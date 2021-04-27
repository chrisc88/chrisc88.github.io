<a name="top"></a>
## Portfolio
###### [ 407.259.0025 ] . [ [chris.cooper.ofl@gmail.com](mailto:chris.cooper.ofl@gmail.com) ] . [ [Resume](https://chrisc88.github.io/Resume) ]

---

### Metabase Dashboard (Prototype)
<br>
<a target="_blank" href="https://chrisc88.github.io/images/MetabaseDashboard.jpg">
  <img src="https://chrisc88.github.io/images/MetabaseDashboard.jpg" alt="Dashboard" style="width:70%">
</a>
<br>

**Tech/Tooling:** Metabase, PostgreSQL, Excel

---

### Sales Reporting (User configurable groupings by: Department, Area, Category, SubCategory, Product, Customer, Employee, Hour) 
<br>
<a target="_blank" href="https://chrisc88.github.io/images/SalesDepartment.jpg">
  <img src="https://chrisc88.github.io/images/SalesDepartment.jpg" alt="SalesDept" style="width:70%">
</a>
<br>
<a target="_blank" href="https://chrisc88.github.io/images/SalesProduct.jpg">
  <img src="https://chrisc88.github.io/images/SalesProduct.jpg" alt="SalesProduct" style="width:70%">
</a>
<br>
<a target="_blank" href="https://chrisc88.github.io/images/SalesHour.jpg">
  <img src="https://chrisc88.github.io/images/SalesHour.jpg" alt="SalesHour" style="width:70%">
</a>

---

### TeeSheet Weather/Revenue Banner (Highlighted in red and noted with an arrow)

<a target="_blank" href="https://chrisc88.github.io/images/InfoBanner.jpg">
  <img src="https://chrisc88.github.io/images/InfoBanner.jpg" alt="Banner" style="width:70%">
</a>
<br>

**Tech/Tooling:** PostMan, Leveraging: Weather Underground API & BRS API

---

### Automated Player Tagging & Color Coding
<br>
<a target="_blank" href="https://chrisc88.github.io/images/Player_Slot_Coloring.jpg">
  <img src="https://chrisc88.github.io/images/Player_Slot_Coloring.jpg" alt="PlayerColor" style="width:70%">
</a>
<br>
_Above: Player tagging configuration page_
<br>
<a target="_blank" href="https://chrisc88.github.io/images/PSCTeeSheet.jpg">
  <img src="https://chrisc88.github.io/images/PSCTeeSheet.jpg" alt="TeeSheet" style="width:70%">
</a>
<br>
_Above: Example of the player tagging in effect within the tee sheet_
 
---

### GL Summary Report
<br>
<a target="_blank" href="https://chrisc88.github.io/images/GLSummary.jpg">
  <img src="https://chrisc88.github.io/images/GLSummary.jpg" alt="GLSummary" style="width:70%">
</a>
<br>

---

### Accounting Reports (QuickBooks IIF & Great Plains)
<br>
<a target="_blank" href="https://chrisc88.github.io/images/QBSales.jpg">
  <img src="https://chrisc88.github.io/images/QBSales.jpg" alt="QuickBooks" style="width:70%">
</a>
<br>
<a target="_blank" href="https://chrisc88.github.io/images/GPSales.jpg">
  <img src="https://chrisc88.github.io/images/GPSales.jpg" alt="QuickBooks" style="width:70%">
</a>
<br>

---

### PayFac - Settlement Report
<br>
<a target="_blank" href="https://chrisc88.github.io/images/SettlementReport.jpg">
  <img src="https://chrisc88.github.io/images/SettlementReport.jpg" alt="Settlement" style="width:70%">
</a>

---

### PayFac - Transaction Report
<br>
<a target="_blank" href="https://chrisc88.github.io/images/TransactionsReport.jpg">
  <img src="https://chrisc88.github.io/images/TransactionsReport.jpg" alt="Settlement" style="width:70%">
</a>

---

### Excel Macro - Compile and Summarize PayFac Data

```VBA
Sub SummarySheet()
'
' Create by Chris Cooper, 2019 - GolfNow
'
'Delete Unused Sheets
    Dim ws As Worksheet
   
    Application.DisplayAlerts = False
   
    For Each ws In Worksheets
        Select Case ws.Name
            Case "Invoice", "SIG", "PIN", "ChainMerchantDifference"
            Case Else
                ws.Delete
        End Select
    Next ws

    Application.DisplayAlerts = True
    
'Unhide Rows/Columns For All Sheets
    For Each ws In ThisWorkbook.Worksheets
            ws.Cells.EntireColumn.Hidden = False
            ws.Cells.EntireRow.Hidden = False
    Next ws
    
'Create New Sheet & Assign Column Headers
    Sheets("Invoice").Select
    Sheets.Add Before:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "SummarySheet"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Merchant"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Interchange"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Invoice"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Pin"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Interchange Volume"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Pin Volume"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Blended Rate"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Interchange %"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Invoice %"""
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Pin %"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Statement Date"
    
'Merge Merchant Names & Remove Duplicates
    Sheets("SIG").Select
    ActiveCell.Offset(1, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("SummarySheet").Select
    ActiveCell.Offset(1, -13).Range("A1").Select
    ActiveSheet.Paste
    Sheets("PIN").Select
    ActiveCell.Offset(1, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("SummarySheet").Select
    ActiveCell.Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Columns("A:A").EntireColumn.Select
    Application.CutCopyMode = False
    ActiveSheet.Range(Selection, Selection.End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo

'Sort and Filter Merchant Names Ascending
    ActiveCell.Cells.Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("SummarySheet").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SummarySheet").AutoFilter.Sort.SortFields.Add2 Key _
        :=ActiveCell.Offset(0, 0).Range("A1:A9312"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("SummarySheet").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Interchange Fee Amount & Apply Down
    ActiveCell.Offset(1, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SIG!C,SummarySheet!RC[-1],SIG!C[7])"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"

'Invoice Fee Amount & Apply Down
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF(Invoice!C[4],SummarySheet!RC[-2],Invoice!C[8])"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"

'Invoice Tab Data Format
    Sheets("Invoice").Select
    ActiveCell.Offset(1, 10).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Selection, DataType:=xlDelimited, _
             TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
             Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
             :=Array(1, 1), TrailingMinusNumbers:=True
            
'Pin Fee Amount & Apply Down
    Sheets("SummarySheet").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(PIN!C[-2],SummarySheet!RC[-3],PIN!C[3])"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
    
'Pin Tab Data Format
    Sheets("PIN").Select
    ActiveCell.Offset(-1, 5).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Selection, DataType:=xlDelimited, _
             TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
             Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
             :=Array(1, 1), TrailingMinusNumbers:=True
             
'Total Fee Sum & Apply Down
    Sheets("SummarySheet").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
    
'Interchange Volume & Apply Down
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(SIG!C[-5],SummarySheet!RC[-6],SIG!C)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
    
'Pin Volume & Apply Down
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(PIN!C2,SummarySheet!RC[-7],PIN!C5)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "$#,##0.00"
    
'Blended Rate & Apply Down
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]/(RC7+RC8)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00%"
    
'Interchange % & Apply Down
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]/(RC7+RC8)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00%"
    
'Invoice % & Apply Down
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]/(RC7+RC8)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00%"
    
'Pin % & Apply Down
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]/(RC7+RC8)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).NumberFormat = "0.00%"

'Statement Date
    Sheets("Invoice").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("SummarySheet").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste

'Detete GhostCell
    On Error Resume Next
 
    Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Columns("F").SpecialCells(xlCellTypeBlanks).EntireColumn.Delete
    Columns("I").SpecialCells(xlCellTypeBlanks).EntireColumn.Delete
    
    Columns("A:N").AutoFit

End Sub
```

---

<a href="#top">Back to top of page</a>