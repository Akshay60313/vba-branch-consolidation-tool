# vba-branch-consolidation-tool
Excel VBA tool that consolidates multi-branch financial data into a single master report in under 10 seconds. Features a custom UserForm, dynamic folder scanning, and full audit trail.

# Multi-Branch Financial Consolidation Tool
### Excel VBA | Automated Financial Reporting | FP&A; Automation

## Overview
A fully automated Excel VBA tool that consolidates financial transaction data from multiple branch
workbooks into a single master report — eliminating manual copy-paste, reducing consolidation time from hours to
seconds, and maintaining a complete audit trail. Built as part of a financial automation portfolio demonstrating
real-world FP&A; process improvement.

## Business Problem Solved 
Finance teams in multi-branch or multi-entity organisations spend significant time manually consolidating data from separate workbooks.
This process is: 
- Time-consuming (2-4 hours per consolidation cycle) 
- Error-prone (manual copy-paste introduces mistakes)
- Difficult to audit (no record of which data came from where) 
- Not scalable (adding branches increases time linearly) This tool solves all four problems simultaneously.

## Features -
- Automated folder scanning 
— detects all Excel files automatically 
- Multi-select UserForm UI
— choose specific branches or select all 
- One-click consolidation 
— runs entire process with single button 
- Source audit column
— every row tagged with originating filename 
- Error handling — skips corrupt or open files gracefully 
- Clean clipboard management
— no lingering copy operations
<img width="1788" height="911" alt="Vba -Consolidation" src="https://github.com/user-attachments/assets/dc812e5c-9765-445c-8882-be903dac1cf9" />




## Technical Implementation Language: 
Excel VBA (Visual Basic for Applications)
Interface: Custom UserForm with ListBox, FileDialog, command buttons Key Methods: Dir() loop,
Workbooks.Open, PasteSpecial xlPasteValues, FileDialog msoFileDialogFolderPicker Patterns: Dynamic array
population, multi-select ListBox, LastRow detection, Application.ScreenUpdating 

## Dataset 
- 6 branch workbooks (North, South, East, West, Central, UAE) - 385-406 transactions per branch - ~2,400 total
transactions consolidated in under 10 seconds - Categories: Revenue, Cost of Goods, Operating Expenses, Payroll


## How It Works
1. User clicks Browse — FileDialog opens for folder selection 
2. Tool scans folder using Dir() loop — populates ListBox with files
3. User selects branches (individual or Select All)
4. Tool opens each workbook, copies data, pastes values to master 
5. Source filename written to audit column for everyrow 
6. Workbook closed — process repeats for next file 
7. Complete consolidated report ready in seconds ---
   
## Key Code
Patterns ```vba ' Dir() loop to scan folder Dim strFile As String strFile = Dir(folderPath & "*.xlsx") Do While
strFile <> "" lstFiles.AddItem strFile strFile = Dir() Loop ' PasteSpecial to avoid formula dependency
wsDest.Cells(destRow, 1).PasteSpecial xlPasteValues Application.CutCopyMode = False ' Audit column — source
tracking wsDest.Cells(destRow, auditCol).Value = wb.Name ``` 

## FP&A; Applications
This pattern applies directly to: 
- Monthly management accounts consolidation (multiple entities) 
- Budget vs actual variance reporting (multiple departments) 
- Sales data aggregation (multiple regions/channels) 
- Cost centre reporting (multiple business units)

## Portfolio Context This is Project 3 of a VBA automation series: 
- Project 1:Automated financial report formatter
- Project 2: Budget variance tracker with MTD/YTD analysis 
- Project 3:Multi-branch consolidation tool (this project)
  
Built by Akshay Jain | linkedin.com/in/akshay-jain-40589633a
