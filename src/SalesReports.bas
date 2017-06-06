Attribute VB_Name = "SalesReports"
Option Explicit

' Macro to validate the Base Margin rates for products and highlight the products with
' multiple margin rates.

' Created by:  Jan Davidson
' Created date:  10 May 2017
' Data:  The data in this workbook was downloaded from the web:
' http://vader.lab.asu.edu/education/papers/2015/Sample%20-%20Superstore%20Sales%20(Excel).xls

' The subroutines and functions (macros) in this module are:
'   AcceptChanges       Copies the new base margin values to the Macro worksheet's Base Margin
'                       column, and recalculates the "Macro" worksheet and all pivot tables;
'                       retains the highlighting in the Current Base column save the workbook.
'   CaclSupportRevenue  Calculates the support revenue for specific product sub-categories;
'   CheckRate           Compares the base margin rates for products with the same Product Name;
'                       if the base margins are different, enters the lower rate in the New Base
'                       column; highlights the row with a different margin.
'   CopyData            Copies Order ID, Product Name, and Base Margin columns from the "Macro"
'                       to the new "BaseCheck" worksheet.
'   CountRows           Function to count the number of rows on a worksheet.
'   FindReturns         Locates the orders that were returned; subtracts the amounts from the
'                       totals; colors the text red for the row with the returned order; copies
'                       returned orders from Returns worksheet; updates pivot tables.
'   NewWorksheet        Creates a new worksheet named "BaseCheck"; adds button "Accept Changes";
'                       adds column headings.
'   SaveWorkbook        Displays the "Save As" disalog box to save the workbook.
'   SortBaseCheckWksht  Sorts the "Base Check" worksheet by Product Name and Order ID.
'   SortMacroWksht      Sorts the "Macro" worksheet by Product Name and Order ID.
'   ValidateMargin      Main routine for Validate Margins button; calls NewWorksheet, CopyData,
'                       and CheckRate.

' Define Public Constants
' Columns for Macro worksheet
Public Const ORDIDCOL As Integer = 1            ' order id column; same column on Macro, BaseCheck worksheets
Public Const ORDDTCOL As Integer = 2            ' order date column
Public Const MPRODCOL As Integer = 8            ' product name column on Macro worksheet
Public Const ORDQTYCOL As Integer = 9           ' order quantity column on Macro worksheet
Public Const SHIPCOSTCOL As Integer = 18        ' shipping cost column on Macro worksheet
Public Const SHIPMODECOL As Integer = 20        ' shipping mode column on Macro worksheet
Public Const SHIPDTCOL As Integer = 21          ' shipping date column on Macro worksheet
Public Const MCURRBASECOL As Integer = 24       ' current base margin column on Macro worksheet
Public Const MGRCOL As Integer = 25             ' manager column
Public Const RETURNSCOL As Integer = 27         ' scratch column for holding returns order ids

' Columns for BaseCheck Worksheet
Public Const NPRODCOL As Integer = 2            ' product name column
Public Const NCURRBASECOL As Integer = 3        ' current base margin column
Public Const NEWBASECOL As Integer = 4          ' new base margin column

' Columns for SuppRev Worksheet
Public Const CUSTCOL As Integer = 1             ' customer name column
Public Const PRODSUBCOL As Integer = 2          ' product sub-category column
Public Const QTYCOL As Integer = 6              ' quantity column
Public Const SUPPHRSCOL As Integer = 7          ' support hours column
Public Const SUPPRTCOL As Integer = 8           ' support rate column
Public Const REVCOL As Integer = 9              ' support revenue column

' lookup table for support hours and rates
Public Const ITEMCOL As Integer = 11            ' product sub-category
Public Const UNDERQTYCOL As Integer = 12        ' quantity (<=) for support hours & rate
Public Const HOURSCOL As Integer = 14           ' hours for sub-category
Public Const RATECOL As Integer = 13            ' support rate

' Columns for PT SuppBase Worksheet
Public Const CUSTNAMECOL As Integer = 1         ' customer name column on support pivot table
Public Const SUMQTYCOL As Integer = 6           ' quantity column on support pivot table

' Column for Returns Worksheet
Public Const RETCOL As Integer = 1              ' return order id column

' Rows
Public Const FirstRow As Integer = 4            ' first data row of Macro and BaseCheck worksheets
Public Const RFIRSTROW As Integer = 2           ' first data row of Returns worksheet

' Public Variables
Public BcWksht As Worksheet                     ' BaseCheck worksheet object
Public MWksht As Worksheet                      ' Macro worksheet object
Public Returns As Worksheet                     ' Returns worksheet object

Public LastRow As Long                          ' last row Macro and BaseCheck worksheets
Public RLastRow As Long                         ' last row Returns on Macro worksheet
Public SLastRow As Long                         ' last row Supp Rev worksheet prod subcat col
Public LLastRow As Long                         ' last row Supp Rev worksheet lookup table item col
Public PLastRow As Long                         ' last row PT SuppBase worksheet (pivot table)
Public RwLastRow As Long                        ' last row of returns col on Returns workskheet

'==================================================================================================

Sub AcceptChanges()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim BcWksht As Worksheet
Dim MWksht As Worksheet
Dim PvtTbl As PivotTable

    ' sort the Macro and BaseCheck worksheets in Product Sub-category
    ' and Order ID order
    Call SortMacroWksht
    Call SortBaseCheckWksht
    
    Set BcWksht = ThisWorkbook.Worksheets("BaseCheck")
    Set MWksht = ThisWorkbook.Worksheets("Sheet14")
    
    ' copy New Base Margin from the BaseCheck to the Macro worksheet
        BcWksht.Select
        Range(Cells(FirstRow, NEWBASECOL), Cells(FirstRow, NEWBASECOL)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
        MWksht.Select
        Cells(FirstRow, MCURRBASECOL).Select
        Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
            , SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    BcWksht.Select
    Cells(FirstRow, ORDIDCOL).Select
    
    MWksht.Select
    Cells(FirstRow, ORDIDCOL).Select
    
    ' Update Pivot Tables
    For Each PvtTbl In ActiveWorkbook.PivotTables
        PvtTbl.RefreshTable
    Next

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

    ' Save Workbook As
    Call SaveWorkbook

End Sub

'==================================================================================================
Sub CalcSupportRevenue()

' NOTE:  For best results, this macro should be run after the FindReturns macro.

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' Define Variables
Dim SuppRev As Worksheet                ' define SuppRev worksheet object
Dim PivotWs As Worksheet                ' define support revenue pivot table worksheet object
Dim PTSuppBase As PivotTable            ' define support revenue pivot table object
Dim PvtTbl As PivotTable                ' define pivot table object for updates

Dim i As Long                           ' for loop counter
Dim j As Long                           ' for loop counter
Dim GrandTotal As Double                ' grand total calculated variable

    ' set worksheet and pivot table objects
    Set SuppRev = ActiveWorkbook.Worksheets("SuppRev")
    Set PivotWs = ActiveWorkbook.Worksheets("PT SuppBase")
    
    ' select pivot table and refresh
    PivotWs.Select
    Set PTSuppBase = ActiveSheet.PivotTables("SuppBase")
    PTSuppBase.RefreshTable
    PLastRow = CountRows(PivotWs, SUMQTYCOL)
    
    ' copy pivot table data and paste on SuppRev worksheet
        PivotWs.Select
        Range(Cells(FirstRow, CUSTNAMECOL), Cells(FirstRow, SUMQTYCOL)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Cells(FirstRow, CUSTNAMECOL), Cells(PLastRow, SUMQTYCOL)).Select
        Selection.Copy
        Sheets("SuppRev").Select
        Range(Cells(FirstRow, CUSTCOL), Cells(FirstRow, CUSTCOL)).Select
        Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
            , SkipBlanks:=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range(Cells(FirstRow, CUSTCOL), Cells(FirstRow, CUSTCOL)).Select
        
        Columns("A:I").Select
        Columns("A:I").EntireColumn.AutoFit
    
    SuppRev.Select
    Cells(FirstRow, CUSTCOL).Select
    
    SLastRow = CountRows(SuppRev, PRODSUBCOL)
    LLastRow = CountRows(SuppRev, ITEMCOL)
    
    GrandTotal = 0

    ' loop through records and match quantity and product subcategory
    For i = FirstRow To SLastRow
        For j = FirstRow To LLastRow
            'find correct hours and rate for quanity and enter values
            If Cells(i, PRODSUBCOL) = Cells(j, ITEMCOL) And Cells(i, QTYCOL) <= Cells(j, UNDERQTYCOL) Then
                If Cells(i, QTYCOL) > 0 Then
                    Cells(i, SUPPHRSCOL) = Cells(j, HOURSCOL)
                    Cells(i, SUPPRTCOL) = Cells(j, RATECOL)
                    Cells(i, REVCOL) = Cells(i, SUPPHRSCOL) * Cells(i, SUPPRTCOL)
                    GrandTotal = GrandTotal + Cells(i, REVCOL)
                    Exit For
                Else
                    Cells(i, SUPPHRSCOL) = 0
                    Cells(i, SUPPRTCOL) = 0
                    Cells(i, REVCOL) = 0
                End If
            End If
        Next j
    Next i
    
    ' format calculated cells
    Range("G:G,I:I").Select
    Range("I4").Activate
    Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
    ' total Support Rev
    Cells(SLastRow + 1, SUPPRTCOL) = "Total Support"
    Cells(SLastRow + 1, REVCOL) = GrandTotal
    
    Cells(FirstRow, CUSTCOL).Select
    
    ' Update Pivot Tables
    Set PvtTbl = ActiveSheet.PivotTables("SuppRev")
    PvtTbl.RefreshTable
        
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
    
    'save workbook
    Call SaveWorkbook

End Sub

'==================================================================================================
Sub CheckRate()

Dim MarginDict As New Scripting.Dictionary      'create dictionary object to store lowest margin by product
Dim PvtTbl As PivotTable
Dim i As Integer                                'for loop counter
Dim j As Integer                                'for loop counter
Dim MinRate As Double                           'stores rate
Dim Prod As String                              'stores prod name

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    'find number of rows
    LastRow = CountRows(BcWksht, ORDIDCOL)
    
    ' find lowest value for each product name
    For i = FirstRow + 1 To LastRow
        Prod = Cells(i, NPRODCOL).Value
        If MarginDict.Exists(Prod) Then
            If MarginDict(Prod) > Cells(i, NCURRBASECOL) Then
                MarginDict(Prod) = Cells(i, NCURRBASECOL)
            End If
        Else
            MarginDict(Prod) = Cells(i, NCURRBASECOL)
        End If
    Next

    ' compare lowest base margin to current base margin by product, enters lowest base margin
    ' in the NewBaseMargin column; highlights the cell light blue
    For i = FirstRow To ProdData.Rows.Count
        Prod = Cells(i, NPRODCOL)
        MinRate = Cells(i, NCURRBASECOL)
       If MarginDict.Exists(Prod) Then
           If MarginDict(Prod) < MinRate Then
               Cells(i, NEWBASECOL) = MarginDict(Prod)
               Cells(i, NORDIDCOL).Select
               Range(Selection, Selection.End(xlToRight)).Select
               Selection.Interior.ColorIndex = 20
           ElseIf MarginDict(Prod) = MinRate Then
               Cells(i, NEWBASECOL) = MarginDict(Prod)
           End If
       End If
    Next i
    
    ' Update Pivot Tables
    For Each PvtTbl In ActiveWorkbook.PivotTables
        PvtTbl.RefreshTable
    Next

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

'==================================================================================================

Sub CopyData()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


    ' copy order id data from macro to basecheck worksheet
    MWksht.Select
    Range(Cells(FirstRow, ORDIDCOL), Cells(FirstRow, ORDIDCOL)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    BcWksht.Select
    Cells(FirstRow, ORDIDCOL).Select
    ActiveSheet.Paste
        
    ' copy product name data from macro to basecheck worksheet
    MWksht.Select
    Range(Cells(FirstRow, MPRODCOL), Cells(FirstRow, MPRODCOL)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    BcWksht.Select
    Cells(FirstRow, NPRODCOL).Select
    ActiveSheet.Paste
        
    ' copy base margin column from macro to basecheck worksheet
    MWksht.Select
    Range(Cells(FirstRow, MCURRBASECOL), Cells(FirstRow, MCURRBASECOL)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    BcWksht.Select
    Cells(FirstRow, NCURRBASECOL).Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    
    MWksht.Select
    Cells(FirstRow, ORDIDCOL).Select
    
    BcWksht.Select
    Cells(FirstRow, ORDIDCOL).Select
      
    'format cells
    Columns("A:D").Select
    Columns("A:D").EntireColumn.AutoFit
    
    Range("D4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.00_);[Red](0.00)"
    
    SLastRow = CountRows(BcWksht, PRODSUBCOL)
    
    'sort data by product name and order ID
    Call SortBaseCheckWksht

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

End Sub

'================================================================================================

Function CountRows(CountSheet As Worksheet, CountCol As Integer) As Long

Dim RowNbr As Long

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    CountSheet.Select
    
    RowNbr = 4
    
        While Cells(RowNbr, CountCol) <> ""
            RowNbr = RowNbr + 1
        Wend
    
    CountRows = RowNbr - 1

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

End Function

'================================================================================================

Sub FindReturns()

Dim RetWksht As Worksheet
Dim OrdID As Range
Dim FoundAddr As Range
Dim LastAddr As Range
Dim PvtTbl As PivotTable
Dim FirstAddr As String
Dim i As Long

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    Set MWksht = ThisWorkbook.Worksheets("Macro")
    Set RetWksht = ThisWorkbook.Worksheets("Returns")
    
    RetWksht.Select
    
    RwLastRow = CountRows(RetWksht, RETCOL)
    
    Range(Cells(FirstRow, RETCOL), Cells(FirstRow, RETCOL)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    MWksht.Select
    Cells(FirstRow, RETURNSCOL).Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
    
    RLastRow = CountRows(MWksht, RETURNSCOL)
    LastRow = CountRows(MWksht, ORDIDCOL)
    
    Set OrdID = Range(Cells(FirstRow, ORDIDCOL), Cells(LastRow, ORDIDCOL))
    
    With OrdID
        Set LastAddr = .Cells(.Cells.Count)
        Debug.Print "LastAddr = " & LastAddr.Rows.Count
    End With
    
    For i = FirstRow To RLastRow
        ' find first occurrence of returned order id
        Set FoundAddr = OrdID.Find(what:=Cells(i, RETURNSCOL), lookat:=xlWhole, after:=LastAddr)
        ' if found, then store found address
        If Not FoundAddr Is Nothing Then
            FirstAddr = FoundAddr.Address
        End If
        Do Until FoundAddr Is Nothing
            'format returned orders and adjust amounts
            'makes the order quantity negative, so the calcs will perform correctly and be negative amounts
            Cells(FoundAddr.Row, ORDQTYCOL) = Cells(FoundAddr.Row, ORDQTYCOL) * -1
            'remove shipping cost, change ship mode to "Returned", and ship date to order date
            Cells(FoundAddr.Row, SHIPCOSTCOL) = 0
            Cells(FoundAddr.Row, SHIPMODECOL) = "Returned"
            Cells(FoundAddr.Row, SHIPDTCOL) = Cells(FoundAddr.Row, ORDDTCOL)
            Range(Cells(FoundAddr.Row, ORDIDCOL), Cells(FoundAddr.Row, ORDIDCOL)).Select
            'change font color to red
            Range(Selection, Selection.End(xlToRight)).Select
            With Selection.Font
                .Color = -16776961 'red
                .TintAndShade = 0
            End With
            'find next returned order id
            Set FoundAddr = OrdID.FindNext(FoundAddr) '(Cells(i, RETURNSCOL), lookat:=xlWhole, after:=LastAddr)
            'if this address already found, look for next one
            If FoundAddr.Address = FirstAddr Then
                Exit Do
            End If
        Loop
    Next
    
    ' Update pivot tables after returns
    For Each PvtTbl In ActiveWorkbook.PivotTables
        PvtTbl.RefreshTable
    Next

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

'================================================================================================

Sub NewWorksheet()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'Create new worksheet named 'BaseCheck' and check if exists
    On Error Resume Next
    Set BcWksht = ThisWorkbook.Worksheets("BaseCheck")
    
    On Error GoTo 0
    If Not BcWksht Is Nothing Then
        MsgBox "Worksheet exists. Please delete and retry macro."
        End
    Else
        Sheets.Add.Name = "BaseCheck"
        Set BcWksht = Worksheets("BaseCheck")
    End If
    
    'Add a button to the worksheet to approve the changes
    BcWksht.Select

    'button add and location of add
    ActiveSheet.Buttons.Add(160, 0.666666666666667, 89, 23.6666666666667).Select 'left,top,witdth,height
    Selection.OnAction = "AcceptChanges"           'name of macro to run when button pressed
    
    Selection.Characters.Text = "Accept Changes"   'button text
    
    With Selection.Font
        .Name = "MS Sans Serif"
        .FontStyle = "Bold"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 10
        '.TintAndShade = 0
    End With

    'Add Column Headings
    Cells(FirstRow - 1, ORDIDCOL) = "Order ID"
    Cells(FirstRow - 1, NPRODCOL) = "Product Name"
    Cells(FirstRow - 1, NCURRBASECOL) = "Current Base Margin"
    Cells(FirstRow - 1, NEWBASECOL) = "New Base Margin"

    'Format Headings
    Range("A3", "D3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    
    With Selection.Font
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .Bold = True
    End With
       
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

'================================================================================================

Sub SaveWorkbook()

Dim WkbkName As String
Dim sFileSaveName As Variant

    WkbkName = ActiveWorkbook.Name
    ' calls Save As dialog box
    sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=WkbkName, fileFilter:="Excel Files (*.xlsm), *.xlsm")
    If sFileSaveName <> False Then
        ActiveWorkbook.SaveAs sFileSaveName
    End If

End Sub

'================================================================================================

Sub SortBaseCheckWksht()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    'sort BaseCheck Worksheet by Product Name and Order ID
    Set BcWksht = ActiveWorkbook.Worksheets("BaseCheck")
    BcWksht.Select
    LastRow = CountRows(BcWksht, ORDIDCOL)
        
    Range(Cells(FirstRow, ORDIDCOL), Cells(LastRow, NCURRBASECOL)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    BcWksht.Sort.SortFields.Clear
    BcWksht.Sort.SortFields.Add Key:=Range(Cells(FirstRow, NPRODCOL), Cells(LastRow, NPRODCOL)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    BcWksht.Sort.SortFields.Add Key:=Range(Cells(FirstRow, ORDIDCOL), Cells(LastRow, ORDIDCOL)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With BcWksht.Sort
        .SetRange Range(Cells(FirstRow, ORDIDCOL), Cells(LastRow, NCURRBASECOL))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(Cells(FirstRow, ORDIDCOL), Cells(FirstRow, ORDIDCOL)).Select

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

'================================================================================================
Sub SortMacroWksht()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    ' sort Macro worksheet by Product Name and Order ID
    Set MWksht = ActiveWorkbook.Worksheets("Macro")
    
    MWksht.Select
    LastRow = CountRows(MWksht, ORDIDCOL)
        
    Range(Cells(FirstRow, ORDIDCOL), Cells(LastRow, MGRCOL)).Select
    Range(Selection, Selection.End(xlDown)).Select
    
    MWksht.Sort.SortFields.Clear
    
    MWksht.Sort.SortFields.Add Key:=Range(Cells(FirstRow, MPRODCOL), Cells(LastRow, MPRODCOL)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    
    MWksht.Sort.SortFields.Add Key:=Range(Cells(FirstRow, ORDIDCOL), Cells(LastRow, ORDIDCOL)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With MWksht.Sort
        .SetRange Range(Cells(FirstRow, ORDIDCOL), Cells(LastRow, MGRCOL))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range(Cells(FirstRow, ORDIDCOL), Cells(FirstRow, ORDIDCOL)).Select

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

'================================================================================================

Sub ValidateMargin()

Application.Workbooks("SampleReports").Activate     ' activate workbook

Application.ScreenUpdating = False                  ' turn off update display
Application.Calculation = xlCalculationManual       ' turn off auto calculation

    Set MWksht = ThisWorkbook.Worksheets("Macro")   ' set MWksht = Macro worksheet
    LastRow = CountRows(MWksht, ORDIDCOL)           ' count number of rows of data on Macro worksheet
    Call NewWorksheet                               ' Insert new worksheet
    Call CopyData                                   ' copy data from Macro to BaseCheck worksheet
    Call CheckRate                                  ' check base margin rates

Application.ScreenUpdating = True                   ' update display
Application.Calculation = xlCalculationAutomatic    ' turn on auto calculation

End Sub

'================================================================================================
















