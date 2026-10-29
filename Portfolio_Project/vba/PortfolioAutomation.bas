Attribute VB_Name = "PortfolioAutomation"
Option Explicit

' Find the transaction table used by the buttons.
Private Function GetTransactionsTable() As ListObject
    Set GetTransactionsTable = ThisWorkbook.Worksheets("Transactions").ListObjects("tblTransactions")
End Function

' Find the investment list used by the buttons.
Private Function GetSecuritiesTable() As ListObject
    Set GetSecuritiesTable = ThisWorkbook.Worksheets("Securities").ListObjects("tblSecurities")
End Function

' Check that the chosen investment is already in the list.
Private Function SecurityExists(ByVal securityID As String) As Boolean
    SecurityExists = Application.CountIf(GetSecuritiesTable.ListColumns("SecurityID").DataBodyRange, securityID) > 0
End Function

' Accept an Excel date or a date typed as yyyy-mm-dd.
Private Function ReadDate(ByVal cellValue As Variant) As Date
    Dim dateText As String
    If VarType(cellValue) = vbDate Then
        ReadDate = CDate(cellValue)
    ElseIf IsNumeric(cellValue) And CStr(cellValue) <> "" Then
        ReadDate = CDate(cellValue)
    Else
        dateText = CStr(cellValue)
        If Len(dateText) <> 10 Or Mid(dateText, 5, 1) <> "-" Or Mid(dateText, 8, 1) <> "-" Then
            Err.Raise 5, , "Use yyyy-mm-dd dates."
        End If
        ReadDate = DateSerial(CInt(Left(dateText, 4)), CInt(Mid(dateText, 6, 2)), CInt(Right(dateText, 2)))
        If Format(ReadDate, "yyyy-mm-dd") <> dateText Then
            Err.Raise 5, , "Invalid date."
        End If
    End If
End Function

' Add the entered values, then copy the formulas from the row above.
Private Sub WriteTransaction(ByVal transactionValues As Variant)
    Dim transactionsTable As ListObject
    Dim newRow As ListRow
    Dim columnNumber As Long
    Set transactionsTable = GetTransactionsTable
    Set newRow = transactionsTable.ListRows.Add
    For columnNumber = 1 To 13
        newRow.Range.Cells(1, columnNumber).Value = transactionValues(columnNumber)
    Next columnNumber
    newRow.Range.Cells(1, 2).NumberFormat = "yyyy-mm-dd"
    newRow.Range.Cells(1, 3).NumberFormat = "yyyy-mm-dd"
    If newRow.Index > 1 Then
        newRow.Range.Cells(1, 14).Resize(1, 12).FormulaR1C1 = transactionsTable.ListRows(newRow.Index - 1).Range.Cells(1, 14).Resize(1, 12).FormulaR1C1
    End If
End Sub

' Put quotes around CSV values so commas in names are handled correctly.
Private Function MakeCsvField(ByVal cellValue As Variant) As String
    Dim quoteMark As String
    quoteMark = Chr(34)
    MakeCsvField = quoteMark & Replace(CStr(cellValue), quoteMark, quoteMark & quoteMark) & quoteMark
End Function

' Run the Python price file, then copy the results into Securities.
Public Sub RefreshPrices()
    On Error GoTo Failed
    Dim csvFile As Object
    Dim commandWindow As Object
    Dim symbolFile As String
    Dim priceFile As String
    Dim pythonFile As String
    Dim pythonCommand As String
    Dim pythonLauncher As String
    Dim csvLine As String
    Dim activeFlag As String
    Dim quoteMark As String
    Dim rowNumber As Long
    Dim exitCode As Long
    Dim priceWorkbook As Workbook
    Dim priceSheet As Worksheet
    Dim securityRow As Range
    Dim matchNumber As Variant
    Dim pricesUpdated As Long
    Dim pricesFailed As Long
    Dim errorMessage As String
    If ThisWorkbook.Path = "" Then
        Err.Raise 5, , "Save the workbook in the extracted project folder first."
    End If
    pythonFile = ThisWorkbook.Path & Application.PathSeparator & "python" & Application.PathSeparator & "refresh_prices.py"
    symbolFile = ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator & "price_symbols.csv"
    priceFile = ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator & "prices_latest.csv"
    If Dir(pythonFile) = "" Then
        Err.Raise 5, , "Missing python/refresh_prices.py. Extract the complete project ZIP and keep the workbook beside the python folder."
    End If
    Set csvFile = CreateObject("ADODB.Stream")
    csvFile.Type = 2
    csvFile.Charset = "utf-8"
    csvFile.Open
    csvFile.WriteText "SecurityID,ProviderSymbol,ExpectedCurrency,ExpectedNameToken,AssetType,ActiveFlag" & vbCrLf
    For rowNumber = 1 To GetSecuritiesTable.ListRows.Count
        With GetSecuritiesTable.ListRows(rowNumber).Range
            If .Cells(1, 12).Value > 0.000001 Then
                activeFlag = "1"
            Else
                activeFlag = "0"
            End If
            csvLine = MakeCsvField(.Cells(1, 1).Value) & "," & MakeCsvField(.Cells(1, 21).Value) & "," & _
            MakeCsvField(.Cells(1, 7).Value) & "," & MakeCsvField(.Cells(1, 22).Value) & "," & MakeCsvField(.Cells(1, 4).Value) & "," & MakeCsvField(activeFlag)
        End With
        csvFile.WriteText csvLine & vbCrLf
    Next rowNumber
    csvFile.SaveToFile symbolFile, 2
    csvFile.Close
    pythonLauncher = Trim(ThisWorkbook.Worksheets("Read Me").Range("G5").Value)
    If pythonLauncher = "" Then
        pythonLauncher = "python"
    End If
    quoteMark = Chr(34)
    pythonCommand = pythonLauncher & " " & quoteMark & pythonFile & quoteMark
    pythonCommand = pythonCommand & " --manifest " & quoteMark & symbolFile & quoteMark
    pythonCommand = pythonCommand & " --output " & quoteMark & priceFile & quoteMark
    ' Wait for Python to finish before opening its output file.
    Set commandWindow = CreateObject("WScript.Shell")
    exitCode = commandWindow.Run(pythonCommand, 1, True)
    If exitCode <> 0 Then
        Err.Raise 5, , "Python failed with exit code " & exitCode & ". Check Python installation and the launcher in Read Me G5."
    End If
    Workbooks.OpenText Filename:=priceFile, DataType:=xlDelimited, Comma:=True, _
    FieldInfo:=Array(Array(1, xlTextFormat), Array(2, xlTextFormat), Array(8, xlTextFormat), _
    Array(9, xlTextFormat), Array(10, xlTextFormat)), Local:=False
    Set priceWorkbook = ActiveWorkbook
    Set priceSheet = priceWorkbook.Worksheets(1)
    If priceSheet.Cells(1, 1).Value <> "SecurityID" Or priceSheet.Cells(1, 11).Value <> "Status" Then
        Err.Raise 5, , "Unexpected quote output format."
    End If
    For rowNumber = 2 To priceSheet.Cells(priceSheet.Rows.Count, 1).End(xlUp).Row
        matchNumber = Application.Match(priceSheet.Cells(rowNumber, 1).Value, GetSecuritiesTable.ListColumns("SecurityID").DataBodyRange, 0)
        If Not IsError(matchNumber) Then
            Set securityRow = GetSecuritiesTable.ListRows(CLng(matchNumber)).Range
            securityRow.Cells(1, 23).Value = priceSheet.Cells(rowNumber, 11).Value
            If priceSheet.Cells(rowNumber, 11).Value = "OK" Then
                If CDbl(priceSheet.Cells(rowNumber, 6).Value) <= 0 Or CDbl(priceSheet.Cells(rowNumber, 7).Value) <= 0 Then
                    Err.Raise 5, , "The price returned was not valid. Try Refresh prices again."
                End If
                securityRow.Cells(1, 6).Value = priceSheet.Cells(rowNumber, 5).Value
                securityRow.Cells(1, 7).Value = priceSheet.Cells(rowNumber, 4).Value
                securityRow.Cells(1, 8).Value = CDbl(priceSheet.Cells(rowNumber, 6).Value)
                securityRow.Cells(1, 9).Value = CDbl(priceSheet.Cells(rowNumber, 7).Value)
                securityRow.Cells(1, 10).Value = ReadDate(priceSheet.Cells(rowNumber, 8).Value)
                securityRow.Cells(1, 10).NumberFormat = "yyyy-mm-dd"
                securityRow.Cells(1, 11).Value = "https://finance.yahoo.com/quote/" & priceSheet.Cells(rowNumber, 2).Value & "/"
                securityRow.Cells(1, 24).Value = priceSheet.Cells(rowNumber, 9).Value
                securityRow.Cells(1, 25).Value = ReadDate(priceSheet.Cells(rowNumber, 10).Value)
                securityRow.Cells(1, 25).NumberFormat = "yyyy-mm-dd"
                pricesUpdated = pricesUpdated + 1
            ElseIf priceSheet.Cells(rowNumber, 11).Value = "FAILED" Then
                ' Keep the old price, but mark it as not ready to use.
                pricesFailed = pricesFailed + 1
            End If
        End If
    Next rowNumber
    priceWorkbook.Close False
    Set priceWorkbook = Nothing
    RefreshPortfolio
    ThisWorkbook.Worksheets("Dashboard").Range("B12").Value = "Prices updated for " & pricesUpdated & " open positions. " & pricesFailed & " prices need attention."
    MsgBox pricesUpdated & " quotes updated; " & pricesFailed & " failed (open positions). Review Securities and data/prices_latest.csv."
    Exit Sub
Failed:
    errorMessage = Err.Description
    On Error Resume Next
    If Not priceWorkbook Is Nothing Then
        priceWorkbook.Close False
    End If
    If Not csvFile Is Nothing Then
        csvFile.Close
    End If
    GetSecuritiesTable.ListColumns("QuoteStatus").DataBodyRange.Value = "FAILED"
    Application.CalculateFull
    MsgBox "Price refresh failed: " & errorMessage & ". Previous numeric quotes retained but marked invalid.", vbExclamation
End Sub

' Ask for the listing details before adding a new investment.
Public Sub AddSecurity()
    On Error GoTo Failed
    Dim securityID As String
    Dim tickerSymbol As String
    Dim securityName As String
    Dim exchangeName As String
    Dim currencyName As String
    Dim newRow As ListRow
    Dim providerSymbol As String
    Dim nameCheck As String
    Dim assetType As String
    securityID = UCase(Trim(InputBox("Enter an investment ID, such as NVDA_US or NVDA_CDR. Each listing needs its own ID.")))
    If securityID = "" Then
        Exit Sub
    End If
    If SecurityExists(securityID) Then
        Err.Raise 5, , "That ID already exists. Use it on Add Transaction."
    End If
    tickerSymbol = UCase(Trim(InputBox("Ticker symbol:")))
    If tickerSymbol = "" Then
        Exit Sub
    End If
    securityName = Trim(InputBox("Full verified security name:"))
    If securityName = "" Then
        Exit Sub
    End If
    exchangeName = UCase(Trim(InputBox("Verified exchange code, for example TSX, NASDAQ, NYSE or CBOE CANADA:")))
    If exchangeName = "" Then
        Exit Sub
    End If
    currencyName = UCase(Trim(InputBox("Quote currency: CAD or USD")))
    If currencyName <> "CAD" And currencyName <> "USD" Then
        Err.Raise 5, , "Only CAD and USD quotes are supported."
    End If
    providerSymbol = Trim(InputBox("Yahoo Finance listing symbol, for example XEQT.TO, FINN.NE or UPS. Include the correct exchange suffix."))
    If providerSymbol = "" Then
        Exit Sub
    End If
    nameCheck = Trim(InputBox("Enter part of the investment name to check, such as Apple or Vanguard:", , Split(securityName, " ")(0)))
    If nameCheck = "" Then
        Exit Sub
    End If
    assetType = UCase(Trim(InputBox("Asset type: Equity, ETF or CDR", , "Equity")))
    If assetType <> "EQUITY" And assetType <> "ETF" And assetType <> "CDR" Then
        Err.Raise 5, , "Use Equity, ETF or CDR."
    End If
    Set newRow = GetSecuritiesTable.ListRows.Add
    newRow.Range.Cells(1, 1).Value = securityID
    newRow.Range.Cells(1, 2).Value = tickerSymbol
    newRow.Range.Cells(1, 3).Value = securityName
    newRow.Range.Cells(1, 4).Value = assetType
    newRow.Range.Cells(1, 5).Value = "Other"
    newRow.Range.Cells(1, 6).Value = exchangeName
    newRow.Range.Cells(1, 7).Value = currencyName
    If currencyName = "CAD" Then
        newRow.Range.Cells(1, 9).Value = 1
    End If
    newRow.Range.Cells(1, 11).Value = ""
    newRow.Range.Cells(1, 21).Value = providerSymbol
    newRow.Range.Cells(1, 22).Value = nameCheck
    newRow.Range.Cells(1, 23).Value = "PENDING"
    newRow.Range.Cells(1, 12).Resize(1, 9).FormulaR1C1 = GetSecuritiesTable.ListRows(newRow.Index - 1).Range.Cells(1, 12).Resize(1, 9).FormulaR1C1
    RefreshPrices
    MsgBox "Security added. Review its matched quote details on Securities."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

' Check the form before adding a purchase, sale or cash entry.
Public Sub AddTransaction()
    On Error GoTo Failed
    Dim entrySheet As Worksheet
    Dim transactionDate As Date
    Dim securityID As String
    Dim transactionType As String
    Dim quantity As Double
    Dim cashAmount As Double
    Dim transactionValues(1 To 13) As Variant
    Dim unitsHeld As Double
    Dim rowNumber As Long
    Dim latestDate As Double
    Set entrySheet = ThisWorkbook.Worksheets("Add Transaction")
    transactionDate = ReadDate(entrySheet.Range("C5").Value)
    securityID = UCase(Trim(entrySheet.Range("C6").Value))
    transactionType = UCase(Trim(entrySheet.Range("C7").Value))
    If InStr("|BUY|SELL|DIV|CONT|NRT|REFER|", "|" & transactionType & "|") = 0 Or transactionType = "" Then
        Err.Raise 5, , "Choose a supported transaction type."
    End If
    If Not IsNumeric(entrySheet.Range("C9").Value) Or entrySheet.Range("C9").Value = "" Then
        Err.Raise 5, , "Enter signed cash amount CAD."
    End If
    cashAmount = CDbl(entrySheet.Range("C9").Value)
    If (transactionType = "BUY" Or transactionType = "NRT") And cashAmount >= 0 Then
        Err.Raise 5, , "BUY / NRT cash amounts must be negative."
    End If
    If transactionType <> "BUY" And transactionType <> "NRT" And cashAmount <= 0 Then
        Err.Raise 5, , "Cash amount must be positive."
    End If
    If transactionType = "BUY" Or transactionType = "SELL" Or transactionType = "DIV" Then
        If Not SecurityExists(securityID) Then
            Err.Raise 5, , "Click Add security first, then try again."
        End If
    Else
        securityID = ""
    End If
    quantity = 0
    If transactionType = "BUY" Or transactionType = "SELL" Then
        If Not IsNumeric(entrySheet.Range("C8").Value) Then
            Err.Raise 5, , "Enter quantity."
        End If
        quantity = CDbl(entrySheet.Range("C8").Value)
        If quantity <= 0 Then
            Err.Raise 5, , "Quantity must be positive."
        End If
    End If
    latestDate = Application.Max(GetTransactionsTable.ListColumns("Date").DataBodyRange)
    If CDbl(transactionDate) < latestDate Then
        Err.Raise 5, , "This date is earlier than the last entry. Put the records in date order before adding it."
    End If
    Application.CalculateFull
    If transactionType = "SELL" Then
        unitsHeld = Application.SumIf(GetTransactionsTable.ListColumns("SecurityID").DataBodyRange, securityID, GetTransactionsTable.ListColumns("SignedQty").DataBodyRange)
        If quantity > unitsHeld + 0.000001 Then
            Err.Raise 5, , "You entered more units than this investment holds."
        End If
        If Application.CountIfs(GetTransactionsTable.ListColumns("SecurityID").DataBodyRange, securityID, GetTransactionsTable.ListColumns("LedgerStatus").DataBodyRange, "REVIEW") > 0 Then
            Err.Raise 5, , "Check the flagged records for this investment before adding its sale."
        End If
    End If
    For rowNumber = 1 To GetTransactionsTable.ListRows.Count
        With GetTransactionsTable.ListRows(rowNumber).Range
            If .Cells(1, 2).Value = transactionDate And _
                .Cells(1, 4).Value = securityID And _
                .Cells(1, 5).Value = transactionType And _
                .Cells(1, 6).Value = quantity And _
                .Cells(1, 7).Value = cashAmount Then
                If MsgBox("This entry looks the same as an existing one. Add it again?", vbYesNo + vbExclamation) <> vbYes Then
                    Exit Sub
                End If
                Exit For
            End If
        End With
    Next rowNumber
    transactionValues(1) = "MAN-" & Format(Now, "yyyymmddhhnnss") & "-" & CStr(GetTransactionsTable.ListRows.Count + 1)
    transactionValues(2) = transactionDate
    transactionValues(3) = transactionDate
    transactionValues(4) = securityID
    transactionValues(5) = transactionType
    transactionValues(6) = quantity
    transactionValues(7) = cashAmount
    transactionValues(8) = "CAD"
    transactionValues(9) = ""
    transactionValues(10) = entrySheet.Range("C11").Value
    transactionValues(11) = ""
    transactionValues(12) = entrySheet.Range("C10").Value
    transactionValues(13) = "Manual"
    WriteTransaction transactionValues
    RefreshPrices
    entrySheet.Range("C8:C10").ClearContents
    MsgBox "Transaction added."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

' Import new records without adding the same RecordID twice.
Public Sub ImportCleanedCSV()
    On Error GoTo Failed
    Dim fileName As Variant
    Dim importWorkbook As Workbook
    Dim importSheet As Worksheet
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim recordID As String
    Dim existingRecords As Object
    Dim securityID As String
    Dim transactionValues(1 To 13) As Variant
    Dim rowsAdded As Long
    Dim rowsSkipped As Long
    Dim latestDate As Double
    Dim lastRow As Long
    Dim transactionDate As Date
    Dim errorMessage As String
    fileName = Application.GetOpenFilename("CSV files (*.csv),*.csv", , "Choose data/transactions_clean.csv")
    If VarType(fileName) = vbBoolean Then
        Exit Sub
    End If
    Workbooks.OpenText Filename:=CStr(fileName), DataType:=xlDelimited, Comma:=True, _
    FieldInfo:=Array(Array(1, xlTextFormat), Array(2, xlTextFormat), Array(3, xlTextFormat), _
    Array(4, xlTextFormat), Array(5, xlTextFormat), Array(8, xlTextFormat), _
    Array(10, xlTextFormat), Array(12, xlTextFormat), Array(13, xlTextFormat)), Local:=False
    Set importWorkbook = ActiveWorkbook
    Set importSheet = importWorkbook.Worksheets(1)
    If importSheet.Cells(1, 1).Value <> "RecordID" Or importSheet.Cells(1, 13).Value <> "Status" Then
        Err.Raise 5, , "Select the cleaned CSV produced by the supplied Python script."
    End If
    Set existingRecords = CreateObject("Scripting.Dictionary")
    For rowNumber = 1 To GetTransactionsTable.ListRows.Count
        existingRecords(CStr(GetTransactionsTable.ListRows(rowNumber).Range.Cells(1, 1).Value)) = True
    Next rowNumber
    latestDate = Application.Max(GetTransactionsTable.ListColumns("Date").DataBodyRange)
    lastRow = importSheet.Cells(importSheet.Rows.Count, 1).End(xlUp).Row
    ' Check the new records before adding any rows.
    For rowNumber = 2 To lastRow
        If Not existingRecords.Exists(CStr(importSheet.Cells(rowNumber, 1).Value)) Then
            transactionDate = ReadDate(importSheet.Cells(rowNumber, 2).Value)
            If transactionDate <= latestDate Then
                Err.Raise 5, , "Some new rows are dated on or before the last saved date. Nothing was added. Check the order of the records."
            End If
            securityID = CStr(importSheet.Cells(rowNumber, 4).Value)
            If securityID <> "" And Not SecurityExists(securityID) Then
                Err.Raise 5, , "Add security ID " & securityID & " first, then retry."
            End If
            If importSheet.Cells(rowNumber, 8).Value <> "CAD" Then
                Err.Raise 5, , "The cash amounts must be in CAD."
            End If
        End If
    Next rowNumber
    For rowNumber = 2 To lastRow
        recordID = CStr(importSheet.Cells(rowNumber, 1).Value)
        If Not existingRecords.Exists(recordID) Then
            For columnNumber = 1 To 13
                transactionValues(columnNumber) = importSheet.Cells(rowNumber, columnNumber).Value
            Next columnNumber
            transactionValues(2) = ReadDate(importSheet.Cells(rowNumber, 2).Value)
            transactionValues(3) = ReadDate(importSheet.Cells(rowNumber, 3).Value)
            WriteTransaction transactionValues
            existingRecords(recordID) = True
            rowsAdded = rowsAdded + 1
        Else
            rowsSkipped = rowsSkipped + 1
        End If
    Next rowNumber
    importWorkbook.Close False
    RefreshPrices
    MsgBox rowsAdded & " added; " & rowsSkipped & " existing IDs skipped. Review Python data_quality.csv for new warnings."
    Exit Sub
Failed:
    errorMessage = Err.Description
    On Error Resume Next
    If Not importWorkbook Is Nothing Then
        importWorkbook.Close False
    End If
    MsgBox "Import stopped: " & errorMessage, vbExclamation
End Sub

' Update the monthly table, open holdings and charts.
Public Sub RefreshPortfolio()
    On Error GoTo Failed
    Dim dashboard As Worksheet
    Dim monthlyTable As ListObject
    Dim knownMonths As Object
    Dim rowNumber As Long
    Dim monthName As Variant
    Dim dashboardRow As Long
    Dim newMonthRow As ListRow
    Dim securitySheetRow As Long
    Dim lastRow As Long
    Dim unitsHeld As Double
    Set knownMonths = CreateObject("Scripting.Dictionary")
    Set monthlyTable = ThisWorkbook.Worksheets("Monthly").ListObjects("tblMonthly")
    For rowNumber = 1 To monthlyTable.ListRows.Count
        knownMonths(CStr(monthlyTable.ListRows(rowNumber).Range.Cells(1, 1).Value)) = True
    Next rowNumber
    For rowNumber = 1 To GetTransactionsTable.ListRows.Count
        monthName = Format(GetTransactionsTable.ListRows(rowNumber).Range.Cells(1, 2).Value, "yyyy-mm")
        If Not knownMonths.Exists(monthName) Then
            Set newMonthRow = monthlyTable.ListRows.Add
            newMonthRow.Range.Cells(1, 1).Value = monthName
            newMonthRow.Range.Cells(1, 2).Resize(1, 7).FormulaR1C1 = monthlyTable.ListRows(newMonthRow.Index - 1).Range.Cells(1, 2).Resize(1, 7).FormulaR1C1
            knownMonths(monthName) = True
        End If
    Next rowNumber
    Application.CalculateFull
    Set dashboard = ThisWorkbook.Worksheets("Dashboard")
    lastRow = Application.Max(34, dashboard.Cells(dashboard.Rows.Count, 2).End(xlUp).Row)
    dashboard.Range("B34:M" & lastRow).ClearContents
    dashboardRow = 34
    For rowNumber = 1 To GetSecuritiesTable.ListRows.Count
        unitsHeld = CDbl(GetSecuritiesTable.ListRows(rowNumber).Range.Cells(1, 12).Value)
        If unitsHeld > 0.000001 Then
            securitySheetRow = GetSecuritiesTable.ListRows(rowNumber).Range.Row
            ' Link the same investment row to each dashboard column.
            dashboard.Cells(dashboardRow, 2).Formula = "=Securities!C" & securitySheetRow
            dashboard.Cells(dashboardRow, 3).Formula = "=Securities!D" & securitySheetRow
            dashboard.Cells(dashboardRow, 4).Formula = "=Securities!E" & securitySheetRow
            dashboard.Cells(dashboardRow, 5).Formula = "=Securities!M" & securitySheetRow
            dashboard.Cells(dashboardRow, 6).Formula = "=Securities!I" & securitySheetRow
            dashboard.Cells(dashboardRow, 7).Formula = "=Securities!J" & securitySheetRow
            dashboard.Cells(dashboardRow, 8).Formula = "=Securities!N" & securitySheetRow
            dashboard.Cells(dashboardRow, 9).Formula = "=Securities!R" & securitySheetRow
            dashboard.Cells(dashboardRow, 10).Formula = "=Securities!S" & securitySheetRow
            dashboard.Cells(dashboardRow, 11).Formula = "=Securities!T" & securitySheetRow
            dashboard.Cells(dashboardRow, 12).Formula = "=Securities!G" & securitySheetRow
            dashboard.Cells(dashboardRow, 13).Formula = "=Securities!U" & securitySheetRow
            dashboard.Cells(dashboardRow, 6).NumberFormat = Chr(34) & _
                GetSecuritiesTable.ListRows(rowNumber).Range.Cells(1, 7).Value & " " & Chr(34) & "0.00"
            dashboardRow = dashboardRow + 1
        End If
    Next rowNumber
    lastRow = dashboardRow - 1
    If lastRow >= 34 Then
        Application.CalculateFull
        dashboard.Range("B33:M" & lastRow).Sort Key1:=dashboard.Range("I34"), Order1:=xlDescending, Header:=xlYes
        dashboard.Range("B34:M" & lastRow).RowHeight = 34
        dashboard.Range("C34:C" & lastRow).WrapText = True
        dashboard.Range("E34:E" & lastRow).NumberFormat = "0.0000"
        dashboard.Range("H34:J" & lastRow).NumberFormat = "#,##0.00;(#,##0.00)"
        dashboard.Range("G34:G" & lastRow).NumberFormat = "0.0000"
        dashboard.Range("K34:K" & lastRow).NumberFormat = "0.0%"
        dashboard.Range("L34:L" & lastRow).NumberFormat = "General"
        With dashboard.ChartObjects(1).Chart.SeriesCollection(1)
            .XValues = dashboard.Range("B34:B" & Application.Min(41, lastRow))
            .Values = dashboard.Range("I34:I" & Application.Min(41, lastRow))
        End With
    End If
    With ThisWorkbook.Worksheets("Monthly")
        With .ChartObjects(1).Chart.SeriesCollection(1)
            .XValues = monthlyTable.ListColumns("Month").DataBodyRange
            .Values = monthlyTable.ListColumns("ContributionsCAD").DataBodyRange
        End With
        With .ChartObjects(2).Chart.SeriesCollection(1)
            .XValues = monthlyTable.ListColumns("Month").DataBodyRange
            .Values = monthlyTable.ListColumns("DividendsCAD").DataBodyRange
        End With
    End With
    Application.CalculateFull
    Exit Sub
Failed:
    MsgBox "Recalculation stopped: " & Err.Description, vbExclamation
End Sub

' Keep one saved portfolio value for each day.
Public Sub SaveSnapshot()
    On Error GoTo Failed
    Dim snapshotSheet As Worksheet
    Dim dashboard As Worksheet
    Dim nextRow As Long
    RefreshPortfolio
    Set dashboard = ThisWorkbook.Worksheets("Dashboard")
    If Not IsNumeric(dashboard.Range("B6").Value) Then
        Err.Raise 5, , "Update prices and check any flagged records before saving a snapshot."
    End If
    On Error Resume Next
    Set snapshotSheet = ThisWorkbook.Worksheets("Snapshots")
    On Error GoTo Failed
    If snapshotSheet Is Nothing Then
        Set snapshotSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        snapshotSheet.Name = "Snapshots"
        snapshotSheet.Range("A1:E1").Value = Array("Date", "ValueCAD", "ContributionsCAD", "CashCAD", "Basis")
    End If
    If Application.CountIf(snapshotSheet.Columns(1), CLng(Date)) > 0 Then
        Err.Raise 5, , "A snapshot already exists for today."
    End If
    nextRow = snapshotSheet.Cells(snapshotSheet.Rows.Count, 1).End(xlUp).Row + 1
    snapshotSheet.Cells(nextRow, 1).Value = Date
    snapshotSheet.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd"
    snapshotSheet.Cells(nextRow, 2).Value = dashboard.Range("B6").Value
    snapshotSheet.Cells(nextRow, 3).Value = dashboard.Range("F6").Value
    snapshotSheet.Cells(nextRow, 4).Value = dashboard.Range("D10").Value
    snapshotSheet.Cells(nextRow, 5).Value = "Entered quote dates; CAD"
    snapshotSheet.Columns("A:E").AutoFit
    MsgBox "Snapshot saved."
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub

' Create the six buttons and the investment dropdown.
Public Sub InstallButtons()
    Dim entrySheet As Worksheet
    Dim newButton As Button
    Dim buttonNumber As Long
    Dim buttonNames As Variant
    Dim macroNames As Variant
    Set entrySheet = ThisWorkbook.Worksheets("Add Transaction")
    buttonNames = Array("Add transaction", "Add security", "Refresh prices", "Recalculate / charts", "Import cleaned CSV", "Save snapshot")
    macroNames = Array("AddTransaction", "AddSecurity", "RefreshPrices", "RefreshPortfolio", "ImportCleanedCSV", "SaveSnapshot")
    For buttonNumber = entrySheet.Buttons.Count To 1 Step -1
        If Left(entrySheet.Buttons(buttonNumber).Name, 3) = "pf_" Then
            entrySheet.Buttons(buttonNumber).Delete
        End If
    Next buttonNumber
    entrySheet.Range("E5:H20").UnMerge
    entrySheet.Range("E5:H20").Clear
    For buttonNumber = 0 To 5
        Set newButton = entrySheet.Buttons.Add(entrySheet.Range("E5").Left, entrySheet.Range("E5").Top + buttonNumber * 38, 170, 30)
        newButton.Name = "pf_" & macroNames(buttonNumber)
        newButton.Caption = buttonNames(buttonNumber)
        newButton.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroNames(buttonNumber)
    Next buttonNumber
    On Error Resume Next
    ThisWorkbook.Names("SecurityIDs").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="SecurityIDs", RefersTo:="=tblSecurities[SecurityID]"
    With entrySheet.Range("C6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=SecurityIDs"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    MsgBox "Six buttons and security dropdown installed. Save as .xlsm to retain VBA."
End Sub
