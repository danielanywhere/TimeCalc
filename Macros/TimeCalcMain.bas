Attribute VB_Name = "TimeCalcMain"
'TimeCalcMain.bas
' Common functionality for all sheets and forms of the workbook.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

'Month columns.
' 0 - Active.
Public Const colActive = 0
' 1 - Sent.
Public Const colSent = 1
' 2 - Service.
Public Const colService = 2
' 3 - Project.
Public Const colProject = 3
' 4 - Task.
Public Const colTask = 4
' 5 - Start.
Public Const colStart = 5
' 6 - End.
Public Const colEnd = 6
' 7 - MH.
Public Const colMH = 7
' 8 - Billable.
Public Const colBillable = 8
' 9 - Charge.
Public Const colCharge = 9
' 10 - Invoiced.
Public Const colInvoiced = 10
' 11 - Received.
Public Const colReceived = 11
' 12 - Due.
Public Const colDue = 12

'*------------------------------------------------------------------------*
'* AddOutlookInvoiceNotification																					*
'*------------------------------------------------------------------------*
''<Description>
''Create an Outlook notification for the specified invoice and due date.
''</Description>
''<Param name="InvoiceNumber">
''The number of the invoice for which a reminder will be set.
''</Param>
''<Param name="DueDate">
''The date upon which the reminder will occur.
''</Param>
Public Sub AddOutlookInvoiceNotification( _
	InvoiceNumber As String, DueDate As Date)
Dim oa As Object    'Outlook application.
Dim oi As Object    'Outlook item.

	Set oa = CreateObject("Outlook.Application")
	Set oi = oa.CreateItem(1)
	oi.Subject = "TimeCalc Invoice " & InvoiceNumber & " Due"
	oi.Location = "Accounting"
	oi.Start = Format(DueDate, "mm/dd/yyyy")
	'1 minute.
	oi.Duration = 1
	'olFree = 0
	oi.BusyStatus = 0
	oi.ReminderSet = True
	oi.ReminderMinutesBeforeStart = 0
	oi.Body = "TimeCalc invoice #" & InvoiceNumber & " is now due."
	oi.Save
	Set oi = Nothing
	Set oa = Nothing

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* AddService																															*
'*------------------------------------------------------------------------*
''<Description>
''Add a new service to the Services table.
''</Description>
''<Param name="ContactCode">
''The code of the contact to which the service will be associated.
''</Param>
''<Param name="ServiceName">
''Name of the service to add.
''</Param>
''<Param name="Rate">
''The hourly rate for this contact / service.
''</Param>
''<Param name="Commission">
''The percentage of commission to be paid out for this contact / service.
''</Param>
Public Sub AddService(ContactCode As String, ServiceName As String, _
	Rate As Double, Commission As Double)
Dim cc As Integer     'Column count.
Dim ciCommission As Integer   'Column indexes.
Dim ciCustomer As Integer
Dim ciRate As Integer
Dim ciService As Integer
Dim cp As Integer     'Column position.
Dim lp As Integer     'List position.
Dim rc As Integer     'Row count.
Dim rg As Range       'Current range.
Dim sh As Worksheet   'Current sheet.
Dim tb As ListObject  'Current table.
Dim ws As String      'Working string.

	Set sh = Sheets("Services")
	Set tb = sh.ListObjects("Services")

	'Get the list of columns.
	Set rg = sh.ListObjects("Services").HeaderRowRange
	cc = rg.Columns.Count
	If cc > 0 Then
		'Columns are found.
		For cp = 1 To cc
			ws = rg.Cells(1, cp).Value
			Select Case ws
				Case "Commission":
					ciCommission = rg.Cells(1, cp).Column
				Case "Customer":
					ciCustomer = rg.Cells(1, cp).Column
				Case "Rate per hr":
					ciRate = rg.Cells(1, cp).Column
				Case "Service":
					ciService = rg.Cells(1, cp).Column
			End Select
		Next cp
	End If

	rc = tb.ListRows.Count
	With tb.DataBodyRange
		If ciCustomer > 0 And ciService > 0 Then
			If Len(.Cells(rc, ciCustomer).Value) > 0 Or _
				Len(.Cells(rc, ciService).Value) > 0 Then
				'Add a new row if blank isn't present.
				tb.ListRows.Add
				rc = tb.ListRows.Count
			End If
			.Cells(rc, ciCustomer).Value = ContactCode
			.Cells(rc, ciService).Value = ServiceName
			If ciRate > 0 Then
				.Cells(rc, ciRate).Value = Rate
			End If
			If ciCommission > 0 Then
				.Cells(rc, ciCommission).Value = Commission
			End If
		End If
	End With

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* AddUniqueSorted																												*
'*------------------------------------------------------------------------*
''<Description>
''Add an item to the raw strings collection, in a sorted position, but only
'' if it is unique.
''</Description>
''<Param name="Items">
''Raw collection of string values.
''</Param>
''<Param name="Value">
''The value to add if unique.
''</Param>
Public Sub AddUniqueSorted(Items As Collection, Value As String)
Dim bf As Boolean 'Flag - found.
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.
Dim ls As String  'Lower case search value.
Dim ws As String  'Working string.

	ls = LCase(Value)
	lc = Items.Count()
	For lp = 1 To lc
		ws = LCase(Items.Item(lp))
		If ws = ls Then
			'Skip non-unique.
			bf = True
			Exit For
		ElseIf ws > ls Then
			'Add sorted.
			Items.Add Value, Before:=lp
			bf = True
			Exit For
		End If
	Next lp
	If bf = False Then
		Items.Add Value
	End If

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* AlternateBackColor																											*
'*------------------------------------------------------------------------*
''<Description>
''Set the selected area to rows of alternating background color.
''</Description>
Public Sub AlternateBackColor()
Dim cb As String    'Column Begin.
Dim cc As Long      'Column Count.
Dim ce As String    'Column End.
Dim cp As Long      'Column Position.
Dim re As Integer   'Row End.
Dim ri As Integer   'Row Index.
Dim rng As Range
Dim rp As Integer   'Row Position.

	ri = 1
	Set rng = Selection
	cp = rng.Column
	cc = rng.Columns.Count
	cb = ColToChar(cp)
	ce = ColToChar(cp + cc - 1)
	
	re = rng.Row + rng.Rows.Count - 1
	For rp = rng.Row To re
		Range(cb & CStr(rp) & ":" & ce & CStr(rp)).Select
		If ri Mod 2 = 0 Then
			'Shade.
			With Selection.Interior
				.Pattern = xlSolid
				.PatternColorIndex = xlAutomatic
				.ThemeColor = xlThemeColorAccent1
				.TintAndShade = 0.799981688894314
				.PatternTintAndShade = 0
			End With
		Else
			'Normal.
			With Selection.Interior
				.Pattern = xlNone
				.TintAndShade = 0
				.PatternTintAndShade = 0
			End With
		End If
		ri = ri + 1
	Next rp

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* CellHasValidation																											*
'*------------------------------------------------------------------------*
''<Description>
''Return a value indicating whether the specified cell has data validation.
''</Description>
''<Param name="Cell">
''Range representing the cell to test.
''</Param>
''<Returns>
''True if the cell has data validation activated. Otherwise, false.
''</Returns>
Public Function CellHasValidation(Cell As Range) As Boolean
Dim rv As Boolean   'Return value.
Dim tp As Variant   'Validation type.

	Err.Clear
	On Local Error Resume Next
	tp = Cell.Validation.Type
	
	If Err.Number = 0 Then
		rv = True
	Else
		Err.Clear
	End If

	CellHasValidation = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ClearStatusBar																													*
'*------------------------------------------------------------------------*
''<Description>
''Clear the status bar on this application.
''</Description>
Public Sub ClearStatusBar()
	Application.StatusBar = False
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ConvertTime																														*
'*------------------------------------------------------------------------*
''<Description>
''Convert from one time domain to another.
''</Description>
''<Param name="FromUnit">
''The source unit from which to convert.
'' Available sources are:
'' sec, min, hr, day, dy, week, wk, month, mo, year, yr.
''</Param>
''<Param name="ToUnit">
''The target unit to which the value will be converted.
'' Available targets are:
'' sec, min, hr, day, dy, week, wk, month, mo, year, yr.
''</Param>
Public Function ConvertTime(Value As Double, _
	FromUnit As String, ToUnit As String) As Double
Dim lf As String    'Lower From Unit.
Dim lt As String    'Lower To Unit.
Dim rv As Double    'Return Value.

	lf = LCase(FromUnit)
	lt = LCase(ToUnit)
	Select Case lf
		Case "sec":
			Select Case lt
				Case "sec":
					rv = Value * 1
				Case "min":
					rv = Value * 0.016666667
				Case "hr":
					rv = Value * 0.000277778
				Case "day", "dy":
					rv = Value * 0.000011574
				Case "week", "wk":
					rv = Value * 0.000001653
				Case "month", "mo":
					rv = Value * 0.000000382
				Case "year", "yr":
					rv = Value * 0.000000032
			End Select
		Case "min":
			Select Case lt
				Case "sec":
					rv = Value * 60
				Case "min":
					rv = Value * 1
				Case "hr":
					rv = Value * 0.016666667
				Case "day", "dy":
					rv = Value * 0.000694444
				Case "week", "wk":
					rv = Value * 0.000099206
				Case "month", "mo":
					rv = Value * 0.000022894
				Case "year", "yr":
					rv = Value * 0.000001908
			End Select
		Case "hr":
			Select Case lt
				Case "sec":
					rv = Value * 3600
				Case "min":
					rv = Value * 60
				Case "hr":
					rv = Value * 1
				Case "day", "dy":
					rv = Value * 0.041666667
				Case "week", "wk":
					rv = Value * 0.005952381
				Case "month", "mo":
					rv = Value * 0.001373626
				Case "year", "yr":
					rv = Value * 0.000114469
			End Select
		Case "day", "dy":
			Select Case lt
				Case "sec":
					rv = Value * 86400
				Case "min":
					rv = Value * 1440
				Case "hr":
					rv = Value * 24
				Case "day", "dy":
					rv = Value * 1
				Case "week", "wk":
					rv = Value * 0.142857143
				Case "month", "mo":
					rv = Value * 0.032967033
				Case "year", "yr":
					rv = Value * 0.002747253
			End Select
		Case "week", "wk":
			Select Case lt
				Case "sec":
					rv = Value * 604800
				Case "min":
					rv = Value * 10080
				Case "hr":
					rv = Value * 168
				Case "day", "dy":
					rv = Value * 7
				Case "week", "wk":
					rv = Value * 1
				Case "month", "mo":
					rv = Value * 0.230769231
				Case "year", "yr":
					rv = Value * 0.019230769
			End Select
		Case "month", "mo":
			Select Case lt
				Case "sec":
					rv = Value * 2620800
				Case "min":
					rv = Value * 43680
				Case "hr":
					rv = Value * 728
				Case "day", "dy":
					rv = Value * 30.333333333
				Case "week", "wk":
					rv = Value * 4.333333333
				Case "month", "mo":
					rv = Value * 1
				Case "year", "yr":
					rv = Value * 0.083333333
			End Select
		Case "year", "yr":
			Select Case lt
				Case "sec":
					rv = Value * 31449600
				Case "min":
					rv = Value * 524160
				Case "hr":
					rv = Value * 8736
				Case "day", "dy":
					rv = Value * 364
				Case "week", "wk":
					rv = Value * 52
				Case "month", "mo":
					rv = Value * 12
				Case "year", "yr":
					rv = Value * 1
			End Select
	End Select
	ConvertTime = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* CreateYear																															*
'*------------------------------------------------------------------------*
''<Description>
''Clear the entire year and configure new sheets.
''</Description>
Public Sub CreateYear()
Dim bct As Boolean
Dim da As Integer   'Current Day.
Dim dt As Date      'Current Date.
Dim lp As Integer   'List Position.
Dim mo As Integer   'Current Month.
Dim sp As Integer   'Sheet Position.
Dim st As Integer   'Starting Day.
Dim yr As Integer   'Current Year.

	bct = (MsgBox( _
		"This function will clear all sheets in this workbook." & _
		vbCrLf & _
		"Do you wish to continue?", vbYesNo, "Create Year") = vbYes)
	If bct = True Then
		'User will continue.
		yr = CLng(InputBox("Current Year:", "Create Year", CStr(Year(Now()))))
		'Get first Monday of year.
		mo = 1  'January 1.
		da = 1
		dt = CDate(CStr(mo) & "/" & CStr(da) & "/" & CStr(yr))
		st = Weekday(dt)
		Do While st <> vbMonday
			dt = dt + 1
			st = Weekday(dt)
		Loop
		If dt > 3 Then dt = dt - 7

		Sheets("Timesheet").Select
		lp = 2
		Do While Len(Range("B" & CStr(lp)).Value) > 0
			Range("A" & CStr(lp)).FormulaR1C1 = Format(dt, "MM/DD/YYYY")
			Range("C" & CStr(lp) & ":I" & CStr(lp + 3)).ClearContents
			lp = lp + 7
			dt = dt + 7
		Loop

		'Monthly Time Sheets.
		For sp = 1 To 12
			Sheets(PadZeros(CStr(sp), 2)).Select
			Do While Len(Range("A3").Value) > 0 Or _
				Len(Range("B3").Value) > 0 Or _
				Len(Range("C3").Value) > 0 Or _
				Len(Range("D3").Value) > 0 Or _
				Len(Range("E3").Value) > 0 Or _
				Len(Range("F3").Value) > 0 Or _
				Len(Range("G3").Value) > 0 Or _
				Len(Range("H3").Value) > 0 Or _
				Len(Range("I3").Value) > 0 Or _
				Len(Range("J3").Value) > 0 Or _
				Len(Range("K3").Value) > 0 Or _
				Len(Range("L3").Value) > 0 Or _
				Len(Range("M3").Value) > 0
				Range("A3").EntireRow.Delete
			Loop
		Next sp

		Sheets("Timesheet").Select

	End If
		
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GenerateInvoice																												*
'*------------------------------------------------------------------------*
''<Description>
''Load the invoice creation form and allow generation of the invoice.
''</Description>
Public Sub GenerateInvoice()
	frmCreateInvoice.Show
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GenerateLog																														*
'*------------------------------------------------------------------------*
''<Description>
''Load the log creation form and allow generation of the log report.
''</Description>
Public Sub GenerateLog()
	frmCreateLog.Show
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetClockTime																														*
'*------------------------------------------------------------------------*
''<Description>
''Return the time of day from any date.
''</Description>
''<Param name="Value">
''The date value to inspect.
''</Param>
''<Returns>
''The time of day, in hours and minutes, formatted as a binary date.
''</Returns>
Public Function GetClockTime(Value As Double) As Double
Dim hv As Double
Dim mv As Double
Dim rv As Double
Dim vl As Double

	vl = Value
	Do While vl > 24
		vl = vl - 24
	Loop
	hv = Fix(vl)
	mv = vl - hv
	mv = Fix(mv * 60)
	rv = CVDate(PadZeros(CStr(hv), 2) & ":" & _
		PadZeros(CStr(mv), 2) & IIf(hv = "12", " PM", ""))
	GetClockTime = rv
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetColumnName																													*
'*------------------------------------------------------------------------*
''<Description>
''Return the letter name of the column specified by the header row title.
''</Description>
''<Param name="Sheet">
''Reference to the worksheet to inspect.
''</Param>
''<Param name="Row">
''Index of the header row on the sheet.
''</Param>
''<Returns>
''Letter name of the specified column.
''</Returns>
Public Function GetColumnName(Sheet As Worksheet, Row As Integer, _
	ColumnTitle As String) As String
Dim cp As Integer 'Column position.
Dim cv As String  'Character position.
Dim rv As String  'Return value.
Dim tl As String  'Lower case column title.
Dim ws As String  'Working string.

	tl = LCase(ColumnTitle)
	cp = 1
	cv = ColToChar(cp)
	ws = LCase(Sheet.Range(cv & CStr(Row)).Value)
	Do While Len(ws) > 0
		If ws = tl Then
			rv = cv
			Exit Do
		End If
		cp = cp + 1
		cv = ColToChar(cp)
		ws = LCase(Sheet.Range(cv & CStr(Row)).Value)
	Loop
	GetColumnName = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetColumnNames																													*
'*------------------------------------------------------------------------*
''<Description>
''Return the month-sheet specific array of letter names occupied by the
'' common column titles.
''</Description>
''<Param name="Sheet">
''Reference to the worksheet to inspect.
''</Param>
''<Returns>
''Array of strings, in the order of the colX constants at the top of this
'' module.
''</Returns>
Public Function GetColumnNames(Sheet As Worksheet) As String()
Dim cp As Integer   'Column position.
Dim cv As String    'Character position.
Dim lp As Integer   'List position.
Dim rv() As String  'Return value.

	'Column order.
	'0  - Active.
	'1  - Sent.
	'2  - Service.
	'3  - Project.
	'4  - Task.
	'5  - Start.
	'6  - End.
	'7  - MH.
	'8  - Billable.
	'9  - Charge.
	'10 - Invoiced.
	'11 - Received.
	'12 - Due.
	ReDim rv(12)
	lp = 2
	cp = 1
	cv = ColToChar(cp)
	ws = LCase(Sheet.Range(cv & CStr(lp)).Value)
	Do While Len(ws) > 0
		Select Case ws
			Case "active":
				rv(colActive) = cv
			Case "sent":
				rv(colSent) = cv
			Case "service":
				rv(colService) = cv
			Case "project":
				rv(colProject) = cv
			Case "task":
				rv(colTask) = cv
			Case "start":
				rv(colStart) = cv
			Case "end":
				rv(colEnd) = cv
			Case "mh":
				rv(colMH) = cv
			Case "billable":
				rv(colBillable) = cv
			Case "charge":
				rv(colCharge) = cv
			Case "invoiced":
				rv(colInvoiced) = cv
			Case "received":
				rv(colReceived) = cv
			Case "due":
				rv(colDue) = cv
		End Select
		cp = cp + 1
		cv = ColToChar(cp)
		ws = LCase(Sheet.Range(cv & CStr(lp)).Value)
	Loop
	GetColumnNames = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetColumnNamesEx																												*
'*------------------------------------------------------------------------*
''<Description>
''Return a 2-dimensional 1-based array of column names for the specified
'' header row of the caller's sheet.
''</Description>
''<Param name="Sheet">
''Reference to the worksheet to inspect.
''</Param>
''<Param name="RowIndex">
''Index of the header row on the sheet.
''</Param>
''<Returns>
''2-dimensional, 1-based string array with the syntax
'' cn(ColumnName, ColumnLetter).
''</Returns>
Public Function GetColumnNamesEx(Sheet As Worksheet, _
	RowIndex As Integer) As String()
Dim cc As Integer   'Column count.
Dim cn() As String  'Column names.
Dim co As New Collection
Dim cp As Integer   'Column position.
Dim ws As String    'Working string.

	ReDim cn(0, 0)
	cp = 1
	ws = Sheet.Range(ColToChar(cp) & CStr(RowIndex)).Value
	Do While Len(ws) > 0
		co.Add ws
		cp = cp + 1
		ws = Sheet.Range(ColToChar(cp) & CStr(RowIndex)).Value
	Loop
	cc = co.Count()
	If cc > 0 Then
		ReDim cn(cc, 1)
		For cp = 1 To cc
			cn(cp, 0) = co.Item(cp)
			cn(cp, 1) = ColToChar(cp)
		Next cp
	End If
	GetColumnNamesEx = cn

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetConfigRow																														*
'*------------------------------------------------------------------------*
''<Description>
''Return the row index of the specified configuration value.
''</Description>
''<Param name="Name">
''The name of the configuration variable to find.
''</Param>
''<Returns>
''Row index of the configuration value on the Config sheet.
''</Returns>
Public Function GetConfigRow(Name As String) As Integer
Dim lp As Integer   'List Position.
Dim ls As String    'Line String.
Dim rv As Integer   'Return Value.
Dim sh As Worksheet 'Local Worksheet.

	rv = 0
	ls = Replace(LCase(Name), " ", "")
	Set sh = Sheets("Config")
	lp = 2
	Do While Len(sh.Range("A" & CStr(lp)).Value) > 0
		If ls = Replace(LCase( _
			sh.Range("A" & CStr(lp)).Value), " ", "") Then
			rv = lp
			Exit Do
		End If
		lp = lp + 1
	Loop
	GetConfigRow = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetConfigValue																													*
'*------------------------------------------------------------------------*
''<Description>
''Return the value found a the specified configuration line.
''</Description>
''<Param name="Name">
''Name of the configuration item to look up.
''</Param>
''<Returns>
''The value of the specified configuration item.
''</Returns>
Public Function GetConfigValue(Name As String) As String
Dim lp As Integer   'List Position.
Dim ls As String    'Line String.
Dim rv As String    'Return Value.
Dim sh As Worksheet 'Local Worksheet.

	ls = Replace(LCase(Name), " ", "")
	Set sh = Sheets("Config")
	lp = 2
	Do While Len(sh.Range("A" & CStr(lp)).Value) > 0
		If ls = Replace(LCase( _
			sh.Range("A" & CStr(lp)).Value), " ", "") Then
			rv = sh.Range("B" & CStr(lp)).Value
			Exit Do
		End If
		lp = lp + 1
	Loop
	GetConfigValue = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetDateDouble																													*
'*------------------------------------------------------------------------*
''<Description>
''Return the binary double version of the date.
''</Description>
''<Param name="Value">
''Date to convert.
''</Param>
''<Returns>
''Double precision floating point representation of the caller's date.
''</Returns>
Public Function GetDateDouble(Value As Date) As Double
	GetDateDouble = CVDate(Value)
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetDecimalTime																													*
'*------------------------------------------------------------------------*
''<Description>
''Return the decimal time from an hourly time-formatted value.
''</Description>
''<Param name="Value">
''The value to convert.
''</Param>
''<Returns>
''Decimal time.
''</Returns>
Public Function GetDecimalTime(Value As String) As Double
Dim hr As String
Dim hv As Integer
Dim mn As String
Dim mv As Double
Dim rv As Double
Dim wv As Double
Dim va() As String

	If Len(Value) = 2 Then
		Value = Value & ":00"
	End If

	On Error Resume Next
	wv = CVDate(CDbl(Value))
	If Err.Number <> 0 Then
		wv = CVDate(Value)
		Err.Clear
	End If
	On Error GoTo 0

	va = Split(Format(wv, "HH:nn"), ":")
	If UBound(va) > 0 Then
		hr = va(0)
	End If
	If UBound(va) >= 1 Then
		mn = va(1)
	End If
	hv = 0
	mv = 0
	On Error Resume Next
	hv = CInt(hr)
	mv = CInt(mn)
	On Error GoTo 0
	mv = mv / 60
	rv = hv + (mv)
	GetDecimalTime = rv
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetHours																																*
'*------------------------------------------------------------------------*
''<Description>
''Return the number of decimal hours between the start and end dates.
''</Description>
''<Param name="StartDate">
''The starting date and time.
''</Param>
''<Param name="EndDate">
''The ending date and time.
''</Param>
''<Returns>
''Time elapsed, in decimal format.
''</Returns>
Public Function GetHours(StartDate As Variant, _
	EndDate As Variant) As Single
Dim de As Double
Dim ds As Double
Dim rv As Long

	On Error Resume Next
	Err.Clear
	de = CVDate(EndDate)
	If Err.Number <> 0 Then
		de = 0
		Err.Clear
	End If
	ds = CVDate(StartDate)
	If Err.Number <> 0 Then
		ds = 0
		Err.Clear
	End If
	If de = 0 Then de = Now()
	If ds = 0 Then ds = Now()

	rv = DateDiff("n", ds, de)
	GetHours = CSng(Format(rv / 60, "0.00"))

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetLastColumnIndex																											*
'*------------------------------------------------------------------------*
''<Description>
''Return the last column index in range.
''</Description>
''<Param name="Data">
''The range to check.
''</Param>
''<Returns>
''The 1-based index of the last column in the specified range.
''</Returns>
Public Function GetLastColumnIndex(Data As Range) As Integer
Dim rv As Integer 'Return value.

	rv = Data.Columns.Item(Data.Columns.Count).Column
	GetLastColumnIndex = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetLastRowIndex																												*
'*------------------------------------------------------------------------*
''<Description>
''Return the last row index in the range.
''</Description>
''<Param name="Data">
''Range to inspect.
''</Param>
''<Returns>
''The 1-based index of the last row in the range.
''</Returns>
Public Function GetLastRowIndex(Data As Range) As Integer
Dim rv As Integer   'Return value.

	rv = Data.Rows.Item(Data.Rows.Count).Row
	GetLastRowIndex = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetRangeLookup																													*
'*------------------------------------------------------------------------*
''<Description>
''Return the value looked up from within a range.
''</Description>
''<Param name="Data">
''The range to search.
''</Param>
''<Param name="KeyColumn">
''The column letter containing the key value.
''</Param>
''<Param name="KeyValue">
''The value to match in the key column.
''</Param>
''<Param name="DataColumn">
''The column letter containing the data to return.
''</Param>
''<Returns>
''Value found in the DataColumn cell for the matched key.
''</Returns>
Public Function GetRangeLookup(Data As Range, _
	KeyColumn As String, KeyValue As String, DataColumn As String) As String
Dim cl As Integer   'Last column index.
Dim dc As Integer   'Data column.
Dim kc As Integer   'Key column.
Dim lb As Integer   'List begin.
Dim le As Integer   'List end.
Dim lp As Integer   'List position.
Dim rl As Integer   'Last row index.
Dim rv As String    'Return value.
Dim sh As Worksheet 'Target sheet.
Dim ws As String    'Working string.

	cl = GetLastColumnIndex(Data)
	rl = GetLastRowIndex(Data)
	kc = CharToCol(KeyColumn)
	dc = CharToCol(DataColumn)
	If kc >= Data.Column And kc <= cl And _
		dc >= Data.Column And dc <= cl Then
		Set sh = Data.Worksheet
		'Both key and data columns are present within the range.
		lb = Data.Row
		le = rl
		For lp = lb To le
			ws = sh.Range(KeyColumn & CStr(lp)).Value
			If ws = KeyValue Then
				'Lookup was found.
				rv = sh.Range(DataColumn & CStr(lp)).Value
				Exit For
			End If
		Next lp
	End If
	GetRangeLookup = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* MarkItemsSent																													*
'*------------------------------------------------------------------------*
''<Description>
''On the physical sheets, mark all items from the collection as sent.
''</Description>
''<Param name="Items">
''Collection of items to find and mark sent.
''</Param>
Public Sub MarkItemsSent(Items As InvoiceCollection)
Dim cn() As String  'Column names.
Dim ic As Integer   'Item count.
Dim ii As InvoiceItem
Dim ip As Integer   'Item position.
Dim sh As Worksheet 'Active sheet.
Dim sn As String    'Current sheet name.

	sn = ""
	ic = Items.Count()
	For ip = 1 To ic
		Set ii = Items.Item(ip)
		If sn <> ii.SheetName Then
			'Select the sheet.
			Set sh = Sheets(ii.SheetName)
			cn = GetColumnNames(sh)
			sn = ii.SheetName
		End If
		'Set the sent and invoiced values.
		sh.Range(cn(colSent) & CStr(ii.RowIndex)).Value = 1
		sh.Range(cn(colInvoiced) & CStr(ii.RowIndex)).Value = ii.InvoicedAmount
	Next ip
	
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* MonthSheetsRenameService																								*
'*------------------------------------------------------------------------*
''<Description>
''Rename matching spreadsheet service selection items from the collection.
''</Description>
''<Param name="Items">
''Collection of items to find and rename.
''</Param>
''<Param name="FromContactCode">
''Original contact code to match.
''</Param>
''<Param name="FromServiceName">
''Original service name to match.
''</Param>
''<Param name="ToContactCode">
''New contact code to set.
''</Param>
''<Param name="ToServiceName">
''New service name to match.
''</Param>
Public Sub MonthSheetsRenameService(Items As InvoiceCollection, _
 FromContactCode As String, FromServiceName As String, _
 ToContactCode As String, ToServiceName As String)
Dim cn() As String  'Column names.
Dim fs As String    'From string.
Dim ic As Integer   'Item count.
Dim ii As InvoiceItem
Dim ip As Integer   'Item position.
Dim sh As Worksheet 'Current sheet.
Dim sn As String    'Current sheet name.
Dim ts As String    'To string.

	sn = ""
	ic = Items.Count()
	For ip = 1 To ic
		Set ii = Items.Item(ip)
		If ii.ContactCode = FromContactCode And _
			ii.ServiceName = FromServiceName And _
			Len(ii.SheetName) > 0 And _
			ii.RowIndex > 2 Then
			'Matching cell found in this sheet.
			If sn <> ii.SheetName Then
				'Select the sheet.
				Set sh = Sheets(ii.SheetName)
				cn = GetColumnNames(sh)
				sn = ii.SheetName
			End If
			'Rename the service cell.
			fs = FromContactCode & "_" & FromServiceName
			ts = ToContactCode & "_" & ToServiceName
			If sh.Range(cn(colService) & CStr(ii.RowIndex)).Value = fs Then
				sh.Range(cn(colService) & CStr(ii.RowIndex)).Value = ts
			End If
		End If
	Next ip

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* PadZeros																																*
'*------------------------------------------------------------------------*
''<Description>
''Left pad the caller's value with zeros.
''</Description>
''<Param name="Value">
''The value to inspect.
''</Param>
''<Param name="Length">
''The desired minimum length of the value.
''</Param>
''<Returns>
''Caller's value, left-padded with zeros as appropriate.
''</Returns>
Public Function PadZeros(Value As String, Length As Integer) As String
Dim rs As String

	rs = Value
	Do While Len(rs) < Length
		rs = "0" & rs
	Loop
	PadZeros = rs
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ReceivePayment																													*
'*------------------------------------------------------------------------*
''<Description>
''Display the payment receipt form, allowing the user to receive payments.
''</Description>
Public Sub ReceivePayment()
	frmReceivePayment.Show
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* SetConfigValue																													*
'*------------------------------------------------------------------------*
''<Description>
''Set the value of the specified configuration item.
''</Description>
''<Param name="Name">
''Name of the configuration item to look up.
''</Param>
''<Param name="Value">
''Value to place in the specified item.
''</Param>
''<Returns>
''Response message indicating status of operation. If OK, the operation
'' was successful.
''</Returns>
Public Function SetConfigValue(Name As String, Value As String) As String
Dim lp As Integer   'List Position.
Dim ls As String    'Line String.
Dim rv As String    'Return Value.
Dim sh As Worksheet 'Local Worksheet.

	rv = "Not Found"
	ls = Replace(LCase(Name), " ", "")
	Set sh = Sheets("Config")
	lp = 2
	Do While Len(sh.Range("A" & CStr(lp)).Value) > 0
		If ls = Replace(LCase( _
			sh.Range("A" & CStr(lp)).Value), " ", "") Then
			sh.Range("B" & CStr(lp)).Value = Value
			rv = "OK"
			Exit Do
		End If
		lp = lp + 1
	Loop
	SetConfigValue = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* SetReceivedAmounts																											*
'*------------------------------------------------------------------------*
''<Description>
''Set the received amounts for all items in the provided collection.
''</Description>
''<Param name="Items">
''Collection of items where the received amount has potentially been
'' changed.
''</Param>
Public Sub SetReceivedAmounts(Items As InvoiceCollection)
Dim cn() As String  'Column names.
Dim ic As Integer   'Item count.
Dim ii As InvoiceItem
Dim ip As Integer   'Item position.
Dim sh As Worksheet 'Active sheet.
Dim sn As String    'Current sheet name.

	sn = ""
	ic = Items.Count()
	For ip = 1 To ic
		Set ii = Items.Item(ip)
		If sn <> ii.SheetName Then
			'Select the sheet.
			Set sh = Sheets(ii.SheetName)
			cn = GetColumnNames(sh)
			sn = ii.SheetName
		End If
		'Set the received amount.
		If ii.ReceivedAmount = 0# Then
			'0 is blank on this column.
			sh.Range(cn(colReceived) & CStr(ii.RowIndex)).Value = ""
		Else
			sh.Range(cn(colReceived) & CStr(ii.RowIndex)).Value = _
				ii.ReceivedAmount
		End If
	Next ip
	
End Sub
'*------------------------------------------------------------------------*
