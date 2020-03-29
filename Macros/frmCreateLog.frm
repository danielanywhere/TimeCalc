VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateLog 
	 Caption         =   "Create Log"
	 ClientHeight    =   2385
	 ClientLeft      =   120
	 ClientTop       =   468
	 ClientWidth     =   4560
	 OleObjectBlob   =   "frmCreateLog.frx":0000
	 StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'frmCreateLog.frm
' Log creation form.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

'*------------------------------------------------------------------------*
'* cmdCancel_Click																												*
'*------------------------------------------------------------------------*
''<Description>
''Cancel button has been clicked.
''</Description>
Private Sub cmdCancel_Click()
	Unload Me
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* cmdOK_Click																														*
'*------------------------------------------------------------------------*
''<Description>
''OK button has been clicked.
''</Description>
Private Sub cmdOK_Click()
Dim bc As Boolean   'Flag - Continue.
Dim bf As Boolean   'Flag - At least one row formatted.
Dim dd As Date      'Date Due.
Dim di As Date      'Invoice Date.
Dim ic As LogCollection   'Working Log Items.
Dim id As Integer             'Detail Index.
Dim ii As LogItem         'Current Item.
Dim lp As Integer   'List Position.
Dim ni As String    'Invoice Number.
Dim rb As Integer   'Bottom Row.
Dim rc As Integer   'Row Count.
Dim rd As Integer   'Detail Row.
Dim rt As Integer   'Top Row.
Dim sc As Worksheet 'Contact Sheet.
Dim sh As Worksheet 'Current Sheet.
Dim ws As String    'Working String.

	bc = True
	bf = False
	On Error Resume Next
	Err.Clear
	If Len(txtStartDate.Text) > 0 And _
		Len(txtEndDate.Text) > 0 Then
		dd = CDate(txtStartDate.Text)
		If Err.Number <> 0 Then
			ws = "Error: Invalid Start Date..."
			bc = False
			Err.Clear
		End If
		If bc = True Then
			dd = CDate(txtEndDate.Text)
			If Err.Number <> 0 Then
				ws = "Error: Invalid End Date..."
				bc = False
				Err.Clear
			End If
		End If
		If bc = True Then
			'Values are known. Generate Log Report.
			Set sh = Sheets("LogReport")
			rt = Range("logReportHeaderRow").Row
			rb = Range("logReportFooterRow").Row
			lp = rt + 1
			Set ic = EnumerateItemsCL()
			If ic.Count >= 0 Then
				'Build the appropriate number of rows on the Log Report page.
				sh.Select
				Columns("K:K").EntireColumn.Hidden = False
				Columns("K:K").EntireColumn.ColumnWidth = _
					Range("B1").ColumnWidth + _
					Range("C1").ColumnWidth + _
					Range("D1").ColumnWidth + _
					Range("E1").ColumnWidth
				Columns("K:K").EntireColumn.Hidden = True
				
				rd = rt + 1
				rc = (rb - rt) - 1
				If ic.Count > rc Then
					'Rows need to be added.
					If rc = 1 Then
						'Inserted row will need to be formatted.
						Rows(CStr(rd) & ":" & CStr(rd)).Select
						Selection.Insert _
							Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
						FormatRowCL rd
						bf = True
						rc = rc + 1
						rb = rb + 1
					End If
					'At this time, there are at least two items to log on, and
					' at least two rows.
					Do While rc < ic.Count
						Rows(CStr(rd + 1) & ":" & CStr(rd + 1)).Select
						Selection.Insert _
							Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
						FormatRowCL rd + 1
						bf = True
						rc = rc + 1
						rb = rb + 1
					Loop
					'At this time, the number of items is equal to the number of rows.
					Range("A" & CStr(rd)).Select
				ElseIf (ic.Count = rc) Or (ic.Count = 0 And rc = 1) Then
					'Row counts are equal.
				Else
					'Rows need to be deleted.
					Do While (rc > 1 And rc > ic.Count)
						Rows(CStr(rd) & ":" & CStr(rd)).Select
						Selection.EntireRow.Delete
						rc = rc - 1
						rb = rb - 1
					Loop
					Range("A" & CStr(rd)).Select
				End If
				'At this point, the number of rows on the sheet are equal
				' to the number of items found.
				'Fill detail items.
				Me.Hide
				id = 1
				For lp = rd To rb - 1
					If ic.Count > 0 Then
						Set ii = ic.Item(id)
						Range("A" & CStr(lp)).Value = ii.ProjectName
						Range("B" & CStr(lp)).Value = ii.TaskName
						Range("K" & CStr(lp)).Value = ii.TaskName
						Range("F" & CStr(lp)).Value = Format(ii.ItemDateStart, "MM/DD/YYYY HH:nn")
						Range("G" & CStr(lp)).Value = Format(ii.ItemDateEnd, "MM/DD/YYYY HH:nn")
						Range("I" & CStr(lp)).Value = Round(ii.ManHoursSpent, 2)
						Range("J" & CStr(lp)).Formula = "=configHourlyRate*I" & CStr(lp)
						FormatRowCL lp, IIf(lp = rb - 1, True, False)
						Rows(CStr(lp) & ":" & CStr(lp)).AutoFit = True
						id = id + 1
					Else
						Range("A" & CStr(lp)).Value = ""
						Range("B" & CStr(lp)).Value = ""
						Range("K" & CStr(lp)).Value = ""
						Range("F" & CStr(lp)).Value = ""
						Range("G" & CStr(lp)).Value = ""
						Range("I" & CStr(lp)).Value = Round(0, 2)
						Range("J" & CStr(lp)).Formula = "=configHourlyRate*I" & CStr(lp)
					End If
				Next lp
				'Color with alternate background.
				Range("A" & CStr(rd) & ":J" & CStr(rb - 1)).Select
				AlternateBackColor
				Range("A" & CStr(rb)).Select
				'Set Report Filter.
				Range("logReportFilter").Value = txtFilter.Text
				'Set Report Title.
				Range("logReportTitle").Value = _
					"Log Report - " & _
					Format(txtStartDate.Text, "dddd, mmmm dd, yyyy") & " through " & _
					Format(txtEndDate.Text, "dddd, mmmm dd, yyyy")
				'Show Print Dialog.
				Application.CommandBars.ExecuteMso "PrintPreviewAndPrint"
				Unload Me
			Else
				'No items found.
				bc = False
			End If
		Else
			'Error found. Exit with message.
			MsgBox ws, vbOKOnly, "Generate Log Report"
		End If
	Else
		'Error found. Exit with message.
		ws = "Please set all required fields first."
		MsgBox ws, vbOKOnly, "Generate Log Report"
	End If
	If bc = False Then
		cmdCancel_Click
	End If

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* EnumerateItemsCL																												*
'*------------------------------------------------------------------------*
''<Description>
''Enumerate items for the log.
''</Description>
''<Returns>
''Collection of log items.
''</Returns>
Private Function EnumerateItemsCL() As LogCollection
Dim db As Date        'Begin Date.
Dim dc As Date        'Current Date.
Dim de As Date        'End Date.
Dim ii As LogItem     'Current Log Item.
Dim ic As New LogCollection
Dim lb As Integer     'List Begin.
Dim le As Integer     'List End.
Dim lp As Integer     'List position.
Dim sh As Worksheet   'Current Sheet.
Dim si As Integer     'Sheet Index.
Dim ws As String      'Working String.

	On Error Resume Next
	Err.Clear
	db = CDate(txtStartDate.Text)
	de = CDate(txtEndDate.Text)
	For si = 1 To 12
		Set sh = Sheets(PadZeros(CStr(si), 2))
		lp = 3
		Do While Len(sh.Range("D" & CStr(lp)).Value) > 0
			If Len(sh.Range("A" & CStr(lp)).Value) > 0 And _
				sh.Range("A" & CStr(lp)).Value <> 0 And _
				sh.Range("B" & CStr(lp)).Value <> 1 And _
				Len(sh.Range("E" & CStr(lp)).Value) > 0 And _
				Len(sh.Range("F" & CStr(lp)).Value) > 0 Then
				'The item is active and has not been sent.
				dc = CDate(sh.Range("E" & CStr(lp)).Value)
				If dc >= db And dc < de + 1 Then
					Set ii = New LogItem
					ii.SheetName = sh.Name
					ii.RowIndex = lp
					ii.ItemDateStart = Format(sh.Range("E" & CStr(lp)).Value, "MM/DD/YYYY HH:nn")
					ii.ItemDateEnd = Format(sh.Range("F" & CStr(lp)).Value, "MM/DD/YYYY HH:nn")
					ii.ManHoursSpent = sh.Range("H" & CStr(lp)).Value
					ii.ProjectName = sh.Range("C" & CStr(lp)).Value
					ii.TaskName = sh.Range("D" & CStr(lp)).Value
					If Len(txtFilter.Text) = 0 Or _
						InStr(LCase(ii.ProjectName), LCase(txtFilter.Text)) > 0 Or _
						InStr(LCase(ii.TaskName), LCase(txtFilter.Text)) > 0 Then
						ic.Add ii
					End If
				End If
			End If
			lp = lp + 1
		Loop
	Next si
	Set EnumerateItemsCL = ic

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'*	FormatRowCL																														*
'*------------------------------------------------------------------------*
''<Description>
''Format the row for output to the log.
''</Description>
''<Param name="RowIndex">
''1-based sheet row index to format.
''</Param>
''<Param name="IsBottom">
''Value indicating whether the row being formatted in the bottom row in
'' the set.
''</Param>
Private Sub FormatRowCL(RowIndex As Integer, _
	Optional IsBottom As Boolean = False)
Dim rd As Integer   'Detail Row.

	On Error Resume Next
	Err.Clear
	rd = RowIndex
	Rows(CStr(rd) & ":" & CStr(rd)).Select
	Selection.Font.Bold = False
	Selection.Borders(xlDiagonalDown).LineStyle = xlNone
	Selection.Borders(xlDiagonalUp).LineStyle = xlNone
	Selection.Borders(xlEdgeLeft).LineStyle = xlNone
	Selection.Borders(xlEdgeRight).LineStyle = xlNone
	Selection.Borders(xlInsideVertical).LineStyle = xlNone
	Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
	If Not IsBottom Then
		Selection.Borders(xlEdgeBottom).LineStyle = xlNone
	End If
	Range("A" & CStr(rd) & ":J" & CStr(rd)).Select
	With Selection
		.HorizontalAlignment = xlGeneral
		.VerticalAlignment = xlTop
	End With
	If IsBottom Then
		With Selection.Borders(xlEdgeBottom)
			.LineStyle = xlContinuous
			.ColorIndex = xlAutomatic
			.TintAndShade = 0
			.Weight = xlMedium
		End With
	End If
	Range("K" & CStr(rd)).Select
	With Selection
		.WrapText = True
	End With
	DoEvents
	Range("B" & CStr(rd) & ":E" & CStr(rd)).Select
	With Selection
		.WrapText = True
		.Orientation = 0
		.AddIndent = False
		.IndentLevel = 0
		.ShrinkToFit = False
		.ReadingOrder = xlContext
		.MergeCells = True
	End With
	DoEvents
	Range("A" & CStr(rd) & ":E" & CStr(rd)).Select
	With Selection
		.HorizontalAlignment = xlLeft
	End With
	If Err.Number <> 0 Then
		MsgBox "Error: " & Err.Description
		Err.Clear
	End If
	DoEvents
	Range("F" & CStr(rd) & ":G" & CStr(rd)).Select
	With Selection
		.HorizontalAlignment = xlRight
		.NumberFormat = "mm/dd/yyyy hh:mm"
	End With
	DoEvents
	Range("H" & CStr(rd) & ":J" & CStr(rd)).Select
	With Selection
		.HorizontalAlignment = xlRight
		.NumberFormat = "#,##0.00"
	End With
	DoEvents

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UserForm_Initialize																										*
'*------------------------------------------------------------------------*
''<Description>
''The user form is initializing.
''</Description>
Private Sub UserForm_Initialize()
	txtEndDate.Text = Format(Date, "MM/DD/YYYY")
	txtStartDate.Text = _
		Format(DateAdd("d", -6, txtEndDate.Text), "MM/DD/YYYY")
	txtFilter.Text = Range("logReportFilter").Value
End Sub
'*------------------------------------------------------------------------*
