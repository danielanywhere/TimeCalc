VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'ThisWorkbook.cls
' Handling for events and functionality that occur on every sheet.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mColumnNames() As String
Private mSheetName As String
Private mUpdating As Boolean

'*------------------------------------------------------------------------*
'* SetCharge																															*
'*------------------------------------------------------------------------*
''<Description>
''Set the active charge calculation for the specified row, based upon
'' selected service.
''</Description>
''<Param name="RowIndex">
''1-based index of the row for which the charge will be calculated.
''</Param>
Private Sub SetCharge(RowIndex As Integer)
Dim shServices As Worksheet
Dim vl As String    'Working value.

	On Local Error Resume Next
	If Len(Range(mColumnNames(colService) & CStr(RowIndex)).Value) = 0 Then
		'No service selected.
		Range(mColumnNames(colCharge) & CStr(RowIndex)).Formula = _
			"=(" & mColumnNames(colBillable) & CStr(RowIndex) & "*configHourlyRate)"
	Else
		'Charge according to the selected service.
		Set shServices = Sheets("Services")
		vl = GetRangeLookup(shServices.ListObjects("Services").DataBodyRange, _
			GetColumnName(shServices, 1, "CustomerService"), _
			Range(mColumnNames(colService) & CStr(RowIndex)).Value, _
			GetColumnName(shServices, 1, "Rate per hr"))
		Range(mColumnNames(colCharge) & CStr(RowIndex)).Formula = _
			"=" & mColumnNames(colMH) & CStr(RowIndex) & "*" & CStr(vl)
	End If

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Workbook_Open																													*
'*------------------------------------------------------------------------*
''<Description>
''The workbook has been opened. If the file doesn't have a unique
'' identifier, then create one.
''</Description>
Private Sub Workbook_Open()
	If Len(GetConfigValue("productid")) = 0 Then
		'Create a GUID.
		SetConfigValue "productid", _
			Format(Left( _
			Replace(Replace(Replace(NewGuid(), "-", ""), _
			"{", ""), "}", ""), 16), _
			"@@@@-@@@@-@@@@-@@@@")
	End If
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Workbook_SheetChange																										*
'*------------------------------------------------------------------------*
''<Description>
''A value on a sheet has changed.
''</Description>
''<Param name="sh">
''Reference to the sheet containing the changed value.
''</Param>
''<Param name="Target">
''Target range whose value has changed.
''</Param>
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
Dim bc As Boolean   'Date changed.
Dim cn() As String  'Column names, positions.
Dim cv As String    'Column value.
Dim dv As Double    'Date value.
Dim ln As String    'Line name.
Dim lp As Integer   'List position.
Dim rg As Range     'Working range.
Dim shServices As Worksheet
Dim ws As String    'Working string.

	tf = GetConfigValue("timeformat")
	If Len(sh.Name) > 0 Then
		ws = Left(sh.Name, 1)
		If ws >= "0" And ws <= "9" And _
			Target.Columns.Count = 1 And Target.Rows.Count = 1 Then
			'Numbered Month Sheet.
			If Len(mSheetName) = 0 Or mSheetName <> sh.Name Then
				mColumnNames = GetColumnNames(sh)
				mSheetName = sh.Name
			End If
			cn = mColumnNames
			cv = ColToChar(Target.Column)
			lp = Target.Row
			If (cn(colStart) = cv Or cn(colEnd) = cv) Then
				If Len(Target.Value) > 0 Then
					'Start or end times.
					If LCase(CStr(Target.Value)) = "now" Then
						'Shorthand for now.
						dv = CVDate(Now())
						bc = True
	'          Target.Value = Format(dv, "mm/dd/yyyy " & tf)
					Else
						'User has entered time.
						dv = CVDate(Target.Value)
					End If
					If dv < 1 Then
						'Time Only.
						dv = dv * 24
						bc = True
					End If
					If dv < 24 Then
						'Even time.
						dv = dv * 60
						dv = DateAdd("n", dv, CVDate(Format(Now(), "mm/dd/yyyy")))
						bc = True
					End If
					If bc Then
						Target.Value = Format(dv, "mm/dd/yyyy " & tf)
					End If
					If Len(Range(cn(colActive) & CStr(lp)).Value) = 0 Then
						'Set the Active cell to a default value of 1.
						Range(cn(colActive) & CStr(lp)).Value = "1"
					End If
					If Len(Range(cn(colMH) & CStr(Target.Row)).FormulaR1C1) = 0 Then
						'Initialize the calculations for each time row.
						'Man-hours calculation.
						Range(cn(colMH) & CStr(Target.Row)).Formula = _
							"=GetHours(" & cn(colStart) & CStr(Target.Row) & "," & _
							cn(colEnd) & CStr(Target.Row) & ")"
						'Number of billable man-hours.
						Range(cn(colBillable) & CStr(Target.Row)).Formula = _
							"=IF(" & cn(colActive) & CStr(Target.Row) & "=1," & _
							cn(colMH) & CStr(Target.Row) & ",0)"
						'Calculated charge.
						SetCharge Target.Row
						'Due.
						Range(cn(colDue) & CStr(lp)).Formula = _
							"=" & cn(colInvoiced) & CStr(lp) & "-" & _
							cn(colReceived) & CStr(lp)
						'Conditional formatting - Active.
						Set rg = Range(cn(colActive) & CStr(3))
						If rg.FormatConditions.Count = 0 Then
							rg.FormatConditions.Add Type:=xlExpression, _
								Formula1:="=$" & cn(colActive) & "3>0"
							With rg.FormatConditions(1)
								.SetFirstPriority
								.StopIfTrue = False
								With .Font
									'Dark green text.
									.Color = -16744448
									.TintAndShade = 0
								End With
								With .Interior
									'Light green background.
									.PatternColorIndex = xlAutomatic
									.Color = 13434828
									.TintAndShade = 0
								End With
								With .Borders(xlLeft)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlRight)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlTop)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlBottom)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
							End With
						End If
						'Extend the range to this cell.
						rg.FormatConditions(1).ModifyAppliesToRange _
							Range(cn(colActive) & "3:" & cn(colActive) & CStr(lp))
						'Conditional formatting - Sent.
						Set rg = Range(cn(colSent) & CStr(3))
						If rg.FormatConditions.Count = 0 Then
							rg.FormatConditions.Add Type:=xlExpression, _
								Formula1:="=$" & cn(colSent) & "3>0"
							With rg.FormatConditions(1)
								.SetFirstPriority
								With .Font
									'Dark yellow text.
									.Color = -16764058
									.TintAndShade = 0
								End With
								With .Interior
									'Light yellow background.
									.PatternColorIndex = xlAutomatic
									.Color = 13434879
									.TintAndShade = 0
								End With
								With .Borders(xlLeft)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlRight)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlTop)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlBottom)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
							End With
						End If
						'Extend the range to this cell.
						rg.FormatConditions(1).ModifyAppliesToRange _
							Range(cn(colSent) & "3:" & cn(colSent) & CStr(lp))
						'Conditional formatting - Due.
						Set rg = Range(cn(colDue) & CStr(3))
						If rg.FormatConditions.Count = 0 Then
							rg.FormatConditions.Add Type:=xlExpression, _
								Formula1:="=$" & cn(colDue) & "3>0"
							With rg.FormatConditions(1)
								.SetFirstPriority
								With .Interior
									'Light red background.
									.PatternColorIndex = xlAutomatic
									.Color = 10066431
									.TintAndShade = 0
								End With
								With .Borders(xlLeft)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlRight)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlTop)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
								With .Borders(xlBottom)
									.LineStyle = xlContinuous
									.ThemeColor = 1
									.TintAndShade = -0.14996795556505
									.Weight = xlThin
								End With
							End With
						End If
						'Extend the range to this cell.
						rg.FormatConditions(1).ModifyAppliesToRange _
							Range(cn(colDue) & "3:" & cn(colDue) & CStr(lp))
					End If
				End If
			ElseIf cv = cn(colService) Then
				'Service column has changed.
				SetCharge Target.Row
			ElseIf cv = cn(colSent) Then
				'Sent column has changed.
			End If
		ElseIf LCase(sh.Name) = "timesheet" Then
			If mUpdating = False And Target.Cells.Count = 1 Then
				mUpdating = True
				If Target.Column >= 3 And Target.Column <= 9 And _
					Target.Row - 2 Mod 7 <> 4 Then
					If Len(Target.Value) > 0 Then
						If LCase(CStr(Target.Value)) = "now" Then
							dv = CVDate(Format(Now(), tf))
							Target.Value = Format(dv, tf)
						Else
							dv = CVDate(Format(Target.Value, tf))
							Target.Value = Format(dv, tf)
						End If
					End If
				End If
				mUpdating = False
			End If
		ElseIf LCase(sh.Name) = "config" Then
			If Target.Cells.Count = 1 And _
				Target.Column = 2 Then
				ln = Replace( _
					LCase( _
					sh.Range("A" & CStr(Target.Row)).Value), _
					" ", "")
				If ln = "timeformat" Then
					'Time Format.
					If Len(Target.Value) > 0 Then
						'Time format has been changed.
						ws = Replace(Target.Value, "n", "m")
						Sheets("Work").Select
						Range("B7").NumberFormat = ws
						Range("D9").NumberFormat = ws
						Range("B13").NumberFormat = ws
						Range("D16").NumberFormat = ws
						Range("B24:B25").NumberFormat = ws
						Range("B40").NumberFormat = ws
						Range("B60:B61").NumberFormat = ws
						Sheets("Timesheet").Select
						For lp = 2 To 366 Step 7
							Range( _
								"C" & CStr(lp) & ":I" & CStr(lp + 3)).NumberFormat = ws
						Next lp
						For lp = 1 To 12
							Sheets(PadZeros(CStr(lp), 2)).Select
							Columns("B:C").Select
							Selection.NumberFormat = "MM/DD/YYYY " & ws
							Range("B2").Select
						Next lp
						Sheets("Config").Select
					End If
				ElseIf ln = "useproject" Then
'          'Use Project.
				ElseIf ln = "usetaskname" Then
'          'Use Task Name.
				End If
			End If
		End If
	End If

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Workbook_SheetSelectionChange																					*
'*------------------------------------------------------------------------*
''<Description>
''The selection area has changed on the specified sheet.
''</Description>
''<Param name="sh">
''Reference to the sheet where the selection has changed.
''</Param>
''<Param name="Target">
''Reference to the target range containing the selection.
''</Param>
Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, _
	ByVal Target As Range)
Dim cn() As String  'Column names.
Dim cv As String  'Column letter value.
Dim lp As Integer 'List position.
Dim ws As String  'Working string.

	If Len(sh.Name) > 0 Then
		ws = Left(sh.Name, 1)
		If ws >= "0" And ws <= "9" And _
			Target.Columns.Count = 1 And Target.Rows.Count = 1 Then
			'Numbered month sheet.
			If Len(mSheetName) = 0 Or mSheetName <> sh.Name Then
				mColumnNames = GetColumnNames(sh)
				mSheetName = sh.Name
			End If
			cn = mColumnNames
			cv = ColToChar(Target.Column)
			lp = Target.Row
			If cv = cn(colService) Then
				If Not CellHasValidation(Target) Then
					With Target.Validation
						.Delete
						.Add Type:=xlValidateList, _
							AlertStyle:=xlValidAlertStop, _
							Operator:=xlBetween, _
							Formula1:="=Services_CustomerService"
						.IgnoreBlank = True
						.InCellDropdown = True
						.InputTitle = ""
						.ErrorTitle = ""
						.InputMessage = ""
						.ErrorMessage = ""
						.ShowInput = True
						.ShowError = True
					End With
				End If
			End If
		End If
	End If

End Sub
'*------------------------------------------------------------------------*
