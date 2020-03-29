VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateInvoice 
	 Caption         =   "Create Invoice"
	 ClientHeight    =   5100
	 ClientLeft      =   120
	 ClientTop       =   468
	 ClientWidth     =   5928
	 OleObjectBlob   =   "frmCreateInvoice.frx":0000
	 StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'frmCreateInvoice.frm
' Invoice creation form.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mContacts As New ContactCollection
Private mDefaultDate As String
Private mDefaultInvoice As String
Private mRenamedServices As New NameValueCollection
Private mServices As New ContactServiceCollection

'*** PRIVATE ***
'*------------------------------------------------------------------------*
'* cmboContact_Change																											*
'*------------------------------------------------------------------------*
''<Description>
''The selected item on the contact drop-down list has changed.
''</Description>
Private Sub cmboContact_Change()
	UpdateServicesList
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* cmboScope_Change																												*
'*------------------------------------------------------------------------*
''<Description>
''The selected item on the scope drop-down list has changed.
''</Description>
Private Sub cmboScope_Change()
	UpdateContactsList
	UpdateServicesList
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* cmdCancel_Click																												*
'*------------------------------------------------------------------------*
''<Description>
''The Cancel button has been clicked.
''</Description>
Private Sub cmdCancel_Click()
	Unload Me
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* cmdOK_Click																														*
'*------------------------------------------------------------------------*
''<Description>
''The OK button has been clicked.
''</Description>
Private Sub cmdOK_Click()
Dim bc As Boolean   'Flag - Continue.
Dim bf As Boolean   'Flag - At least one row formatted.
Dim dd As Date      'Date due.
Dim di As Date      'Invoice date.
Dim cc As String    'Customer code.
Dim ci As Integer		'Customer index.
Dim cn As String    'Customer name.
Dim ic As InvoiceCollection   'Working invoice items.
Dim id As Integer             'Detail index.
Dim ii As InvoiceItem         'Current item.
Dim lp As Integer   'List position.
Dim ni As String    'Invoice number.
Dim pc As New Collection  'Project names.
Dim rb As Integer   'Bottom row.
Dim rc As Integer   'Row count.
Dim rd As Integer   'Detail row.
Dim rt As Integer   'Top row.
Dim sc As Worksheet 'Contact sheet.
Dim sh As Worksheet 'Current sheet.
Dim si As ContactServiceItem  'Current service item.
Dim tc As ListObject					'Contact table.
Dim tg As Range								'Table data body range.
Dim ws As String    'Working string.

	bc = True
	bf = False
	On Error Resume Next
	Err.Clear
	If cmboContact.ListIndex > -1 And _
		cmboScope.ListIndex > -1 And _
		Len(txtDueDate.Text) > 0 And _
		Len(txtInvoiceDate.Text) > 0 And _
		Len(txtInvoiceNumber.Text) > 0 Then
		'A contact and scope have been selected, the due date has
		' been filled, the date of invoice has been set, and the
		' invoice number is ready.
		'Check for data errors.
		dd = CDate(txtDueDate.Text)
		If Err.Number <> 0 Then
			ws = "Error: Invalid Due Date..."
			bc = False
			Err.Clear
		End If
		If bc = True Then
			di = CDate(txtInvoiceDate.Text)
			If Err.Number <> 0 Then
				ws = "Error: Invalid Invoice Date..."
				bc = False
				Err.Clear
			End If
		End If
		If bc = True Then
			If Not IsNumeric(txtInvoiceNumber.Text) Then
				ws = "Error: Invalid Invoice Number..."
				bc = False
				Err.Clear
			End If
		End If
		If bc = True Then
			'Get company code.
			cn = cmboContact.Text
			cc = mContacts.GetCodeFromName(cn)
			If Len(cc) = 0 Then
				'Company code not found.
				bc = False
			Else
				ci = mContacts.GetCodeIndex(cc)
				If ci = 0 Then bc = False
			End If
		End If
		If bc = True Then
			'Valid values. Generate Invoice.
			Set sh = Sheets("ServiceInvoice")
			rt = Range("serviceInvoiceHeaderRow").Row
			rb = Range("serviceInvoiceFooterRow").Row
			lp = rt + 1
			Set ic = EnumerateItemsCI(True, True)
			If ic.Count > 0 Then
				'The following line was added in an attempt to keep the
				' Invoice Detail from getting filled with Log Report rows.
				Range("A1").Select
				'Build the appropriate number of rows on the Service Invoice page.
				sh.Select
				Columns("K:K").EntireColumn.Hidden = False
				Columns("K:K").EntireColumn.ColumnWidth = _
					Range("C1").ColumnWidth + _
					Range("D1").ColumnWidth + _
					Range("E1").ColumnWidth + _
					Range("F1").ColumnWidth
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
						FormatRowCI rd
						bf = True
						rc = rc + 1
						rb = rb + 1
					End If
					'At this time, there are at least two items to invoice on, and
					' at least two rows.
					Do While rc < ic.Count
						Rows(CStr(rd + 1) & ":" & CStr(rd + 1)).Select
						Selection.Insert _
							Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
						FormatRowCI rd + 1
						bf = True
						rc = rc + 1
						rb = rb + 1
					Loop
					'At this time, the number of items is equal to the number of rows.
					Range("A" & CStr(rd)).Select
				ElseIf ic.Count = rc Then
					'Row counts are equal.
				Else
					'Rows need to be deleted.
					Do While rc > ic.Count
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
					Set ii = ic.Item(id)
					'Service invoice columns.
					'Detail:
					' A - Date.
					' B - Project name.
					' C - Task.
					' G - Man hours.
					' I - Rate.
					' J - Charge.
					' K - Scratchpad cell.
					Range("A" & CStr(lp)).Value = Format(ii.ItemDate, "MM/DD/YYYY")
					Range("B" & CStr(lp)).Value = ii.ProjectName
					Set si = Nothing
					If Len(ii.ServiceName) > 0 Then
						Set si = mServices.Item(ii.ContactCode & "_" & ii.ServiceName)
						ws = si.ServiceName & " : " & ii.TaskName
					Else
						ws = ii.TaskName
					End If
					Range("C" & CStr(lp)).Value = ws
					Range("K" & CStr(lp)).Value = ws
					Range("G" & CStr(lp)).Value = Round(ii.ManHoursSpent, 2)
					If Not si Is Nothing Then
						Range("I" & CStr(lp)).Value = si.Rate
					Else
						Range("I" & CStr(lp)).Value = Range("configHourlyRate").Value
					End If
					ii.InvoicedAmount = Round(ii.ManHoursSpent, 2) * si.Rate
					Range("J" & CStr(lp)).Formula = "=G" & CStr(lp) & "*I" & CStr(lp)
					FormatRowCI lp, IIf(lp = rb - 1, True, False)
					Rows(CStr(lp) & ":" & CStr(lp)).AutoFit = True
					id = id + 1
				Next lp
				'Color with alternate background.
				Range("A" & CStr(rd) & ":J" & CStr(rb - 1)).Select
				AlternateBackColor
				Range("A" & CStr(rb)).Select
				Set tc = Sheets("Contacts").ListObjects("Contacts")
				Set tg = tc.DataBodyRange
				'Set Bill-To.
				Range("serviceInvoiceBillToName").Value = _
					tg.Cells(ci, _
					tc.ListColumns("Bill To Name").Index)
				Range("serviceInvoiceBillToAddress").Value = _
					tg.Cells(ci, _
					tc.ListColumns("Bill To Address").Index)
				Range("serviceInvoiceBillToCityStateZip").Value = _
					tg.Cells(ci, _
					tc.ListColumns("Bill To City State Zip").Index)
				'Set Ship-To.
				Range("serviceInvoiceShipToName").Value = _
					tg.Cells(ci, _
					tc.ListColumns("Ship To Name").Index)
				Range("serviceInvoiceShipToAddress").Value = _
					tg.Cells(ci, _
					tc.ListColumns("Ship To Address").Index)
				Range("serviceInvoiceShipToCityStateZip").Value = _
					tg.Cells(ci, _
					tc.ListColumns("Ship To City State Zip").Index)
				'Set Invoice Number.
				Range("serviceInvoiceNumber").Value = _
					txtInvoiceNumber.Text
				'Set Invoice Date.
				Range("serviceInvoiceDate").Value = _
					Format(txtInvoiceDate.Text, "MM/DD/YYYY")
        'Set Project Name.
				ws = ""
				sh.Range("serviceInvoiceProjectName").Value = ""
				If chkProjectCite.Value = True Then
					'Cite all project names in the project cell.
					rc = ic.Count()
					For rd = 1 To rc
						Set ii = ic.Item(rd)
						If Len(ii.ProjectName) > 0 Then
							AddUniqueSorted pc, ii.ProjectName
						End If
					Next rd
					rc = pc.Count()
					For rd = 1 To rc
						If rd > 1 Then
							ws = ws & ", "
						End If
						ws = ws & pc.Item(rd)
					Next rd
					If Len(ws) > 0 Then
						sh.Range("serviceInvoiceProjectName").Value = ws
					End If
				End If
				'Set Due Date.
				Range("serviceInvoiceDueDate").Value = _
					Format(txtDueDate.Text, "MM/DD/YYYY")
				'Set Last-Known Invoice.
				Range("configLastInvoice").Value = txtInvoiceNumber.Text
				'Show Print Dialog.
				Application.CommandBars.ExecuteMso "PrintPreviewAndPrint"
				'Mark items as sent.
				MarkItemsSent ic
				'Set the Outlook notification.
				If GetConfigValue("Outlook Notify") = "1" Then
					AddOutlookInvoiceNotification _
						txtInvoiceNumber.Text, CDate(txtDueDate.Text)
				End If
				Unload Me
			Else
				'No items found.
				bc = False
			End If
		Else
			'Error found. Exit with message.
			MsgBox ws, vbOKOnly, "Generate Invoice"
		End If
	Else
		'Error found. Exit with message.
		If Len(txtProject.Text) = 0 Then
			ws = "Project Name is optional. " & _
				"Please set all other empty fields."
		Else
			ws = "Please set all empty fields first."
		End If
		MsgBox ws, vbOKOnly, "Generate Invoice"
	End If
	If bc = False Then
		cmdCancel_Click
	End If

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* EnumerateItemsCI																												*
'*------------------------------------------------------------------------*
''<Description>
''Enumerate the items to be invoiced, and return the collection to the
'' caller.
''</Description>
''<Param name="ShowMessage">
''Value indicating whether to show UI messages based on status. Default =
'' True.
''</Param>
''<Param name="ServiceFilter">
''Value indicating whether to filter by the selected items in the services
'' list. Default = False.
''</Param>
''<Returns>
''Collection of invoice items matching the scope, selected company,
'' project name, and optionally, services.
''</Returns>
Private Function EnumerateItemsCI( _
	Optional ShowMessage As Boolean = True, _
	Optional ServiceFilter As Boolean = False) _
	As InvoiceCollection
'Enumerate the items to be invoiced, and return the collection to the caller.
' 0 - All uninvoiced items.
' 1 - Uninvoiced items on sheet.
' 2 - Selected items only.
Dim cn() As String    'Column names.
Dim ii As InvoiceItem 'Current Invoice Item.
Dim ic As New InvoiceCollection
Dim lb As Integer     'List begin.
Dim le As Integer     'List end.
Dim lp As Integer     'List position.
Dim rc As New InvoiceCollection 'Return collection.
Dim sh As Worksheet   'Current sheet.
Dim si As Integer     'Sheet index.
Dim ss As String      'Source string.
Dim ts As String      'Target string.

	'Get the contact selected.
	'Name is listed locally, code is found in services list.
	Select Case cmboScope.ListIndex
		Case 0:   'Invoice on all uninvoiced items.
			For si = 1 To 12
				Set sh = Sheets(PadZeros(CStr(si), 2))
				cn = GetColumnNames(sh)
				lp = 3
				Do While Len(sh.Range(cn(colTask) & CStr(lp)).Value) > 0
					If Len(sh.Range(cn(colActive) & CStr(lp)).Value) > 0 And _
						sh.Range(cn(colActive) & CStr(lp)).Value <> 0 And _
						sh.Range(cn(colSent) & CStr(lp)).Value <> 1 And _
						Len(sh.Range(cn(colStart) & CStr(lp)).Value) > 0 And _
						Len(sh.Range(cn(colEnd) & CStr(lp)).Value) > 0 Then
						'The item is active and has not been sent.
						Set ii = New InvoiceItem
						ii.SheetName = sh.Name
						ii.RowIndex = lp
						ii.ItemDate = _
							Format(sh.Range(cn(colEnd) & CStr(lp)).Value, "MM/DD/YYYY")
						ii.ManHoursSpent = sh.Range(cn(colBillable) & CStr(lp)).Value
						ii.ProjectName = sh.Range(cn(colProject) & CStr(lp)).Value
						ii.TaskName = sh.Range(cn(colTask) & CStr(lp)).Value
						ss = sh.Range(cn(colService) & CStr(lp)).Value
						ts = mRenamedServices.GetValue(ss)
						If Len(ts) > 0 Then
							ss = ts
						End If
						ii.SetContactCodeService ss
						ic.Add ii
					End If
					lp = lp + 1
				Loop
			Next si
			If ic.Count = 0 And ShowMessage = True Then
				MsgBox "No outstanding items were found to invoice for. " & _
					"Process cancelled...", vbOKOnly, "Generate Invoice"
			End If
		Case 1:     'Invoice non-invoiced items on current sheet.
			If IsNumeric(ActiveSheet.Name) Then
				'Sheet is selected.
				Set sh = ActiveSheet
				cn = GetColumnNames(sh)
				lp = 3
				Do While Len(sh.Range(cn(colTask) & CStr(lp)).Value) > 0
					If Len(sh.Range(cn(colActive) & CStr(lp)).Value) > 0 And _
						sh.Range(cn(colActive) & CStr(lp)).Value <> 0 And _
						sh.Range(cn(colSent) & CStr(lp)).Value <> 1 And _
						Len(sh.Range(cn(colStart) & CStr(lp)).Value) > 0 And _
						Len(sh.Range(cn(colEnd) & CStr(lp)).Value) > 0 Then
						'The item is active and has not been sent.
						Set ii = New InvoiceItem
						ii.SheetName = sh.Name
						ii.RowIndex = lp
						ii.ItemDate = _
							Format(sh.Range(cn(colEnd) & CStr(lp)).Value, "MM/DD/YYYY")
						ii.ManHoursSpent = sh.Range(cn(colBillable) & CStr(lp)).Value
						ii.ProjectName = sh.Range(cn(colProject) & CStr(lp)).Value
						ii.TaskName = sh.Range(cn(colTask) & CStr(lp)).Value
						ss = sh.Range(cn(colService) & CStr(lp)).Value
						ts = mRenamedServices.GetValue(ss)
						If Len(ts) > 0 Then
							ss = ts
						End If
						ii.SetContactCodeService ss
						ic.Add ii
					End If
					lp = lp + 1
				Loop
				If ic.Count = 0 And ShowMessage = True Then
					MsgBox "No outstanding items were found in the current sheet. " & _
						"Process cancelled...", vbOKOnly, "Generate Invoice"
				End If
			ElseIf ShowMessage = True Then
				MsgBox "Please first select a monthly journal sheet to invoice " & _
					"for items on current sheet. " & _
					"Process cancelled...", vbOKOnly, "Generate Invoice"
			End If
		Case 2:     'Invoice Selected Items on Current Sheet.
			If IsNumeric(ActiveSheet.Name) Then
				'Sheet is selected.
				Set sh = ActiveSheet
				cn = GetColumnNames(sh)
				If Selection.Row >= 3 And Selection.Rows.Count > 0 Then
					lb = Selection.Row
					le = lb + Selection.Rows.Count - 1
					For lp = lb To le
						If Len(sh.Range(cn(colActive) & CStr(lp)).Value) > 0 And _
							sh.Range(cn(colActive) & CStr(lp)).Value <> 0 And _
							sh.Range(cn(colSent) & CStr(lp)).Value <> 1 And _
							Len(sh.Range(cn(colStart) & CStr(lp)).Value) > 0 And _
							Len(sh.Range(cn(colEnd) & CStr(lp)).Value) > 0 Then
							'The item is active and has not been sent.
							Set ii = New InvoiceItem
							ii.SheetName = sh.Name
							ii.RowIndex = lp
							ii.ItemDate = _
								Format(sh.Range(cn(colEnd) & CStr(lp)).Value, "MM/DD/YYYY")
							ii.ManHoursSpent = sh.Range(cn(colBillable) & CStr(lp)).Value
							ii.ProjectName = sh.Range(cn(colProject) & CStr(lp)).Value
							ii.TaskName = sh.Range(cn(colTask) & CStr(lp)).Value
							ss = sh.Range(cn(colService) & CStr(lp)).Value
							ts = mRenamedServices.GetValue(ss)
							If Len(ts) > 0 Then
								ss = ts
							End If
							ii.SetContactCodeService ss
							ic.Add ii
						End If
					Next lp
					If ic.Count = 0 And ShowMessage = True Then
						MsgBox "No outstanding items were found in the current " & _
							"selection. " & _
							"Process cancelled...", vbOKOnly, "Generate Invoice"
					End If
				ElseIf ShowMessage = True Then
					MsgBox "No outstanding items were found in the current " & _
						"selection. " & _
						"Process cancelled...", vbOKOnly, "Generate Invoice"
				End If
			ElseIf ShowMessage = True Then
				MsgBox "Please first select a monthly journal sheet to invoice " & _
					"for selected items on current sheet. " & _
					"Process cancelled...", vbOKOnly, "Generate Invoice"
			End If
	End Select
	If cmboContact.ListCount > 0 Then
		Set rc = FilterCompanyProjectService(ic, ServiceFilter)
	Else
		Set rc = ic
	End If
	Set EnumerateItemsCI = rc

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* FilterCompanyProjectService																						*
'*------------------------------------------------------------------------*
''<Description>
''Return a collection of invoice items that are filtered by company,
'' project name, and optionally, selected service items.
''</Description>
''<Param name="Invoices">
''Collection of previously selected invoice items.
''</Param>
''<Param name="ServiceFilter">
''Value indicating whether selected service items should be applied to
'' the filter.
''</Param>
''<Returns>
''Collection of invoice items matching the additional filters.
''</Returns>
Private Function FilterCompanyProjectService( _
	Invoices As InvoiceCollection, _
	Optional ServiceFilter As Boolean = False) As InvoiceCollection
Dim bc As Boolean   'Flag - Continue.
Dim cc As String    'Company code to which the filter applies.
Dim cf As String    'Company name found.
Dim cn As String    'Company name to which the filter applies.
Dim ic As InvoiceCollection 'Collection of items picked.
Dim ii As InvoiceItem       'Current picked item.
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim rc As New InvoiceCollection   'Return collection.
Dim ri As InvoiceItem   'Item in the return collection.
Dim sc As Integer   'Services count.
Dim sn As String    'Service name.
Dim sp As Integer   'Services position.

	'Filter from the raw list to result in items matching contact,
	' project name, and services list.
	Set ic = Invoices
	cn = cmboContact.Text
	If Len(cn) > 0 Then
		cc = mContacts.GetCodeFromName(cn)
	End If
	If Len(cc) > 0 Then
		lc = ic.Count()
		For lp = 1 To lc
			'Split CompanyCode_ServiceName to find matching code.
			cf = ""
			sn = ""
			Set ii = ic.Item(lp)
			If Len(ii.ContactCode) = 0 And Len(ii.ServiceName) = 0 Then
				'Service name not specified on sheet. This item can be
				' placed on any selected customer if "(All including unassigned)"
				' is selected on the services list.
				bc = False
				If ServiceFilter = False Or lstServices.ListCount = 0 Then
					bc = True
				ElseIf lstServices.Selected(0) = True Then
					'This statement is only evaluated if there are
					' known to be items in the list.
					'The bug requiring this separation appears to have been
					' corrected in recent versions of VBA.
					'TODO: Test to see if the above expression can reside on the
					' same line as the previous expression when the lstServices
					' list contains zero items. If it can, then the bc flag can
					' be eliminated for setting cf.
					bc = True
				End If
				If bc = True Then
					cf = cc
				End If
			Else
				bc = False
				If ServiceFilter = False Or lstServices.ListCount = 0 Then
					bc = True
				ElseIf lstServices.Selected(0) = True Then
					bc = True
				Else
					'lstServices is 0-based.
					sc = lstServices.ListCount - 1
					For sp = 1 To sc
						If lstServices.List(sp) = ii.ServiceName And _
							lstServices.Selected(sp) = True Then
							bc = True
						End If
					Next sp
				End If
				If bc = True Then
					cf = ii.ContactCode
					sn = ii.ServiceName
				End If
			End If
			'If the company and project both match, then save the result.
			If cf = cc And _
				(LCase(txtProject.Text) = LCase(ii.ProjectName) Or _
				Len(txtProject.Text) = 0) Then
				'Matching company / project / service found.
				Set ri = New InvoiceItem
				ri.ContactCode = cc
				ri.ItemDate = ii.ItemDate
				ri.ManHoursSpent = ii.ManHoursSpent
				ri.ProjectName = ii.ProjectName
				ri.RowIndex = ii.RowIndex
				ri.ServiceName = sn
				ri.SheetName = ii.SheetName
				ri.TaskName = ii.TaskName
				rc.Add ri
			End If
		Next lp
	End If
	Set FilterCompanyProjectService = rc

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* FormatRowCI																														*
'*------------------------------------------------------------------------*
''<Description>
''Format the specified invoice row.
''</Description>
''<Param name="RowIndex">
''Spreadsheet row to format.
''</Param>
''<Param name="IsBottom">
''Value indicating whether the row to be formatted is the last row in the
'' list. Default = False.
''</Param>
Private Sub FormatRowCI(RowIndex As Integer, _
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
	Range("C" & CStr(rd) & ":F" & CStr(rd)).Select
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
	Range("A" & CStr(rd) & ":F" & CStr(rd)).Select
	With Selection
		.HorizontalAlignment = xlLeft
	End With
	If Err.Number <> 0 Then
		MsgBox "Error: " & Err.Description
		Err.Clear
	End If
	DoEvents
	Range("G" & CStr(rd) & ":J" & CStr(rd)).Select
	With Selection
		.HorizontalAlignment = xlRight
		.NumberFormat = "#,##0.00"
	End With
	DoEvents

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* txtProject_Exit																												*
'*------------------------------------------------------------------------*
''<Description>
''The project name textbox has lost focus.
''</Description>
''<Param name="Cancel">
''Value indicating whether to cancel the event.
''</Param>
Private Sub txtProject_Exit(ByVal Cancel As MSForms.ReturnBoolean)
	UpdateServicesList
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UpdateContactsList																											*
'*------------------------------------------------------------------------*
''<Description>
''Update the list of contacts based on the current selection of the scope.
''</Description>
Private Sub UpdateContactsList()
Dim ci As ContactItem
Dim ic As InvoiceCollection
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.
Dim ws As String  'Working string.

	cmboContact.Clear
	Set ic = EnumerateItemsCI(False, False)
	lc = mContacts.Count()
	For lp = 1 To lc
		Set ci = mContacts.Item(lp)
		If ic.ContactCodeExists(ci.Code) Then
			cmboContact.AddItem ci.BillToName
		End If
	Next lp
	If cmboContact.ListCount > 0 Then
		cmboContact.ListIndex = 0
	End If

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UpdateServicesList																											*
'*------------------------------------------------------------------------*
''<Description>
''Update the list of services based on the current settings and selections
'' on the form.
''</Description>
Private Sub UpdateServicesList()
Dim bc As Boolean 'Flag - Continue.
Dim bf As Boolean 'Flag - Found.
Dim cc As String  'Contact code.
Dim ci As ContactItem
Dim cn As String    'Contact name.
Dim cs() As String  'Code_Service.
Dim fu As frmUndefinedServices
Dim ic As InvoiceCollection
Dim ii As InvoiceItem
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim nc As New NameValueCollection 'Names and values.
Dim ni As NameValueItem
Dim sc As Collection  'Service name collection.
Dim si As ContactServiceItem
Dim sn As String    'Service name.
Dim tc As Integer   'Target count.
Dim tp As Integer   'Target position.
Dim ws As String    'Working string.

	'Reset the services selection list.
	lstServices.Clear
	'Get the contact selected.
	'Name is listed locally, code is found in services list.
	cn = cmboContact.Text
	If Len(cn) > 0 Then
		cc = mContacts.GetCodeFromName(cn)
	End If
	If Len(cc) > 0 Then
		Set ic = EnumerateItemsCI(False, False)
		lc = ic.Count
		If lc > 0 Then
			'Matching items for this contact selection have been found.
			Set sc = ic.GetServiceList()
			lc = sc.Count()
			If lc > 0 Then
				lstServices.AddItem "(All including unassigned)"
			End If
			'Add services for selected contact.
			For lp = 1 To lc
				sn = sc.Item(lp)
				If Not mServices.IsDefinedContactCodeService(cc, sn) Then
					'Service not defined for this contact.
					nc.Add cc, sn
				Else
					'Service is defined for contact.
					lstServices.AddItem sn
				End If
			Next lp
			'Resolve any missing contact / service relation.
			lc = nc.Count()
			If lc > 0 Then
				'Resolve the relations of missing services on this contact.
				For lp = 1 To lc
					Set ni = nc.Item(lp)
					Set ii = ic.FirstContactCodeServiceName(ni.Name, ni.Value)
					Set fu = New frmUndefinedServices
					fu.SetMonthSheet ii.SheetName
					fu.SetHourlyRates _
						CDbl(GetConfigValue("Hourly Rate")), _
						CDbl(GetConfigValue("Commission"))
					fu.SetContactService ni.Name, ni.Value
					fu.Show 1
					bc = False
					Err.Clear
					On Local Error Resume Next
					If fu.IsSubmitted Then
						If Err.Number = 0 Then bc = True
					End If
					On Error GoTo 0
					If bc = True Then
						'The OK button was clicked. Set the choice.
						If fu.ChangeContact = True Then
							'Change to another contact with the same service.
							'These items will disappear from the current invoice
							' setup because of selected customer.
							mRenamedServices.Add _
								ni.Name & "_" & ni.Value, _
								fu.ChangeContactCode & "_" & ni.Value
							If fu.UpdateSheet = True Then
								'Update the data sheets if specified.
								MonthSheetsRenameService ic, _
									ni.Name, ni.Value, _
									fu.ChangeContactCode, ni.Value
							End If
							tc = ic.Count()
							For tp = 1 To tc
								Set ii = ic.Item(tp)
								If ii.ContactCode = ni.Name And _
									ii.ServiceName = ni.Value Then
									'Remove the item from this list.
									ic.Remove tp
									tc = tc - 1   'Decount.
									tp = tp - 1   'Deindex.
									If tc = 0 Then Exit For
								End If
							Next tp
							If lstServices.ListCount = 1 Then
								lstServices.Clear
							End If
						ElseIf fu.ChangeService = True Then
							'Change to another service under the same contact.
							'These items will remain in the list.
							mRenamedServices.Add _
								ni.Name & "_" & ni.Value, _
								ni.Name & "_" & fu.ChangeServiceName
							If fu.UpdateSheet = True Then
								'Update the data sheets if specified.
								MonthSheetsRenameService ic, _
									ni.Name, ni.Value, _
									ni.Name, fu.ChangeServiceName
							End If
							'Update the service list, if applicable.
							bf = False
							ws = fu.ChangeServiceName
							tc = lstServices.ListCount
							If tc > 1 Then
								For tp = 1 To tc
									If lstServices.List(tp) = ws Then
										bf = True
										Exit For
									End If
								Next tp
							End If
							If bf = False Then
								'The service can be added to the list.
								lstServices.AddItem ws
							End If
							'Update the loaded service reference.
							tc = ic.Count
							For tp = 1 To tc
								Set ii = ic.Item(tp)
								If ii.ContactCode = ni.Name And _
									ii.ServiceName = ni.Value Then
									'Change the underlying reference to the new service.
									ii.ServiceName = ws
								End If
							Next tp
						ElseIf fu.NewService Then
							'Assign the default rate to the currently selected
							' Contact/Service combination.
							Set ci = mContacts.Code(ni.Name)
							If Not ci Is Nothing Then
								'Contact exists.
								'Create the local reference.
								Set si = New ContactServiceItem
								Set si.Contact = ci
								si.Commission = CDbl(fu.NewCommissionValue)
								si.Rate = CDbl(fu.NewRateValue)
								si.ServiceName = ni.Value
								mServices.Add si
								'Create the record in the Services table.
								If fu.UpdateSheet = True Then
									AddService ni.Name, ni.Value, si.Rate, si.Commission
								End If
								'Update the service list, if applicable.
								bf = False
								ws = ni.Value
								tc = lstServices.ListCount
								If tc > 1 Then
									For tp = 1 To tc
										If lstServices.List(tp) = ws Then
											bf = True
											Exit For
										End If
									Next tp
								End If
								If bf = False Then
									'The service can be added to the list.
									lstServices.AddItem ws
								End If
							Else
								'Contact needs to be created.
							End If
						End If
					Else
						'If a decision was not made for a non-existent relation,
						' then let's just unload.
						Unload Me
					End If
				Next lp
			End If
			If lstServices.ListCount > 0 Then
				'First item is selected by default.
				lstServices.Selected(0) = True
			End If
		End If
	End If
	If lstServices.ListCount > 0 Then
		cmdOK.Enabled = True
	Else
		cmdOK.Enabled = False
	End If

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UserForm_Initialize																										*
'*------------------------------------------------------------------------*
''<Description>
''User form is initializing.
''</Description>
Private Sub UserForm_Initialize()
Dim dt As Date      'Working Date.
Dim lp As Integer   'List Position.
Dim sh As Worksheet 'Working Sheet.
Dim ws As String    'Working String.

	On Error Resume Next
	'Contacts reference.
	mContacts.FillFromSheet
	'Services reference.
	mServices.FillFromSheet mContacts
	'Scope.
	cmboScope.List = Array("All uninvoiced items", _
		"Uninvoiced items on sheet", _
		"Selected items only")
	If cmboScope.ListCount > 0 Then
		cmboScope.ListIndex = 0
	End If
	'Invoice Date.
	mDefaultDate = Format(Date, "MM/DD/YYYY")
	txtInvoiceDate.Text = mDefaultDate
	'Invoice Number.
	ws = Range("configLastInvoice").Value
	If Len(ws) = 10 Then
		'Date Timestamp + Counter.
		dt = CDate(Mid(ws, 5, 2) & "/" & _
			Mid(ws, 7, 2) & "/" & _
			Mid(ws, 1, 4))
		If dt = Date Then
			'Previous invoice has been issued today.
			mDefaultInvoice = Format(Date, "YYYYMMDD") & _
				PadZeros(CStr(CInt(Right(ws, 2)) + 1), 2)
		Else
			'Previous invoice was issued on another day.
			mDefaultInvoice = Format(Date, "YYYYMMDD") & _
				"01"
		End If
	ElseIf IsNumeric(ws) Then
		'Some numeric value.
		mDefaultInvoice = CStr(CDbl(ws) + 1)
	End If
	txtInvoiceNumber.Text = mDefaultInvoice
	'Due Date.
	txtDueDate.Text = Format(DateAdd("d", 15, CDate(txtInvoiceDate.Text)), _
		"MM/DD/YYYY")

End Sub
'*------------------------------------------------------------------------*
