VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceivePayment 
	 Caption         =   "Receive Payment"
	 ClientHeight    =   5424
	 ClientLeft      =   108
	 ClientTop       =   456
	 ClientWidth     =   4920
	 OleObjectBlob   =   "frmReceivePayment.frx":0000
	 StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceivePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmReceivePayment.frm
' Payment receipt form.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mContacts As New ContactCollection
Private mItems As New InvoiceCollection

'*------------------------------------------------------------------------*
'* cmboContact_Change																											*
'*------------------------------------------------------------------------*
''<Description>
''Contact selection has changed.
''</Description>
Private Sub cmboContact_Change()
Dim ii As InvoiceItem
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim ou As Double    'Outstanding total.

	lstItems.Clear
	Set mItems = EnumerateOutstanding()
	lc = mItems.Count()
	If lc > 0 Then
		lstItems.AddItem "(All outstanding items)"
		For lp = 1 To lc
			Set ii = mItems.Item(lp)
			lstItems.AddItem _
				Format(ii.ItemDate, "mm/dd/yyyy") & "-" & ii.TaskName
			lstItems.List(lstItems.ListCount - 1, 1) = _
				Format(ii.InvoicedAmount - ii.ReceivedAmount, "0.00")
			ou = ou + (ii.InvoicedAmount - ii.ReceivedAmount)
		Next lp
		lstItems.List(0, 1) = Format(ou, "#,##0.00")
	End If
	txtAmountOutstanding.Text = Format(ou, "#,##0.00")
	UpdateBalance

End Sub
'*------------------------------------------------------------------------*

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
Dim bf As Boolean 'Flag - Found.
Dim ic As Integer 'Item count.
Dim ii As InvoiceItem
Dim ip As Integer 'Item position.
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.
Dim ob As Double  'Output balance.
Dim wv As Double  'Working value.

	If Len(txtAmountReceived.Text) > 0 Then
		ob = IIf( _
			IsNumeric(txtAmountReceived.Text), _
			CDbl(txtAmountReceived.Text), 0#)
	Else
		ob = 0#
	End If

	If ob > 0# And lstItems.ListCount > 0 Then
		If lstItems.Selected(0) = True Then
			'If the first item is selected, then all items are picked.
			ic = mItems.Count()
			For ip = 1 To ic
				Set ii = mItems.Item(ip)
				wv = ii.InvoicedAmount - ii.ReceivedAmount
				If wv <= ob Then
					'This item can be fully applied.
					ii.ReceivedAmount = ii.ReceivedAmount + wv
					ob = ob - wv
					bf = True
				ElseIf ob > 0# And wv > 0# Then
					ii.ReceivedAmount = ii.ReceivedAmount + ob
					ob = 0#
					bf = True
					Exit For
				End If
			Next ip
		Else
			'If the first item is not selected, then tally the items
			' that are selected.
			'List is 0-based.
			lc = lstItems.ListCount - 1
			For lp = 0 To lc
				If lstItems.Selected(lp) = True Then
					'The item is selected. Get the associated item.
					'Although list is 0-based, the first item represents all.
					Set ii = mItems.Item(lp)
					wv = ii.InvoicedAmount - ii.ReceivedAmount
					If wv <= ob Then
						'This item can be fully applied.
						ii.ReceivedAmount = ii.ReceivedAmount + wv
						ob = ob - wv
						bf = True
					ElseIf ob > 0# And wv > 0# Then
						ii.ReceivedAmount = ii.ReceivedAmount + ob
						ob = 0#
						bf = True
						Exit For
					End If
				End If
			Next lp
		End If
	End If
	If bf = True Then
		SetReceivedAmounts mItems
	End If
	Unload Me

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* EnumerateOutstanding																										*
'*------------------------------------------------------------------------*
''<Description>
''Enumerate the outstanding invoice items.
''</Description>
''<Returns>
''Invoice collection containing the items having outstanding payment
'' balances.
''</Returns>
Private Function EnumerateOutstanding() As InvoiceCollection
Dim cc As String    'Contact code.
Dim cn() As String  'Column names.
Dim cs As String    'Selected contact name.
Dim ic As New InvoiceCollection
Dim ii As InvoiceItem
Dim lp As Integer   'List position.
Dim sh As Worksheet 'Current sheet.
Dim si As Integer   'Sheet index.
Dim sn As String    'Sheet name.
Dim ws As String    'Working string.

	cs = cmboContact.Text
	cc = mContacts.GetCodeFromName(cs)
	For si = 1 To 12
		sn = PadZeros(CStr(si), 2)
		Set sh = Sheets(sn)
		cn = GetColumnNames(sh)
		lp = 3
		ws = sh.Range(cn(colTask) & CStr(lp)).Value
		Do While Len(ws) > 0
			Set ii = New InvoiceItem
			ii.InvoicedAmount = sh.Range(cn(colInvoiced) & CStr(lp)).Value
			ii.ItemDate = sh.Range(cn(colEnd) & CStr(lp)).Value
			ii.ManHoursSpent = sh.Range(cn(colMH) & CStr(lp)).Value
			ii.ProjectName = sh.Range(cn(colProject) & CStr(lp)).Value
			ii.ReceivedAmount = sh.Range(cn(colReceived) & CStr(lp)).Value
			ii.RowIndex = lp
			ii.Sent = _
				IIf(sh.Range(cn(colSent) & CStr(lp)).Value = 1, True, False)
			ii.SetContactCodeService sh.Range(cn(colService) & CStr(lp)).Value
			ii.SheetName = sn
			ii.TaskName = sh.Range(cn(colTask) & CStr(lp)).Value
			If ii.ContactCode = cc And ii.Sent = True And _
				ii.InvoicedAmount - ii.ReceivedAmount <> 0# Then
				ic.Add ii
			End If
			lp = lp + 1
			ws = sh.Range(cn(colTask) & CStr(lp)).Value
		Loop
	Next si
	Set EnumerateOutstanding = ic

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* lstItems_Change																												*
'*------------------------------------------------------------------------*
''<Description>
''Selection has changed on the items list.
''</Description>
Private Sub lstItems_Change()
	UpdateBalance
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* txtAmountReceived_Change																								*
'*------------------------------------------------------------------------*
''<Description>
''The amount received text has changed.
''</Description>
Private Sub txtAmountReceived_Change()
	UpdateBalance
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UpdateBalance																													*
'*------------------------------------------------------------------------*
''<Description>
''Update the balance of payment received minus selected items.
''</Description>
Private Sub UpdateBalance()
Dim bf As Boolean 'Flag - Found.
Dim ic As Integer 'Item count.
Dim ii As InvoiceItem
Dim ip As Integer 'Item position.
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.
Dim ob As Double  'Output balance.
Dim os As Double  'Output selected.
Dim wv As Double  'Working value.

	If Len(txtAmountReceived.Text) > 0 Then
		ob = IIf( _
			IsNumeric(txtAmountReceived.Text), _
			CDbl(txtAmountReceived.Text), 0#)
	Else
		ob = 0#
	End If
	os = 0#
	If lstItems.ListCount > 0 Then
		If lstItems.Selected(0) = True Then
			'If the first item is selected, then all items are picked.
			ic = mItems.Count()
			For ip = 1 To ic
				Set ii = mItems.Item(ip)
				wv = ii.InvoicedAmount - ii.ReceivedAmount
				os = os + wv
				ob = ob - wv
			Next ip
		Else
			'If the first item is not selected, then tally the items
			' that are selected.
			'List is 0-based.
			lc = lstItems.ListCount - 1
			For lp = 0 To lc
				If lstItems.Selected(lp) = True Then
					'The item is selected. Get the associated item.
					'Although list is 0-based, the first item represents all.
					Set ii = mItems.Item(lp)
					wv = ii.InvoicedAmount - ii.ReceivedAmount
					os = os + wv
					ob = ob - wv
				End If
			Next lp
		End If
	End If
	txtSelected.Text = Format(os, "#,##0.00")
	txtBalance.Text = Format(ob, "#,##0.00")
	bf = False
	lc = lstItems.ListCount - 1
	For lp = 0 To lc
		If lstItems.Selected(lp) = True Then
			bf = True
			Exit For
		End If
	Next lp
	wv = 0#
	If Len(txtAmountReceived.Text) > 0 Then
		wv = IIf( _
			IsNumeric(txtAmountReceived.Text), _
			CDbl(txtAmountReceived.Text), 0#)
	End If
	cmdOK.Enabled = (bf = True And wv > 0#)

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UserForm_Initialize																										*
'*------------------------------------------------------------------------*
''<Description>
''The user form is being initialized.
''</Description>
Private Sub UserForm_Initialize()
Dim ci As ContactItem
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.

	mContacts.FillFromSheet
	lc = mContacts.Count()
	For lp = 1 To lc
		Set ci = mContacts.Item(lp)
		cmboContact.AddItem ci.BillToName
	Next lp

End Sub
'*------------------------------------------------------------------------*
