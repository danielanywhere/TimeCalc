VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUndefinedServices 
	 Caption         =   "Undefined Service"
	 ClientHeight    =   6432
	 ClientLeft      =   108
	 ClientTop       =   456
	 ClientWidth     =   4644
	 OleObjectBlob   =   "frmUndefinedServices.frx":0000
	 StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUndefinedServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmUndefinedServices.frm
' Form to resolve conflict for services not found.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mCommission As Double
Private mContacts As New ContactCollection
Private mContactSelected As ContactItem
Private mRate As Double
Private mServices As New ContactServiceCollection
Private mServiceSelected As ContactServiceItem
Private mSubmitted As Boolean 'Value indicating whether the form was submitted.

'*** PRIVATE ***
'*------------------------------------------------------------------------*
'* cmdCancel_Click																												*
'*------------------------------------------------------------------------*
''<Description>
''Cancel button has been clicked.
''</Description>
Private Sub cmdCancel_Click()
	Me.Hide
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* cmdOK_Click																														*
'*------------------------------------------------------------------------*
''<Description>
''OK button has been clicked.
''</Description>
Private Sub cmdOK_Click()
	mSubmitted = True
	Me.Hide
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* optServiceNewService_Change																						*
'*------------------------------------------------------------------------*
''<Description>
''The value of the 'Add a new service definition with the following rate'
'' option has changed.
''</Description>
Private Sub optServiceNewService_Change()
	If optServiceNewService.Value = True Then
		cmboServiceOtherContact.Enabled = False
		cmboServiceOtherService.Enabled = False
		lblServiceNewServiceCommission.Enabled = True
		txtServiceNewServiceCommission.Enabled = True
		lblServiceNewServiceRate.Enabled = True
		txtServiceNewServiceRate.Enabled = True
	End If
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* optServiceOtherContact_Change																					*
'*------------------------------------------------------------------------*
''<Description>
''The value of the 'Select service of the same name under another contact.'
'' option has changed.
''</Description>
Private Sub optServiceOtherContact_Change()
	If optServiceOtherContact.Value = True Then
		cmboServiceOtherService.Enabled = False
		lblServiceNewServiceCommission.Enabled = False
		txtServiceNewServiceCommission.Enabled = False
		lblServiceNewServiceRate.Enabled = False
		txtServiceNewServiceRate.Enabled = False
		cmboServiceOtherContact.Enabled = True
	End If
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* optServiceOtherService_Change																					*
'*------------------------------------------------------------------------*
''<Description>
''The value of the 'Select another service that is already defined for
'' this contact.' option has changed.
''</Description>
Private Sub optServiceOtherService_Change()
'The other service option has changed.
	If optServiceOtherService.Value = True Then
		cmboServiceOtherContact.Enabled = False
		lblServiceNewServiceCommission.Enabled = False
		txtServiceNewServiceCommission.Enabled = False
		lblServiceNewServiceRate.Enabled = False
		txtServiceNewServiceRate.Enabled = False
		cmboServiceOtherService.Enabled = True
	End If
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UserForm_Initialize																										*
'*------------------------------------------------------------------------*
''<Description>
''The user form is being initialized.
''</Description>
Private Sub UserForm_Initialize()

	mContacts.FillFromSheet
	mServices.FillFromSheet mContacts
	'The form is cancelled by default unless OK is clicked.
	mSubmitted = False

End Sub
'*------------------------------------------------------------------------*

'*** PUBLIC ***
'*------------------------------------------------------------------------*
'* ChangeContact																													*
'*------------------------------------------------------------------------*
''<Description>
''Get a value indicating whether the contact will be changed.
''</Description>
Public Property Get ChangeContact() As Boolean
	ChangeContact = optServiceOtherContact.Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ChangeContactCode																											*
'*------------------------------------------------------------------------*
''<Description>
''Get the new contact code to which the references will be changed.
''</Description>
Public Property Get ChangeContactCode() As String
Dim cn As String    'Contact name.
Dim rv As String    'Return value.
Dim sa() As String  'String array.

	rv = ""
	If optServiceOtherContact.Value = True And _
		cmboServiceOtherContact.ListIndex > -1 Then
		sa = Split(cmboServiceOtherContact, ",")
		cn = Trim(sa(0))
		rv = mContacts.GetCodeFromName(cn)
	End If
	ChangeContactCode = rv

End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ChangeService																													*
'*------------------------------------------------------------------------*
''<Description>
''Get a value indicating whether the service will be changed.
''</Description>
Public Property Get ChangeService() As Boolean
	ChangeService = optServiceOtherService.Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ChangeServiceName																											*
'*------------------------------------------------------------------------*
''<Description>
''Get the new service name to which the references will be changed.
''</Description>
Public Property Get ChangeServiceName() As String
Dim rv As String    'Return value.
Dim sa() As String  'String array.

	rv = ""
	If optServiceOtherService.Value = True And _
		cmboServiceOtherService.ListIndex > -1 Then
		sa = Split(cmboServiceOtherService.Text, ",")
		rv = Trim(sa(0))
	End If
	ChangeServiceName = rv

End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* IsSubmitted																														*
'*------------------------------------------------------------------------*
''<Description>
''Get a value indicating whether the form was submitted.
''</Description>
Public Property Get IsSubmitted() As Boolean
	IsSubmitted = mSubmitted
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* NewCommissionValue																											*
'*------------------------------------------------------------------------*
''<Description>
''Get the new commission to be used in the new service.
''</Description>
Public Property Get NewCommissionValue() As Double
Dim rv As Double  'Return value.
Dim ws As String  'Working string.

	rv = 0#
	ws = Replace(Replace( _
		txtServiceNewServiceCommission.Text, "$", ""), ",", "")
	If optServiceNewService.Value = True And _
		IsNumeric(ws) Then
		rv = CDbl(ws)
	End If
	NewCommissionValue = rv

End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* NewRateValue																														*
'*------------------------------------------------------------------------*
''<Description>
''Get the new rate to be used in the new service.
''</Description>
Public Property Get NewRateValue() As Double
Dim rv As Double  'Return value.
Dim ws As String  'Working string.

	rv = 0#
	ws = Replace(Replace(txtServiceNewServiceRate.Text, "$", ""), ",", "")
	If optServiceNewService.Value = True And _
		IsNumeric(ws) Then
		rv = CDbl(ws)
	End If
	NewRateValue = rv

End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* NewService																															*
'*------------------------------------------------------------------------*
''<Description>
''Get a value indicating whether to create a new service with the
'' specified rate.
''</Description>
Public Property Get NewService() As Boolean
Dim rv As Boolean   'Return value.

	If optServiceNewService.Value = True Then
		rv = True
	End If
	NewService = rv

End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* SetContactService																											*
'*------------------------------------------------------------------------*
''<Description>
''Set the selected contact by its code, and related service.
''</Description>
''<Param name="ContactCode">
''The contact code to select.
''</Param>
''<Param name="ServiceName">
''Name of the service to select.
''</Param>
Public Sub SetContactService(ContactCode As String, ServiceName As String)
Dim cc As ContactCollection
Dim ci As ContactItem
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim sc As ContactServiceCollection
Dim si As ContactServiceItem

	cmboServiceOtherContact.Clear
	cmboServiceOtherService.Clear

	Set mContactSelected = mContacts.Code(ContactCode)
	If Not mContactSelected Is Nothing Then
		txtContact.Text = mContactSelected.BillToName
		'Get the other contacts for which this service name is defined.
		Set cc = mServices.ContactsWithService(ServiceName)
		lc = cc.Count()
		If lc > 0 Then
			For lp = 1 To lc
				Set ci = cc.Item(lp)
				Set si = mServices.ContactCodeService(ci.Code, ServiceName)
				If Not si Is Nothing Then
					'Add the contact and rate to the drop-down list.
					cmboServiceOtherContact.AddItem ci.BillToName & ", " & _
						Format(si.Rate, "0.00")
				End If
			Next lp
			If cmboServiceOtherContact.ListCount > 0 Then
				optServiceOtherContact.Enabled = True
			End If
		End If
		'Get the other services for the present contact.
		Set sc = mServices.ContactCodeItems(ContactCode)
		lc = sc.Count()
		If lc > 0 Then
			For lp = 1 To lc
				Set si = sc.Item(lp)
				cmboServiceOtherService.AddItem si.ServiceName & ", " & _
					Format(si.Rate, "0.00")
			Next lp
			If cmboServiceOtherService.ListCount > 0 Then
				optServiceOtherService.Enabled = True
			End If
		End If
	Else
		txtContact.Text = ContactCode & " (Unknown contact)"
	End If
	Set mServiceSelected = Nothing
	txtService.Text = ServiceName

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* SetHourlyRates																													*
'*------------------------------------------------------------------------*
''<Description>
''Set the hourly rate for the default selection.
''</Description>
''<Param name="Rate">
''The hourly rate to select.
''</Param>
''<Param name="Commission">
''The commission percentage.
''</Param>
Public Sub SetHourlyRates(Rate As Double, Commission As Double)
	mRate = Rate
	mCommission = Commission
	
	txtServiceNewServiceRate.Text = Format(mRate, "0.00")
	txtServiceNewServiceCommission.Text = Format(mCommission, "0.00")

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* SetMonthSheet																													*
'*------------------------------------------------------------------------*
''<Description>
''Set the month sheet to be displayed on the dialog.
''</Description>
''<Param name="MonthName">
''Name of the month to select.
''</Param>
Public Sub SetMonthSheet(MonthName As String)
	lblInstruction1.Caption = _
		Replace(lblInstruction1.Caption, "{Month}", MonthName)
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* UpdateSheet																														*
'*------------------------------------------------------------------------*
''<Description>
''Get a value indicating whether the user wishes to update the underlying
'' data sheets with effects of this change.
''</Description>
Public Property Get UpdateSheet() As Boolean
	UpdateSheet = chkUpdateSheet.Value
End Property
'*------------------------------------------------------------------------*
