VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "ContactServiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ContactServiceItem.cls
' An item that relates contact to service.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mCommission As Double
Private mContact As ContactItem
Private mRate As Double
Private mServiceName As String

'*** PRIVATE ***
'*------------------------------------------------------------------------*
'* Class_Initialize																												*
'*------------------------------------------------------------------------*
''<Description>
''Object is initializing.
''</Description>
Private Sub Class_Initialize()
	Set mContact = Nothing
End Sub
'*------------------------------------------------------------------------*

'*** PUBLIC ***
'*------------------------------------------------------------------------*
'* Commission																															*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the percentage of commission to pay out for all work on this
'' service.
''</Description>
Public Property Get Commission() As Double
	Commission = mCommission
End Property
Public Property Let Commission(Value As Double)
	mCommission = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Contact																																*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set a reference to the Contact to which this service applies.
''</Description>
Public Property Get Contact() As ContactItem
	Set Contact = mContact
End Property
Public Property Set Contact(Value As ContactItem)
	Set mContact = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetListTitle																														*
'*------------------------------------------------------------------------*
''<Description>
''Return the services list title for this entry.
''</Description>
''<Returns>
''A list title if this item, in the format of ContactCode_ServiceName.
''</Returns>
Public Function GetListTitle() As String
Dim rv As String    'Return value.

	If Not mContact Is Nothing Then
		rv = mContact.Code & "_" & mServiceName
	End If
	GetListTitle = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Rate																																		*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the hourly rate charged for this service and customer.
''</Description>
Public Property Get Rate() As Double
	Rate = mRate
End Property
Public Property Let Rate(Value As Double)
	mRate = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ServiceName																														*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the name of the service.
''</Description>
Public Property Get ServiceName() As String
	ServiceName = mServiceName
End Property
Public Property Let ServiceName(Value As String)
	mServiceName = Value
End Property
'*------------------------------------------------------------------------*
