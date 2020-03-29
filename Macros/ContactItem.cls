VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "ContactItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ContactItem.cls
' Collection of ContactItems.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mBillToAddress As String
Private mBillToCityStateZip As String
Private mBillToName As String
Private mCode As String
Private mShipToAddress As String
Private mShipToCityStateZip As String
Private mShipToName As String

'*------------------------------------------------------------------------*
'* BillToAddress																													*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the bill-to address.
''</Description>
Public Property Get BillToAddress() As String
	BillToAddress = mBillToAddress
End Property
Public Property Let BillToAddress(Value As String)
	mBillToAddress = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* BillToCityStateZip																											*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the bill-to city, state, and zip code.
''</Description>
Public Property Get BillToCityStateZip() As String
	BillToCityStateZip = mBillToCityStateZip
End Property
Public Property Let BillToCityStateZip(Value As String)
	mBillToCityStateZip = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* BillToName																															*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the bill-to name of the contact.
''</Description>
Public Property Get BillToName() As String
	BillToName = mBillToName
End Property
Public Property Let BillToName(Value As String)
	mBillToName = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Code																																		*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the contact code.
''</Description>
Public Property Get Code() As String
	Code = mCode
End Property
Public Property Let Code(Value As String)
	mCode = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ShipToAddress																													*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the ship-to address.
''</Description>
Public Property Get ShipToAddress() As String
	ShipToAddress = mShipToAddress
End Property
Public Property Let ShipToAddress(Value As String)
	mShipToAddress = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ShipToCityStateZip																											*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the ship-to city, state, and zip code.
''</Description>
Public Property Get ShipToCityStateZip() As String
	ShipToCityStateZip = mShipToCityStateZip
End Property
Public Property Let ShipToCityStateZip(Value As String)
	mShipToCityStateZip = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ShipToName																															*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the ship-to contact name.
''</Description>
Public Property Get ShipToName() As String
	ShipToName = mShipToName
End Property
Public Property Let ShipToName(Value As String)
	mShipToName = Value
End Property
'*------------------------------------------------------------------------*
