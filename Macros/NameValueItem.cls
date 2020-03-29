VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "NameValueItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NameValueItem.cls
' Generic name and value string item.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mName As String
Private mValue As String

'*------------------------------------------------------------------------*
'* Name																																		*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the name of the item.
''</Description>
Public Property Get Name() As String
	Name = mName
End Property
Public Property Let Name(ItemValue As String)
	mName = ItemValue
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Value																																	*
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the value of the item.
''</Description>
Public Property Get Value() As String
	Value = mValue
End Property
Public Property Let Value(ItemValue As String)
	mValue = ItemValue
End Property
'*------------------------------------------------------------------------*
