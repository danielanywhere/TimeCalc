VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "NameValueCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'NameValueCollection.cls
' Collection of generic name and value string entries.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mCol As New Collection

'*------------------------------------------------------------------------*
'* Add																																		*
'*------------------------------------------------------------------------*
''<Description>
''Add an item by member values.
''</Description>
''<Param name="Name">
''Name of the entry.
''</Param>
''<Param name="Value">
''Value of the entry.
''</Param>
Public Sub Add(Name As String, Value As String)
Dim ni As New NameValueItem

	ni.Name = Name
	ni.Value = Value
	mCol.Add ni

End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* AddItem																																*
'*------------------------------------------------------------------------*
''<Description>
''Add an existing item to the collection.
''</Description>
''<Param name="Item">
''Reference to the item to be added.
''</Param>
Public Sub AddItem(Item As NameValueItem)
	mCol.Add Item
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Count																																	*
'*------------------------------------------------------------------------*
''<Description>
''Return the count of items in the collection.
''</Description>
''<Returns>
''Count of items in the collection.
''</Returns>
Public Function Count() As Integer
	Count = mCol.Count
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetValue																																*
'*------------------------------------------------------------------------*
''<Description>
''Return the value of the item corresponding to the specified name.
''</Description>
''<Param name="Name">
''Name of the item to find.
''</Param>
''<Returns>
''The value of the item with the specified name, if found.
'' Otherwise, empty string.
''</Returns>
Public Function GetValue(Name As String) As String
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim ni As NameValueItem
Dim rv As String    'Return value.

	rv = ""
	lc = mCol.Count
	For lp = 1 To lc
		Set ni = mCol.Item(lp)
		If ni.Name = Name Then
			rv = ni.Value
			Exit For
		End If
	Next lp
	GetValue = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Item																																		*
'*------------------------------------------------------------------------*
''<Description>
''Return the item at the specified ordinal location in the collection.
''</Description>
''<Param name="Index">
''1-based index of the item to retrieve.
''</Param>
''<Returns>
''Item at the specified location within the collection, if found.
'' Otherwise, Nothing.
''</Returns>
Public Function Item(Index As Variant) As NameValueItem
Dim rv As NameValueItem

	If IsNumeric(Index) Then
		Set rv = mCol.Item(Index)
	End If
	Set Item = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Remove																																	*
'*------------------------------------------------------------------------*
''<Description>
''Remove the item at the ordinal location within the collection.
''</Description>
''<Param name="Index">
''1-based index of the item to be removed.
''</Param>
Public Sub Remove(Index As Variant)

	If IsNumeric(Index) Then
		mCol.Remove Index
	End If

End Sub
'*------------------------------------------------------------------------*
