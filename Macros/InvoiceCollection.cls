VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "InvoiceCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'InvoiceCollection.cls
' Collection of invoice items.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mCol As New Collection

'*------------------------------------------------------------------------*
'* Add																																		*
'*------------------------------------------------------------------------*
''<Description>
''Add an item to the collection.
''</Description>
''<Param name="Value">
''Item to add.
''</Param>
Public Sub Add(Value As InvoiceItem)
	mCol.Add Value
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ContactCodeExists																											*
'*------------------------------------------------------------------------*
''<Description>
''Return a value indicating whether the specified contact code is found in
'' this instance.
''</Description>
''<Param name="ContactCode">
''The contact code to search for.
''</Param>
''<Returns>
''Value indicating whether the specified contact was found.
''</Returns>
Public Function ContactCodeExists(ContactCode As String) As Boolean
Dim ii As InvoiceItem
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim rv As Boolean   'Return value.

	rv = False
	lc = mCol.Count()
	For lp = 1 To lc
		Set ii = mCol.Item(lp)
		If ii.ContactCode = ContactCode Then
			rv = True
			Exit For
		End If
	Next lp
	ContactCodeExists = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Count																																	*
'*------------------------------------------------------------------------*
''<Description>
''Return the count of items in this collection.
''</Description>
''<Returns>
''Count of items currently loaded in the collection.
''</Returns>
Public Function Count() As Long
	Count = mCol.Count
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* FirstContactCodeServiceName																						*
'*------------------------------------------------------------------------*
''<Description>
''Return the first item found with the specified contact code and service
'' item.
''</Description>
''<Param name="ContactCode">
''Contact code to search for.
''</Param>
''<Param name="ServiceName">
''Service name to match.
''</Param>
''<Returns>
''Reference to the matching invoice item, if found. Otherwise, nothing.
''</Returns>
Public Function FirstContactCodeServiceName( _
	ContactCode As String, ServiceName As String) As InvoiceItem
Dim ii As InvoiceItem
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.
Dim rv As InvoiceItem

	Set rv = Nothing
	lc = mCol.Count()
	For lp = 1 To lc
		Set ii = mCol.Item(lp)
		If ii.ContactCode = ContactCode And ii.ServiceName = ServiceName Then
			Set rv = ii
			Exit For
		End If
	Next lp
	Set FirstContactCodeServiceName = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetServiceList																													*
'*------------------------------------------------------------------------*
''<Description>
''Return a list of names found on the service list.
''</Description>
''<Returns>
''Raw collection of strings equal to the unique, sorted group of service
'' names found.
''</Returns>
Public Function GetServiceList() As Collection
Dim bf As Boolean   'Flag - item found.
Dim ii As InvoiceItem
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim ls As String    'Lower case string.
Dim rc As Integer   'Result count.
Dim rp As Integer   'Result position.
Dim rv As New Collection
Dim ws As String    'Working string.

	rc = 0
	lc = mCol.Count
	For lp = 1 To lc
		Set ii = mCol.Item(lp)
		ws = ii.ServiceName
		bf = False
		For rp = 1 To rc
			If rv.Item(rp) = ws Then
				bf = True
				Exit For
			End If
		Next rp
		If bf = False Then
			'Add the service to the list, sorted.
			ls = LCase(ws)
			For rp = 1 To rc
				If LCase(rv.Item(rp)) > ls Then
					rv.Add ws, Before:=rp
					rc = rc + 1
					bf = True
					Exit For
				End If
			Next rp
			If bf = False Then
				rv.Add ws
				rc = rc + 1
			End If
		End If
	Next lp
	Set GetServiceList = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Item																																		*
'*------------------------------------------------------------------------*
''<Description>
''Return the item from the specified location in the collection.
''</Description>
''<Param name="Index">
''Physical 1-based location of the item to retrieve.
''</Param>
''<Returns>
''Reference to the item found at the specified location.
''</Returns>
Public Function Item(Index As Variant) As InvoiceItem
	Set Item = mCol.Item(Index)
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Remove																																	*
'*------------------------------------------------------------------------*
''<Description>
''Remove the item from the specified location in the collection.
''</Description>
''<Param name="Index">
''1-based location of the item to remove.
''</Param>
Public Sub Remove(Index As Variant)
	mCol.Remove Index
End Sub
'*------------------------------------------------------------------------*
