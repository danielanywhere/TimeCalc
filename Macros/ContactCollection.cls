VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "ContactCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ContactCollection.cls
' Collection of ContactItems.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mCol As New Collection

'*------------------------------------------------------------------------*
'* Add																																		*
'*------------------------------------------------------------------------*
''<Description>
''Add an item to to the collection.
''</Description>
''<Param name="Contact">
''Reference to the contact item to be added.
''</Param>
Public Sub Add(Contact As ContactItem)
	mCol.Add Contact
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* BillToName																															*
'*------------------------------------------------------------------------*
''<Description>
''Return a contact by bill-to name.
''</Description>
''<Param name="ContactName">
''Name of the contact to find.
''</Param>
''<Returns>
''Reference to the contact item matching the specified bill-to name.
''</Returns>
Public Function BillToName(ContactName As String) As ContactItem
Dim ci As ContactItem
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim rv As ContactItem 'Return value.

	Set rv = Nothing
	lc = mCol.Count
	For lp = 1 To lc
		Set ci = mCol.Item(lp)
		If ci.BillToName = ContactName Then
			Set rv = ci
			Exit For
		End If
	Next lp
	Set BillToName = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Code																																		*
'*------------------------------------------------------------------------*
''<Description>
''Return a contact item by code.
''</Description>
''<Param name="ContactCode">
''The contact code to find.
''</Param>
''<Returns>
''Reference to the contact item having the specified contact code.
''</Returns>
Public Function Code(ContactCode As String) As ContactItem
'Return a reference to the item with the specified contact code.
Dim ci As ContactItem
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim rv As ContactItem 'Return value.

	Set rv = Nothing
	lc = mCol.Count
	For lp = 1 To lc
		Set ci = mCol.Item(lp)
		If ci.Code = ContactCode Then
			Set rv = ci
			Exit For
		End If
	Next lp
	Set Code = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Count																																	*
'*------------------------------------------------------------------------*
''<Description>
''Return the count of items in the collection.
''</Description>
''<Returns>
''Count of items currently loaded in the collection.
''</Returns>
Public Function Count() As Integer
	Count = mCol.Count
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* FillFromSheet																													*
'*------------------------------------------------------------------------*
''<Description>
''Fill the collection directly from the Contacts table.
''</Description>
''<Returns>
''Value indicating whether the operation was a success.
''</Returns>
Public Function FillFromSheet() As Boolean
'Fill the collection from the source sheet.
Dim cc As Integer     'Column count.
Dim ci As ContactItem 'Current contact.
Dim cn() As String    'Column names.
Dim cp As Integer     'Column position.
Dim ct As String      'Column title.
Dim lc As Integer     'List count.
Dim lp As Integer     'List position.
Dim rg As Range       'Working range.
Dim rr As Range       'Row range.
Dim rv As Boolean     'Return value.
Dim sh As Worksheet   'Working sheet.
Dim ws As String      'Working string.

	rv = True
	Set sh = Sheets("Contacts")
	'Get the names of the columns from the header row.
	Set rg = sh.ListObjects("Contacts").HeaderRowRange
	cc = rg.Columns.Count
	If cc > 0 Then
		'Columns are found.
		ReDim cn(cc, 1)
		For cp = 1 To cc
			cn(cp, 0) = rg.Cells(1, cp).Value
			cn(cp, 1) = ColToChar(cp)
		Next cp
	End If
	'Add the data rows to the collection.
	Set rg = sh.ListObjects("Contacts").DataBodyRange
	lc = rg.Rows.Count
	cc = rg.Columns.Count
	For lp = 1 To lc
		'Each row.
		Set ci = New ContactItem
		For cp = 1 To cc
			'Each column.
			ct = cn(cp, 0)
			ws = rg.Cells(lp, cp).Value
			Select Case ct
				Case "Bill To Address":
					ci.BillToAddress = ws
				Case "Bill To City State Zip":
					ci.BillToCityStateZip = ws
				Case "Bill To Name":
					ci.BillToName = ws
				Case "Code":
					ci.Code = ws
				Case "Ship To Address":
					ci.ShipToAddress = ws
				Case "Ship To City State Zip":
					ci.ShipToCityStateZip = ws
				Case "Ship To Name":
					ci.ShipToName = ws
			End Select
		Next cp
		mCol.Add ci
	Next lp
	FillFromSheet = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetCodeFromName																												*
'*------------------------------------------------------------------------*
''<Description>
''Translate the contact name to contact code.
''</Description>
''<Param name="Name">
''Contact name to find.
''</Param>
''<Returns>
''Contact code.
''</Returns>
Public Function GetCodeFromName(Name As String) As String
Dim ci As ContactItem
Dim rv As String    'Return value.

	Set ci = BillToName(Name)
	If Not ci Is Nothing Then
		rv = ci.Code
	End If
	GetCodeFromName = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* GetCodeIndex																														*
'*------------------------------------------------------------------------*
''<Description>
''Return the location of the item within the collection matching the
'' specified contact code.
''</Description>
''<Param name="ContactCode">
''The contact code to match.
''</Param>
''<Returns>
''1-based index of the specified item within the collection, if found.
'' Otherwise, 0.
''</Returns>
Public Function GetCodeIndex(ContactCode As String) As Integer
Dim ci As ContactItem
Dim lc As Integer		'List count.
Dim lp As Integer		'List position.
Dim rv As Integer		'Return value.

	rv = 0
	lc = mCol.Count()
	For lp = 1 To lc
		Set ci = mCol.Item(lp)
		If ci.Code = ContactCode Then
			rv = lp
			Exit For
		End If
	Next lp
	GetCodeIndex = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Insert																																	*
'*------------------------------------------------------------------------*
''<Description>
''Insert an item at the specified location in the collection.
''</Description>
''<Param name="Index">
''Location at which to insert the item.
''</Param>
''<Param name="Value">
''Reference to the item to insert.
''</Param>
Public Sub Insert(Index As Variant, Value As ContactItem)
'Insert an item at a specified location in the collection.
	If IsNumeric(Index) Then
		mCol.Add Value, Before:=Index
	End If
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Item																																		*
'*------------------------------------------------------------------------*
''<Description>
''Return an item from the specified location in the collection.
''</Description>
''<Param name="Index">
''Index of the item to select.
''</Param>
''<Returns>
''Reference to the contact item found at the specified location.
''</Returns>
Public Function Item(Index As Variant) As ContactItem
Dim rv As ContactItem

	If IsNumeric(Index) Then
		Set rv = mCol(Index)
	End If
	Set Item = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Remove																																	*
'*------------------------------------------------------------------------*
''<Description>
''Remove an item from the specified location in the collection.
''</Description>
''<Param name="Index">
''1-based location at which the item will be removed.
''</Param>
Public Function Remove(Index As Variant)
	If IsNumeric(Index) Then
		mCol.Remove Index
	End If
End Function
'*------------------------------------------------------------------------*
