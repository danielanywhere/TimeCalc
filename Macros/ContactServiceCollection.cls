VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "ContactServiceCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ContactServiceCollection.cls
' Collection of ContactServiceItems.
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
''<Param name="Service">
''Reference to the service item to be added.
''</Param>
Public Sub Add(Service As ContactServiceItem)
	mCol.Add Service
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ContactCodeItems																												*
'*------------------------------------------------------------------------*
''<Description>
''Return a collection of contact service items where all included items
'' are of the specified contact.
''</Description>
''<Param name="ContactCode">
''Contact code by which the collection will be filtered.
''</Param>
''<Returns>
''Collection of services for the specified contact.
''</Returns>
Public Function ContactCodeItems(ContactCode As String) _
	As ContactServiceCollection
'Return a collection of services for the specified contact.
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.
Dim rv As New ContactServiceCollection
Dim si As ContactServiceItem

	lc = mCol.Count()
	For lp = 1 To lc
		Set si = mCol.Item(lp)
		If Not si.Contact Is Nothing Then
			If si.Contact.Code = ContactCode Then
				rv.Add si
			End If
		End If
	Next lp
	Set ContactCodeItems = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ContactCodeService																											*
'*------------------------------------------------------------------------*
''<Description>
''Return the service item corresponding to the specified contact code and
'' service name.
''</Description>
''<Param name="ContactCode">
''Contact code to search for.
''</Param>
''<Param name="ServiceName">
''Name of the service to match.
''</Param>
''<Returns>
''Reference to the service item matching the specified contact code and
'' service name.
''</Returns>
Public Function ContactCodeService(ContactCode As String, _
	ServiceName As String) As ContactServiceItem
'Return a reference to the service instance matching the specified contact
' code and service name.
Dim lc As Integer 'List count.
Dim lp As Integer 'List position.
Dim rv As ContactServiceItem
Dim si As ContactServiceItem

	Set rv = Nothing
	lc = mCol.Count()
	For lp = 1 To lc
		Set si = mCol.Item(lp)
		If Not si.Contact Is Nothing Then
			If si.Contact.Code = ContactCode And si.ServiceName = ServiceName Then
				Set rv = si
				Exit For
			End If
		End If
	Next lp
	Set ContactCodeService = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ContactsWithService																										*
'*------------------------------------------------------------------------*
''<Description>
''Return a collection of contacts that are associated with the specified
'' service.
''</Description>
''<Param name="ServiceName">
''Name of the service to match.
''</Param>
''<Returns>
''Collection of contact records that each have a reference to the specified
'' service.
''</Returns>
Public Function ContactsWithService(ServiceName As String) _
	As ContactCollection
'Return a collection of contacts for which the specified service has been
' defined.
Dim bf As Boolean 'Flag - found.
Dim cc As Integer 'Contact count.
Dim ci As ContactItem
Dim cp As Integer 'Contact position.
Dim cr As String  'Contact code reference.
Dim cs As String  'Contact code selected.
Dim rv As New ContactCollection
Dim sc As Integer 'Service count.
Dim si As ContactServiceItem
Dim sp As Integer 'Service position.

	cc = 0
	sc = mCol.Count()
	For sp = 1 To sc
		Set si = mCol.Item(sp)
		If si.ServiceName = ServiceName Then
			'Service found.
			Set ci = si.Contact
			cr = ci.Code
			bf = False
			'Insert unique, sorted.
			For cp = 1 To cc
				cs = rv.Item(cp).Code
				If cr = cs Then
					bf = True
					Exit For
				ElseIf cr < cs Then
					rv.Insert cp, ci
					bf = True
					Exit For
				End If
			Next cp
			If bf = False Then
				'The item was not found. Append.
				rv.Add ci
			End If
		End If
	Next sp
	Set ContactsWithService = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Count																																	*
'*------------------------------------------------------------------------*
''<Description>
''Return a count of items in the collection.
''</Description>
''<Returns>
''Count of items in the collection.
''</Returns>
Public Function Count() As Integer
	Count = mCol.Count()
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* FillFromSheet																													*
'*------------------------------------------------------------------------*
''<Description>
''Fill the collection directly from the Services table.
''</Description>
''<Returns>
''Value indicating whether the operation was a success.
''</Returns>
Public Function FillFromSheet(Contacts As ContactCollection) As Boolean
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
Dim si As ContactServiceItem  'Current service item.
Dim ws As String      'Working string.

	rv = True
	Set sh = Sheets("Services")
	'Get the names of the columns from the header row.
	Set rg = sh.ListObjects("Services").HeaderRowRange
	cc = rg.Columns.Count
	If cc > 0 Then
		'Columns are found.
		ReDim cn(cc, 1)
		For cp = 1 To cc
			cn(cp, 0) = rg.Cells(1, cp).Value
'      cn(cp, 1) = ColToChar(cp)
			cn(cp, 1) = ColToChar(rg.Cells(1, cp).Column)
		Next cp
	End If
	'Add the data rows to the collection.
	Set rg = sh.ListObjects("Services").DataBodyRange
	lc = rg.Rows.Count
	cc = rg.Columns.Count
	For lp = 1 To lc
		'Each row.
		Set si = New ContactServiceItem
		For cp = 1 To cc
			'Each column.
			ct = cn(cp, 0)
			ws = rg.Cells(lp, cp).Value
			Select Case ct
				Case "Customer":
					Set si.Contact = Contacts.Code(ws)
				Case "Service":
					si.ServiceName = ws
				Case "Rate per hr":
					If IsNumeric(ws) Then
						si.Rate = CDbl(ws)
					End If
				Case "Commission":
					If IsNumeric(ws) Then
						si.Commission = CDbl(ws)
					End If
			End Select
		Next cp
		If Len(si.ServiceName) > 0 Then
			mCol.Add si
		End If
	Next lp
	FillFromSheet = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* IsDefinedContactCodeService																						*
'*------------------------------------------------------------------------*
''<Description>
''Return a value indicating whether the contact code and service name have
'' a defined relation in this collection.
''</Description>
''<Param name="ContactCode">
''Contact code to search for.
''</Param>
''<Param name="ServiceName">
''Name of the service to match.
''</Param>
''<Returns>
''True if the relation between contact code and service name is found.
'' Otherwise, false.
''</Returns>
Public Function IsDefinedContactCodeService( _
	ContactCode As String, ServiceName As String) As Boolean
'Return a value indicating whether the combination of contact code and
' service are defined.
Dim rv As Boolean 'Return value.
Dim si As ContactServiceItem

	Set si = ContactCodeService(ContactCode, ServiceName)
	If Not si Is Nothing Then
		rv = True
	End If
	IsDefinedContactCodeService = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Item																																		*
'*------------------------------------------------------------------------*
''<Description>
''Return the specified item from the collection.
''</Description>
''<Param name="Index">
''Either the specific numeric location of the item in the collection, or
'' a string with the syntax of ContactCode_ServiceName.
''</Param>
''<Returns>
''A contact service item matching the specification.
''</Returns>
Public Function Item(Index As Variant) As ContactServiceItem
Dim id As String    'Index value.
Dim lc As Integer   'List count.
Dim lp As Integer   'List position.
Dim rv As ContactServiceItem  'Return value.
Dim si As ContactServiceItem  'Current item.

	If IsNumeric(Index) Then
		Set rv = mCol.Item(Index)
	ElseIf Not IsNull(Index) Then
		id = CStr(Index)
		lc = mCol.Count
		For lp = 1 To lc
			Set si = mCol.Item(lp)
			If si.GetListTitle() = id Then
				Set rv = si
				Exit For
			End If
		Next lp
	End If
	Set Item = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Remove																																	*
'*------------------------------------------------------------------------*
''<Description>
''Remove the item from the specified location in the collection.
''</Description>
''<Param name="Index">
''Physical 1-based location in the collection from which to remove the
'' item.
''</Param>
Public Sub Remove(Index As Variant)
	If IsNumeric(Index) Then
		mCol.Remove Index
	End If
End Sub
'*------------------------------------------------------------------------*
