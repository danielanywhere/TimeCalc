VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "LogCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'LogCollection.cls
' Collection of items collected for use on a log report.
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
''Reference to the log item to be added.
''</Param>
Public Sub Add(Value As LogItem)
	mCol.Add Value
End Sub
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Count																																	*
'*------------------------------------------------------------------------*
''<Description>
''Return a count of the items in this collection.
''</Description>
''<Returns>
''Count of items in the collection.
''</Returns>
Public Function Count() As Long
	Count = mCol.Count
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Item																																		*
'*------------------------------------------------------------------------*
''<Description>
''Return an item from the collection by its ordinal position.
''</Description>
''<Param name="Index">
''The 1-based ordinal position of the item in the collection.
''</Param>
''<Returns>
''Reference to the item found at the specified location.
''</Returns>
Public Function Item(Index As Variant) As LogItem
	Set Item = mCol.Item(Index)
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Remove																																	*
'*------------------------------------------------------------------------*
''<Description>
''Remove an item from the specified location in the collection.
''</Description>
''<Param name="Index">
''1-based location at which to remove the item.
''</Param>
Public Sub Remove(Index As Variant)
	mCol.Remove Index
End Sub
'*------------------------------------------------------------------------*
