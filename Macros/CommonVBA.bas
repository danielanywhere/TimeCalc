Attribute VB_Name = "CommonVBA"
Option Explicit
'CommonVBA.bas
' VBA methods suitable for use in every Microsoft Office environment.
'
' Copyright (c). 2000 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

'Clipboard Features...
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
'/Clipboard Features...

'*------------------------------------------------------------------------*
'* CharToCol																															*
'*------------------------------------------------------------------------*
''<Description>
''Return the 1-based column index of the column specified by letter.
''</Description>
''<Param name="Name">
''Column name to translate.
''</Param>
''<Returns>
''1-based column index.
''</Returns>
Public Function CharToCol(Name As String) As Integer
Dim cc As Integer   'Character count.
Dim ch As String    'Current character.
Dim cp As Integer   'Character position.
Dim ml As Integer   'Multiplier.
Dim rv As Integer   'Return value.
Dim tu As String    'Upper case character.

	tu = UCase(Name)
	ml = 1
	cc = Len(Name)
	For cp = cc To 1 Step -1
		ch = Mid(tu, cp, 1)
		rv = rv + ((Asc(ch) - 64) * ml)
		ml = ml * 26
	Next cp
	CharToCol = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ClipBoard_SetData																											*
'*------------------------------------------------------------------------*
''<Description>
''Copy the caller's text to the clipboard.
''</Description>
Public Function ClipBoard_SetData(Value As String)
Dim hc As Long	'Clip handle.
Dim hm As Long	'Memory handle.
Dim pm As Long	'Memory pointer.
Dim px As Long	'Clipboard empty pointer.

	' Allocate moveable global memory.
	'-------------------------------------------
	hm = GlobalAlloc(GHND, Len(Value) + 1)

	' Lock the block to get a far pointer
	' to this memory.
	pm = GlobalLock(hm)

	' Copy the string to this global memory.
	pm = lstrcpy(pm, Value)

	' Unlock the memory.
	If GlobalUnlock(hm) <> 0 Then
		MsgBox "Could not unlock memory location. Copy aborted."
	Else
		' Open the Clipboard to copy data to.
		If OpenClipboard(0&) = 0 Then
			MsgBox "Could not open the Clipboard. Copy aborted."
		Else
			' Clear the Clipboard.
			px = EmptyClipboard()
			' Copy the data to the Clipboard.
			hc = SetClipboardData(CF_TEXT, hm)
			If CloseClipboard() = 0 Then
				MsgBox "Could not close Clipboard."
			End If
		End If
	End If

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ColToChar																															*
'*------------------------------------------------------------------------*
''<Description>
''Given the column index, return the letter name of that column.
''</Description>
''<Param name="Index">
''1-based column index to translate.
''</Param>
''<Returns>
''Column letter.
''</Returns>
Public Function ColToChar(Index As Variant) As String
Dim bp As Boolean 'Flag - Prefix Present.
Dim cp As String  'Char Prefix.
Dim cs As String  'Char Suffix.
Dim id As Long    'Working Index.
Dim ip As Long    'Prefix Index.
Dim rs As String  'Return String.

	rs = ""
	bp = False
	cp = ""
	cs = ""
	id = Index
	ip = Int(id / 26)
	If ip > 0 And id Mod 26 = 0 Then
		ip = ip - 1
		id = 26
	End If
	If id Mod 26 <> 0 Then
		id = id Mod 26
	End If
	If ip > 0 Then
		'Prefix was found.
		rs = Chr(ip + 64)
	End If
	rs = rs & Chr(id + 64)
	ColToChar = rs
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* NewGuid																																*
'*------------------------------------------------------------------------*
''<Description>
''Create and return a new globally unique identifier to the caller.
''</Description>
''<Returns>
''Newly created and properly terminated GUID string.
''</Returns>
Public Function NewGuid() As String
Dim rs As String

	rs = CStr(CreateObject("Scriptlet.TypeLib").guid)
	rs = Left(rs, 38)
	NewGuid = rs

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* PadLeft																																*
'*------------------------------------------------------------------------*
''<Description>
''Pad characters to the left of a string shorter than the specified total
'' width.
''</Description>
''<Param name="Value">
''Value to be padded.
''</Param>
''<Param name="TotalWidth">
''Total width of the finished value, in characters.
''</Param>
''<Param name="Char">
''The character to prepend to the string while padding.
''</Param>
''<Returns>
''Caller's string, padded to the left with the specified character where
'' necessary.
''</Returns>
Public Function PadLeft(Value As String, TotalWidth As Integer, _
	Char As String) As String
Dim rv As String  'Return Value.

	rv = Value
	If Len(Char) > 0 Then
		Do While Len(rv) < TotalWidth
			rv = Char & rv
		Loop
	End If
	PadLeft = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* RegExMatches																														*
'*------------------------------------------------------------------------*
''<Description>
''Return a collection of regular expression matches for the specified
'' pattern in the given source.
''</Description>
''<Param name="Source">
''The string to be searched.
''</Param>
''<Param name="Pattern">
''The regular expression pattern to process.
''</Param>
''<Returns>
''Match collection.
''</Returns>
Public Function RegExMatches(Source As String, Pattern As String) As Object
Dim mc As Object
Dim rx As Object

	Set rx = CreateObject("VBScript.RegExp")
	rx.IgnoreCase = True
	rx.Pattern = Pattern
	rx.Global = True    'Set global applicability.

	Set mc = rx.Execute(Source)  ' Execute search.
'  For Each Match In Matches   ' Iterate Matches collection.
'    RetStr = RetStr & "Match found at position "
'    RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
'    RetStr = RetStr & Match.Value & "'." & vbCrLf
'  Next
	Set RegExMatches = mc

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* RegExReplace																														*
'*------------------------------------------------------------------------*
''<Description>
''Replace patterns found in the source string with the replacement pattern.
''</Description>
''<Param name="Source">
''Source string to be inspected.
''</Param>
''<Param name="Find">
''Pattern to find.
''</Param>
''<Param name="Replace">
''Pattern to replace.
''</Param>
''<Returns>
''The caller's string with suitable replacements made.
''</Returns>
Public Function RegExReplace(Source As String, _
	Find As String, Replace As String) As String
Dim rv As String  'Return Value.
Dim rx As Object  'RegExp

	Set rx = CreateObject("VBScript.RegExp")
	rx.IgnoreCase = True
	rx.Pattern = Find
	rv = rx.Replace(Source, Replace)
	RegExReplace = rv

End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* Repeat																																	*
'*------------------------------------------------------------------------*
''<Description>
''Repeat the provided pattern a specified number of times.
''</Description>
''<Param name="Character">
''The character to repeat.
''</Param>
''<Param name="Count">
''Number of times to repeat the character.
''</Param>
''<Returns>
''String of specific character in the specified length.
''</Returns>
Public Function Repeat(Character As String, Count As Integer) As String
	Repeat = Replace(Space(Count), " ", Character)
End Function
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ToTitleCase																														*
'*------------------------------------------------------------------------*
''<Description>
''Convert the caller's value to title case, where the first letter of every
'' word is capitalized, and all non-alphanumeric characters are removed.
''</Description>
''<Param name="Value">
''Value to be converted.
''</Param>
''<Returns>
''Original string, converted to title case.
''</Returns>
Public Function ToTitleCase(Value As String) As String
Dim bh As Boolean     'Flag - Handled.
Dim la                'Value Array.
Dim lc As Integer     'List Count.
Dim lp As Integer     'List Position.
Dim rs As String      'Return String.
Dim tl As String      'Lower Case Working String.
Dim ts As String      'Temporary String.
Dim ws As String      'Working String.

	rs = ""
	la = Split(Value, " ")
	lc = UBound(la)
	If lc > 0 Then
		For lp = 0 To lc
			bh = False
			ws = la(lp)
			tl = LCase(ws)
			If Len(ws) > 0 Then
				If Len(rs) > 0 Then
					rs = rs & " "
				End If
				If Len(ws) = 2 Then
					'2 character strings.
					Select Case tl
						Case "ii", "iv", "vi", "po":
							rs = rs & UCase(ws)
							bh = True
					End Select
				ElseIf Len(ws) = 3 Then
					'3 character strings.
					Select Case tl
						Case "iii", "vii", "llc", "dba", "c/o":
							rs = rs & UCase(ws)
							bh = True
					End Select
				ElseIf Len(ws) = 4 Then
					'4 character strings.
					Select Case tl
						Case "p.o.":
							rs = rs & UCase(ws)
							bh = True
					End Select
				Else
					'Other variations.
					If Left(tl, 2) = "o'" Then
						rs = rs & "O'" & UCase(Mid(ws, 3, 1)) & Right(tl, Len(tl) - 3)
						bh = True
					End If
					If Left(tl, 2) = "mc" Then
						rs = rs & "Mc" & UCase(Mid(ws, 3, 1)) & Right(tl, Len(tl) - 3)
						bh = True
					End If
					If InStr(tl, "-") > 0 Then
						rs = rs & Replace(ToTitleCase(Replace(ws, "-", " ")), " ", "-")
						bh = True
					End If
				End If
				If bh = False Then
					rs = rs & UCase(Left(ws, 1)) & LCase(Right(ws, Len(ws) - 1))
				End If
			End If
		Next lp
	Else
		If Len(Value) > 0 Then
			rs = UCase(Left(Value, 1)) & LCase(Right(Value, Len(Value) - 1))
		End If
	End If
	ToTitleCase = rs

End Function
'*------------------------------------------------------------------------*
