VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "LogItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'LogItem.cls
' Item collected for use on a log report.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mItemDateEnd As Date
Private mItemDateStart As Date
Private mManHoursSpent As Double
Private mProjectName As String
Private mRowIndex As Integer
Private mSheetName As String
Private mTaskName As String

'*** PRIVATE ***
'*------------------------------------------------------------------------*
'* Class_Initialize																												*
'*------------------------------------------------------------------------*
''<Description>
''The object is initializing.
''</Description>
Private Sub Class_Initialize()
	mItemDateEnd = Now
	mItemDateStart = Now
	mManHoursSpent = 0
	mProjectName = ""
	mTaskName = ""
End Sub
'*------------------------------------------------------------------------*

'*** PUBLIC ***
'*------------------------------------------------------------------------*
'* ItemDateEnd                                                            *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the end date of the entry.
''</Description>
Public Property Get ItemDateEnd() As Date
	ItemDateEnd = mItemDateEnd
End Property
Public Property Let ItemDateEnd(Value As Date)
	mItemDateEnd = Value
End Property
'*------------------------------------------------------------------------*
 
'*------------------------------------------------------------------------*
'* ItemDateStart                                                          *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the start date of the entry.
''</Description>
Public Property Get ItemDateStart() As Date
	ItemDateStart = mItemDateStart
End Property
Public Property Let ItemDateStart(Value As Date)
	mItemDateStart = Value
End Property
'*------------------------------------------------------------------------*
 
'*------------------------------------------------------------------------*
'* ManHoursSpent                                                          *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the number of man-hours elapsed on the entry.
''</Description>
Public Property Get ManHoursSpent() As Double
	ManHoursSpent = mManHoursSpent
End Property
Public Property Let ManHoursSpent(Value As Double)
	mManHoursSpent = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ProjectName                                                            *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the name of the project.
''</Description>
Public Property Get ProjectName() As String
	ProjectName = mProjectName
End Property
Public Property Let ProjectName(Value As String)
	mProjectName = Value
End Property
'*------------------------------------------------------------------------*
 
'*------------------------------------------------------------------------*
'* RowIndex                                                               *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the index of the spreadsheet row where the base entry is located.
''</Description>
Public Property Get RowIndex() As Integer
	RowIndex = mRowIndex
End Property
Public Property Let RowIndex(Value As Integer)
	mRowIndex = Value
End Property
'*------------------------------------------------------------------------*
 
'*------------------------------------------------------------------------*
'* SheetName                                                              *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the name of the sheet.
''</Description>
Public Property Get SheetName() As String
	SheetName = mSheetName
End Property
Public Property Let SheetName(Value As String)
	mSheetName = Value
End Property
'*------------------------------------------------------------------------*
 
'*------------------------------------------------------------------------*
'* TaskName                                                               *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the name of the task.
''</Description>
Public Property Get TaskName() As String
	TaskName = mTaskName
End Property
Public Property Let TaskName(Value As String)
	mTaskName = Value
End Property
'*------------------------------------------------------------------------*
