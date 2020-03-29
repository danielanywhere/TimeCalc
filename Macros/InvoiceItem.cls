VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "InvoiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'InvoiceItem.cls
' Item collected for use on an invoice.
'
' Copyright (c). 2008 - 2020 Daniel Patterson, MCSD (danielanywhere)
' Released for public access under the MIT License.
' http://www.opensource.org/licenses/mit-license.php

Private mContactCode As String
Private mInvoicedAmount As Double
Private mItemDate As Date
Private mManHoursSpent As Double
Private mProjectName As String
Private mRate As Double
Private mReceivedAmount As Double
Private mRowIndex As Integer
Private mSent As Boolean
Private mServiceName As String
Private mSheetName As String
Private mTaskName As String

'*** PRIVATE ***
'*------------------------------------------------------------------------*
'* Class_Initialize																												*
'*------------------------------------------------------------------------*
''<Description>
''The class is initializing.
''</Description>
Private Sub Class_Initialize()
	mContactCode = ""
	mInvoicedAmount = 0#
	mItemDate = Date
	mManHoursSpent = 0#
	mProjectName = ""
	mReceivedAmount = 0#
	mSent = False
	mServiceName = ""
	mTaskName = ""
	mRate = 0#
End Sub
'*------------------------------------------------------------------------*

'*** PUBLIC ***
'*------------------------------------------------------------------------*
'* ContactCode                                                            *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the contact code selected in the service cell.
''</Description>
Public Property Get ContactCode() As String
	ContactCode = mContactCode
End Property
Public Property Let ContactCode(Value As String)
	mContactCode = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* InvoicedAmount                                                         *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the amount invoiced for this item.
''</Description>
Public Property Get InvoicedAmount() As Double
	InvoicedAmount = mInvoicedAmount
End Property
Public Property Let InvoicedAmount(Value As Double)
	mInvoicedAmount = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ItemDate                                                               *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the date of the item.
''</Description>
Public Property Get ItemDate() As Date
	ItemDate = mItemDate
End Property
Public Property Let ItemDate(Value As Date)
	mItemDate = Value
End Property
'*------------------------------------------------------------------------*
 
'*------------------------------------------------------------------------*
'* ManHoursSpent                                                          *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the number of man-hours spent on the item.
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
'* Rate                                                                   *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the hourly rate charged for work for customer on this service.
''</Description>
Public Property Get Rate() As Double
	Rate = mRate
End Property
Public Property Let Rate(Value As Double)
	mRate = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ReceivedAmount                                                         *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the amount of payment received.
''</Description>
Public Property Get ReceivedAmount() As Double
	ReceivedAmount = mReceivedAmount
End Property
Public Property Let ReceivedAmount(Value As Double)
	mReceivedAmount = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* RowIndex                                                               *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set the index of the row where this item was found.
''</Description>
Public Property Get RowIndex() As Integer
	RowIndex = mRowIndex
End Property
Public Property Let RowIndex(Value As Integer)
	mRowIndex = Value
End Property
'*------------------------------------------------------------------------*
 
'*------------------------------------------------------------------------*
'* Sent                                                                   *
'*------------------------------------------------------------------------*
''<Description>
''Get/Set a value indicating whether the invoice has been sent for this
'' item.
''</Description>
Public Property Get Sent() As Boolean
	Sent = mSent
End Property
Public Property Let Sent(Value As Boolean)
	mSent = Value
End Property
'*------------------------------------------------------------------------*

'*------------------------------------------------------------------------*
'* ServiceName                                                            *
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

'*------------------------------------------------------------------------*
'* SetContactCodeService                                                  *
'*------------------------------------------------------------------------*
''<Description>
''Set the contact code and/or service name for this entry from a combined
'' value.
''</Description>
Public Sub SetContactCodeService(Value As String)
Dim cs() As String  'Split value.
Dim ws As String    'Working string.

	If InStr(Value, "-") > 0 Then
		cs = Split(Value, "-")
	ElseIf InStr(Value, "_") > 0 Then
		cs = Split(Value, "_")
	Else
		ReDim cs(0)
		cs(0) = ""
	End If
	If UBound(cs) = 0 Then
		'Only one value is presented.
		ws = Trim(cs(0))
		If Len(ws) = 4 And ws = UCase(ws) Then
			'Contact code.
			mContactCode = ws
		ElseIf Len(ws) > 0 Then
			'Service name.
			mServiceName = ws
		End If
	Else
		'ContactCode, ServiceName.
		mContactCode = Trim(cs(0))
		mServiceName = Trim(cs(1))
	End If
	
End Sub
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
