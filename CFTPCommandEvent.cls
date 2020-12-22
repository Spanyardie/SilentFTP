VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFTPCommandEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarEventName As String 'local copy
Private mvarEventValue As String 'local copy
Public Property Let EventValue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EventValue = 5
    mvarEventValue = vData
End Property


Public Property Get EventValue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EventValue
    EventValue = mvarEventValue
End Property



Public Property Let EventName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EventName = 5
    mvarEventName = vData
End Property


Public Property Get EventName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EventName
    EventName = mvarEventName
End Property



