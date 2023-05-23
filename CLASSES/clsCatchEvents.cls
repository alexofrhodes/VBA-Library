VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCatchEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NOTE:
'       There are Attributes assigned in this class
'       They are visible if you export this component to file

'-----------------------
'NOTE : PUT IN MODULE
'-----------------------
'
'Option Explicit
'
''https://stackoverflow.com/questions/61855925/reducing-withevent-declarations-and-subs-with-vba-and-activex#answer-61893857
''Reducing WithEvent declarations and subs with VBA and ActiveX
'
'Private AllControls() As New clsCatchEvents
'
'Sub connect()
'    Dim j As Long
'    With Worksheets("Sheet1")
'    ReDim AllControls(.OLEObjects.Count - 1)
'        For j = 0 To .OLEObjects.Count - 1
'           AllControls(j).item = .OLEObjects(j + 1).Object
'           AllControls(j).Prop = .OLEObjects(j + 1).Name
'        Next
'    End With
'End Sub
'
'Sub disconnect()
'    Dim j As Long
'  For j = LBound(AllControls) To UBound(AllControls)
'          AllControls(j).clear
'   Next j
'    Erase AllControls
'End Sub
'-----------------------



Private Type GUID
      Data1 As Long
      Data2 As Integer
      Data3 As Integer
      Data4(0 To 7) As Byte
End Type

#If VBA7 And Win64 Then
      Private Declare PtrSafe Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, _
              ByRef riidEvent As GUID, ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, _
              Optional ByVal ppcpOut As LongPtr) As Long
#Else
     Private Declare Function ConnectToConnectionPoint Lib "shlwapi" Alias "#168" (ByVal punk As stdole.IUnknown, ByRef riidEvent As GUID, _
              ByVal fConnect As Long, ByVal punkTarget As stdole.IUnknown, ByRef pdwCookie As Long, Optional ByVal ppcpOut As Long) As Long
#End If

Private EventGuide As GUID
Private Ck As Long
Private ctl As Object
Private CustomProp As String

Public Sub ConnectAllEvents(ByVal connect As Boolean)
    With EventGuide
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    ConnectToConnectionPoint Me, EventGuide, connect, ctl, Ck, 0&
End Sub

Public Property Let Prop(newProp As String)
    CustomProp = newProp
End Property

Public Property Let item(Ctrl As Object)
    Set ctl = Ctrl
    Call ConnectAllEvents(True)
End Property

Public Sub clear()
    If (Ck <> 0) Then Call ConnectAllEvents(False)
    Set ctl = Nothing
End Sub

'Now when you run the connect sub every change of any activeX control is catched
'Edit: after comment to put in all other events; Other events: (all these will work on userforms also)

Public Sub MyChange()
    Debug.Print " Change ControlName " & " Type: " & TypeName(ctl) & " CustomProp: " & CustomProp
End Sub

Public Sub MyListClick()
    Debug.Print "cl1"
End Sub

Public Sub MyClick()
    Debug.Print "cl2"
End Sub

Public Sub MyDropButtonClick()
End Sub

Public Sub MyDblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print "dcl"
End Sub

Public Sub MyKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print "kd"
End Sub

Public Sub MyKeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print "ku"
End Sub

Public Sub MyMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Debug.Print "md"
End Sub

Public Sub MyMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Debug.Print "mm"
End Sub

Public Sub MyMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Debug.Print "mu"
End Sub

Public Sub myKeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Debug.Print "kp"
End Sub

'Then there are 4 (userform) events: Exit, Enter, AfterUpdate and BeforeUpdate
'which are events of the container-control which you can't 'catch' with withevents but in this way you can:

Public Sub myExit(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print "exit"
End Sub

Public Sub MyAfterUpdate()
    Debug.Print "au"
End Sub

Public Sub MyBeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print "bu"
End Sub

Public Sub MyEnter()
    Debug.Print "enter"
End Sub

'On a worksheet you've got LostFocus and GotFocus (1541 and 1542) but these I can't get to work,
'so if anybody knows howto it would be great. final remark: It doesnot work on a mac

