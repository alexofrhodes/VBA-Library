VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zTest 
   Caption         =   "UserForm1"
   ClientHeight    =   4428
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5916
   OleObjectBlob   =   "zTest.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonBar1_OnClick(ByVal ButtonId As Long)
    Stop
End Sub

Private Sub Slider1_Change()
    ProgressBar1.Value = Slider1.Value
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Debug.Print Panel.TEXT
    Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Debug.Print Button.Caption
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Debug.Print ButtonMenu.TEXT
End Sub

Private Sub UserForm_Click()
    goup
    comedown
End Sub

Private Sub comedown()
    Transition Effect(Me, "Top", 200, 2000)
End Sub

Sub goup()
    Transition Effect(Me, "Top", -Me.Height, 2000)
End Sub

Private Sub UserForm_Initialize()

    With ButtonBar1
    
    End With

    With ProgressBar1
        .Value = 50
    End With

    'TOOLBAR

    Toolbar1.Buttons.Add
    Toolbar1.Buttons(1).Caption = "but1"
    Toolbar1.Buttons(1).Style = tbrDropdown
    Toolbar1.Buttons(1).ButtonMenus.Add
    Toolbar1.Buttons(1).ButtonMenus(1).TEXT = "men1"
    'Toolbar1.Buttons(1).ButtonMenus(1).b
 
    Toolbar1.Buttons(1).ButtonMenus.item(1).TEXT = "menit1"

    Toolbar1.Buttons.Add
    Toolbar1.Buttons(2).Style = tbrButtonGroup


    'STATUSBAR

    StatusBar1.SimpleText = "asd;j"
    StatusBar1.Panels.Add , , "One"
    StatusBar1.Panels.Add

    With StatusBar1.Panels(1)
        .Bevel = sbrRaised
        .Alignment = sbrRight
        .TEXT = "as; ka slfk flkae"
        .TooltipText = "faosdh"
    End With
    With StatusBar1.Panels(2)
        .Style = sbrTime
    End With
    With StatusBar1.Panels(3)
        .TEXT = "ttt"
        .MinWidth = 5
        .AutoSize = sbrContents
    End With
End Sub
