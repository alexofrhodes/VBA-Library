VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : aListView
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:18    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Private LV As ListView

#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#Else
    private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
#End If

Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

Function Init(TargetListView As ListView) As aListView
    Set LV = TargetListView
    Set Init = Me
End Function

Public Sub AutoSizeColumns(Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
            SendMessage LV.hWnd, LVM_FIRST + 30, c.index - 1, ByVal -2
        Next
    Else
        SendMessage LV.hWnd, LVM_FIRST + 30, Column.index - 1, ByVal -2
    End If
    LV.Refresh
End Sub

Public Sub ChangeValue(RowIndex As Long, ColumnIndex As Long, NewValue As Variant)
    Rem base 1 like range
    If ColumnIndex = 1 Then
        LV.ListItems(RowIndex).TEXT = NewValue
    ElseIf ColumnIndex > 1 Then
        LV.ListItems(RowIndex).ListSubItems(ColumnIndex - 1).TEXT = NewValue
    End If
End Sub

Public Function SELECTED(Optional ColumnIndex As Long = 0, Optional delimeter As String = ",") As Variant
    Rem base 1 like range
Retry:
    If ColumnIndex = 0 Then
        Dim s As String
        s = LV.ListItems(LV.SelectedItem.index)
        Dim counter
        For counter = 1 To LV.ColumnHeaders.Count - 1
            s = s & delimeter & LV.ListItems(LV.SelectedItem.index).ListSubItems(counter)
        Next
        SELECTED = Split(s, delimeter)
    ElseIf ColumnIndex = 1 Then
        SELECTED = LV.ListItems(LV.SelectedItem.index)
    ElseIf ColumnIndex > 1 Then
        SELECTED = LV.ListItems(ColumnIndex).ListSubItems(ColumnIndex - 1)
    Else
        ColumnIndex = -1
        GoTo Retry
    End If
End Function


Function LoadCSV(TargetFile As String)
    LoadArray CsvToArray(TargetFile)
End Function

Public Sub LoadArray(inputArray As Variant, Optional HeadersArray As Variant)
    clear
    Dim vListItem  As listItem
    Dim vChildItem As ListSubItem
    Dim vHeader    As Variant
    Dim iRows      As Long, iColumns As Long
    If IsMissing(HeadersArray) Then
        For iColumns = LBound(inputArray, 2) To UBound(inputArray, 2)
            Set vHeader = LV.ColumnHeaders.Add(, , inputArray(LBound(inputArray, 1), iColumns))
        Next
    Else
        Dim dif As Long
        Select Case LBound(inputArray, 2) - LBound(HeadersArray)
            Case 0:         dif = 0
            Case Is < 0:    dif = -1
            Case Is > 0:    dif = 1
        End Select
        For iColumns = LBound(inputArray, 2) To UBound(inputArray, 2)
            Set vHeader = LV.ColumnHeaders.Add(, , HeadersArray(iColumns - dif))
        Next
    End If
    For iRows = LBound(inputArray, 1) + IIf(IsMissing(HeadersArray), 1, 0) To UBound(inputArray, 1)
        Set vListItem = LV.ListItems.Add(, , inputArray(iRows, 1))
        For iColumns = LBound(inputArray, 2) + 1 To UBound(inputArray, 2)
            Set vChildItem = vListItem.ListSubItems.Add(, , inputArray(iRows, iColumns))
        Next
    Next
    LV.View = lvwReport
End Sub

Public Sub clear()
    Dim i As Long
    For i = LV.ListItems.Count To 1 Step -1
        LV.ListItems.Remove i
    Next
    LV.Refresh
End Sub




'Private Sub ListViewControl_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim dbBefore As Long: dbBefore = ListBox1.ListCount
'    Dim FileFullPath As String
'    Dim fileItem As Long
'    Dim objFSO As scripting.FileSystemObject
'    Dim objTopFolder As scripting.Folder
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'    Dim var As Variant, element As Variant
'    For fileItem = 1 To Data.Files.Count
'        FileFullPath = Data.Files(fileItem)
'        If oLogFiles = True Then
'            If LCase(IsFileFolderURL(FileFullPath)) = "f" Then
'                var = Split(TextBox2.TEXT, ",")
'                On Error Resume Next
'                If Left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" And (var(0) = "*" Or var(0) = "") Then GoTo PASS
'                For Each element In var
'                    If InStr(1, FileFullPath, element, vbTextCompare) > 0 And Left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" Then
'PASS:
'                        If Not aListBox.Init(ListBox1).Contains(FileFullPath) Then
'                            AddToListBox ListBox1, FileFullPath
'                        End If
'                    End If
'                Next
'            Else        'if drag dropped folder
'                Set objTopFolder = objFSO.getFolder(FileFullPath)
'                FileRecursive objTopFolder, oSearchInSubfolders.Value
'            End If
'        End If
'        If oLogFolders = True Then
'            If UCase(IsFileFolderURL(FileFullPath)) = "D" Then
'                Set objTopFolder = objFSO.getFolder(FileFullPath)
'                If Not aListBox.Init(ListBox1).Contains(objTopFolder.Path & "\") Then
'                    AddToListBox ListBox1, objTopFolder.Path
'                End If
'                FolderRecursive objTopFolder, oSearchInSubfolders.Value
'            End If
'        End If
'    Next fileItem
'
'    If ListBox1.ListCount - dbBefore > 0 Then ListboxToDatabaseSheet
'
'    Set objFSO = Nothing
'    Set objTopFolder = Nothing
'End Sub
