VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aSettingsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : aSettingsTable
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:19    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



'4 columns like Settings_Registry: Application, Section, Key, Value

Option Explicit

Private Const SheetName = "Settings Table"
Private Const SettingsTableName = "tSettings"

Private SettingsTable As ListObject
Private SettingsSheet As Worksheet

Private Const colApp = 1
Private Const colSection = 2
Private Const colKey = 3
Private Const colValue = 4

Private Sub Class_Initialize()
    If Not WorksheetExists(SheetName, ThisWorkbook) Then CreateSheet
    Set SettingsSheet = ThisWorkbook.SHEETS(SheetName)
    Set SettingsTable = SettingsSheet.ListObjects(SettingsTableName)
End Sub

Private Sub CreateSheet()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.SHEETS(SheetName).Delete
    On Error GoTo 0
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS.Add(): ws.Name = SheetName
    Set SettingsSheet = ThisWorkbook.SHEETS(SheetName)
    Dim Headers As Variant: Headers = Array("Application", "Section", "Key", "Value")
    ArrayToRange1d Headers, True, SettingsSheet.Range("A1")
    Set SettingsTable = SettingsSheet.ListObjects.Add(xlSrcRange, SettingsSheet.Range("A1").CurrentRegion, , xlYes)
    SettingsTable.Name = SettingsTableName
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub AddOrModify(App, Section, key, Value)
    Dim var: var = Array(App, Section, key, Value)

    With SettingsTable
        If Not IsArray(Filter(App, Section, key)) Then
            .ListRows.Add
            .ListRows(.ListRows.Count).Range.Value = var
        Else
            Dim i As Long
            For i = 1 To .ListRows.Count
                If .ListRows(i).Range.Cells(colApp).Value = App _
                And .ListRows(i).Range.Cells(colSection).Value = Section _
                And .ListRows(i).Range.Cells(colKey).Value = key Then
                    .ListRows(i).Range.Cells(colValue).Value = Value
                End If
            Next
        End If
    End With
End Sub

Private Function Filter(App, Section, key)
    Dim var: var = Array(App, Section, key)
    Dim tVar
    tVar = ArrayFilter2d(SettingsTable.Range.Value, 1, App, True)
    tVar = ArrayFilter2d(tVar, 2, Section, False)
    tVar = ArrayFilter2d(tVar, 3, key, False)
    Filter = tVar
End Function

Public Function Apps() As Variant
    Dim arr
    arr = ArrayColumn(SettingsTable.DataBodyRange.Value, 1)
    Dim BA As New BetterArray
    BA.Items = arr
    BA.Unique
    Apps = BA.Items
    Apps = WorksheetFunction.Transpose(Apps)
End Function

Public Function Sections(App)
    Dim arr: arr = ArrayFilter2d(SettingsTable.DataBodyRange.Value, colApp, App, False)
    arr = ArrayColumn(arr, colSection)
    Dim BA As New BetterArray
    BA.Items = arr
    BA.Unique
    Sections = BA.Items
    
End Function

Public Function Keys(App, Section)
    Dim arr
    arr = ArrayFilter2d(SettingsTable.DataBodyRange.Value, colApp, App, False)
    arr = ArrayFilter2d(arr, colSection, Section, False)
    Keys = ArrayColumn(arr, colKey)
    
End Function

Public Function Value(App, Section, key)
    Value = Filter(App, Section, key)(1, colValue)
End Function

Function toTreeviewArray(App)
    Dim arr
    ReDim arr(1 To SettingsTable.ListColumns.Count, 1 To 1)
    Dim i As Long
    Dim Section
    Dim key

    arr(1, UBound(arr, 2)) = App
    For Each Section In Sections(App)
        ReDim Preserve arr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + 1)
        arr(2, UBound(arr, 2)) = Section
        
        For Each key In Keys(App, Section)
            ReDim Preserve arr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + 1)
            arr(3, UBound(arr, 2)) = key
            ReDim Preserve arr(1 To UBound(arr, 1), 1 To UBound(arr, 2) + 1)
            arr(4, UBound(arr, 2)) = Value(App, Section, key)
        Next
    Next
    
    toTreeviewArray = WorksheetFunction.Transpose(arr)
End Function

Function toINI(App)
    Dim Section
    Dim key
    
    Dim out As String
    
    For Each Section In Sections(App)
        out = out & IIf(out <> "", vbNewLine, "") & "[" & Section & "]"
        For Each key In Keys(App, Section)
            out = out & IIf(out <> "", vbNewLine, "") & Space(4) & key & "=" & Value(App, Section, key)
        Next
    Next
    
    toINI = out

End Function

Function toXML(App)
    Dim Section
    Dim key
    
    Dim var
    Dim out As String
    Dim indentation As Long
    
    out = out & IIf(out <> "", vbNewLine, "") & "<" & App & ">"
        For Each Section In Sections(App)
            out = out & IIf(out <> "", vbNewLine, "") & Space(4) & "<" & Section & ">"
            var = Keys(App, Section)
            For Each key In var
                out = out & IIf(out <> "", vbNewLine, "") & Space(8) & "<" & key & ">" & Value(App, Section, key) & "</" & key & ">"
                If key = var(UBound(var)) Then
                    out = out & IIf(out <> "", vbNewLine, "") & Space(4) & "</" & Section & ">"
                End If
            Next
        Next
    out = out & IIf(out <> "", vbNewLine, "") & "</" & App & ">"
    
    toXML = out
End Function
