
#If VBA7 Then
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#Else
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#End If

'____MODIFY THESE TO MATCH YOUR LOCAL AND GITHUB DIRECTORIES________

Public Const GITHUB_LIBRARY = "https://raw.githubusercontent.com/USERNAME/RepoName/"	'<---
    Public Const GITHUB_LIBRARY_DECLARATIONS = GITHUB_LIBRARY & "Declarations/"
    Public Const GITHUB_LIBRARY_PROCEDURES = GITHUB_LIBRARY & "Procedures/"
    Public Const GITHUB_LIBRARY_USERFORMS = GITHUB_LIBRARY & "Userforms/"
    Public Const GITHUB_LIBRARY_CLASSES = GITHUB_LIBRARY & "Classes/"

Public Const GITHUB_LOCAL_LIBRARY = "C:\Users\USERNAME\Documents\GitHub\VBA-Library\"   '<---
    Public Const GITHUB_LOCAL_LIBRARY_DECLARATIONS = GITHUB_LOCAL_LIBRARY & "Declarations\"
    Public Const GITHUB_LOCAL_LIBRARY_PROCEDURES = GITHUB_LOCAL_LIBRARY & "Procedures\"
    Public Const GITHUB_LOCAL_LIBRARY_USERFORMS = GITHUB_LOCAL_LIBRARY & "Userforms\"
    Public Const GITHUB_LOCAL_LIBRARY_CLASSES = GITHUB_LOCAL_LIBRARY & "Classes\"

'___________________________________________________________________

Sub AddLinkedListsToActiveProcedure()
    AddLinkedLists ThisWorkbook, ActiveModule, ActiveProcedure
End Sub

Sub ExportActiveProcedure()
    ExportProcedure ThisWorkbook, ActiveModule, ActiveProcedure, ExportMergedTxt:=True
End Sub

Sub ExportAllProceduresOfThisWorkbook()
    ExportAllProcedures ThisWorkbook
End Sub

Sub ImportActiveProcedureDependencies()
    ImportProcedureDependencies ThisWorkbook, ActiveModule, ActiveProcedure, Overwrite:=True
End Sub

Sub ExportAllProcedures(TargetWorkbook As Workbook)
    Dim procedure
    Dim module As VBComponent
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.Type = vbext_ct_StdModule Then
            For Each procedure In ProceduresOfModule
                ExportProcedure TargetWorkbook, module, procedure, False
            Next procedure
        End If
    Next module
End Sub

Sub RemoveComments(TargetWorkbook As Workbook)
    Dim module As VBComponent
    Dim s As String
    Dim i As Long
    For Each module In TargetWorkbook.VBProject.VBComponents
        For i = module.CodeModule.CountOfLines To 1 Step -1
            s = Trim(module.CodeModule.Lines(i, 1))
            If s Like "'*" Or s Like "Rem *" Then module.CodeModule.DeleteLines i, 1
        Next i
    Next
End Sub
Function ArrayAppend(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    Dim holdarr As Variant
    Dim ub1 As Long
    Dim ub2 As Long
    Dim i As Long
    Dim newind As Long
    If IsEmpty(arr1) Or Not IsArray(arr1) Then
        arr1 = Array()
    End If
    If IsEmpty(arr2) Or Not IsArray(arr2) Then
        arr2 = Array()
    End If
    ub1 = UBound(arr1)
    ub2 = UBound(arr2)
    If ub1 = -1 Then
        ArrayAppend = arr2
        Exit Function
    End If
    If ub2 = -1 Then
        ArrayAppend = arr1
        Exit Function
    End If
    holdarr = arr1
    ReDim Preserve holdarr(ub1 + ub2 + 1)
    newind = UBound(arr1) + 1
    For i = 0 To ub2
        If VarType(arr2(i)) = vbObject Then
            Set holdarr(newind) = arr2(i)
        Else
            holdarr(newind) = arr2(i)
        End If
        newind = newind + 1
    Next i
    ArrayAppend = holdarr
End Function

Public Sub ArrayQuickSort(ByRef SortableArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant
    If IsEmpty(SortableArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortableArray), "()") < 1 Then
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortableArray)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortableArray)
    End If
    If lngMin >= lngMax Then
        Exit Sub
    End If
    i = lngMin
    j = lngMax
    varMid = Empty
    varMid = SortableArray((lngMin + lngMax) \ 2)
    If IsObject(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If
    While i <= j
        While SortableArray(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortableArray(j) And j > lngMin
            j = j - 1
        Wend
        If i <= j Then
            varX = SortableArray(i)
            SortableArray(i) = SortableArray(j)
            SortableArray(j) = varX
            i = i + 1
            j = j - 1
        End If
    Wend
    If (lngMin < j) Then Call ArrayQuickSort(SortableArray, lngMin, j)
    If (i < lngMax) Then Call ArrayQuickSort(SortableArray, i, lngMax)
End Sub

Public Function cleanArray(varArray As Variant) As Variant()
  Dim TempArray() As Variant
  Dim OldIndex As Integer
  Dim NewIndex As Integer
  Dim Output As String
  If Not ArrayAllocated(varArray) Then Exit Function
  ReDim TempArray(LBound(varArray) To UBound(varArray))
  For OldIndex = LBound(varArray) To UBound(varArray)
    Output = CleanTrim(varArray(OldIndex))
    If Not Output = "" Then
        TempArray(NewIndex) = Output
        NewIndex = NewIndex + 1
    End If
 Next OldIndex
 ReDim Preserve TempArray(LBound(varArray) To NewIndex - 1)
 cleanArray = TempArray
End Function

Function ArrayDuplicatesRemove(myArray As Variant) As Variant
    Dim nFirst As Long, nLast As Long, i As Long
    Dim Item As String

    Dim arrTemp() As String
    Dim coll As New Collection
    If Not ArrayAllocated(myArray) Then Exit Function
    nFirst = LBound(myArray)
    nLast = UBound(myArray)
    ReDim arrTemp(nFirst To nLast)

    For i = nFirst To nLast
        arrTemp(i) = CStr(myArray(i))
    Next i

    On Error Resume Next
    For i = nFirst To nLast
        coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0

    nLast = coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)

    For i = nFirst To nLast
        arrTemp(i) = coll(i - nFirst + 1)
    Next i

    ArrayDuplicatesRemove = arrTemp

End Function

Public Function ArrayToCollection(Items As Variant) As Collection
    If Not ArrayAllocated(Items) Then Exit Function
    Dim coll As New Collection
    Dim i As Integer
    For i = LBound(Items) To UBound(Items)
        coll.Add Items(i)
    Next
    Set ArrayToCollection = coll
End Function

Function CleanTrim(ByVal s As String, Optional ConvertNonBreakingSpace As Boolean = True) As String
    Dim X As Long, CodesToClean As Variant
    CodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                         21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
    If ConvertNonBreakingSpace Then s = Replace(s, Chr(160), " ")
    s = Replace(s, vbCr, "")
    For X = LBound(CodesToClean) To UBound(CodesToClean)
        If InStr(s, Chr(CodesToClean(X))) Then
            s = Replace(s, Chr(CodesToClean(X)), vbNullString)
        End If
    Next
    CleanTrim = s
    CleanTrim = Trim(s)
End Function

Sub AddLinkedLists(Optional TargetWorkbook As Workbook, _
                    Optional module As VBComponent, _
                    Optional procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Exit Sub
    ProcedureLinesRemoveInclude TargetWorkbook, module, procedure
    ProcedureAssignedModuleAdd TargetWorkbook, module, procedure
    AddListOfLinkedProceduresToProcedure TargetWorkbook, module, procedure
    AddListOfLinkedClassesToProcedure TargetWorkbook, module, procedure
    AddListOfLinkedUserformsToProcedure TargetWorkbook, module, procedure
    AddListOfLinkedDeclarationsToProcedure TargetWorkbook, module, procedure
    
End Sub


Sub AddListOfLinkedClassesToProcedure( _
                                     Optional TargetWorkbook As Workbook, _
                                     Optional module As VBComponent, _
                                     Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop
    Dim ListOfImports As String
    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim myClasses As Collection
    Set myClasses = LinkedClasses(TargetWorkbook, module, ProcedureName)
    Dim Element As Variant
    For Each Element In myClasses
        If InStr(1, Code, "@INCLUDE CLASS " & Element) = 0 _
        And InStr(1, ListOfImports, "@INCLUDE CLASS " & Element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE CLASS " & Element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE CLASS " & Element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines _
            ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedDeclarationsToProcedure( _
                                          Optional TargetWorkbook As Workbook, _
                                          Optional module As VBComponent, _
                                          Optional ProcedureName As String)

    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim ListOfImports As String
    If module Is Nothing Then Set module = ModuleOfProcedure(TargetWorkbook, ProcedureName)
    Dim ProcedureText As String
    ProcedureText = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim myDeclarations As Collection
    Set myDeclarations = LinkedDeclarations(TargetWorkbook, module, ProcedureName)
    Dim coll As New Collection
    Dim Element As Variant
    For Each Element In myDeclarations
        If InStr(1, ProcedureText, "'@INCLUDE DECLARATION " & Element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE DECLARATION " & Element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE DECLARATION " & Element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedProceduresToProcedure( _
                                        Optional TargetWorkbook As Workbook, _
                                        Optional module As VBComponent, _
                                        Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop
    Dim Procedures As Collection
    Set Procedures = LinkedProcedures(TargetWorkbook, module, ProcedureName)
    Dim ListOfImports As String
    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim procedure As Variant
    For Each procedure In Procedures
        If InStr(1, Code, "@INCLUDE PROCEDURE " & procedure) = 0 And InStr(1, ListOfImports, "@INCLUDE PROCEDURE " & procedure) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE PROCEDURE " & procedure
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE PROCEDURE " & procedure
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedUserformsToProcedure( _
                                       Optional TargetWorkbook As Workbook, _
                                       Optional module As VBComponent, _
                                       Optional ProcedureName As String)
    
    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop

    Dim ListOfImports As String
    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim myClasses As Collection
    Set myClasses = LinkedUserforms(TargetWorkbook, module, ProcedureName)
    Dim Element As Variant
    For Each Element In myClasses
        If InStr(1, Code, "@INCLUDE USERFORM " & Element) = 0 And InStr(1, ListOfImports, "@INCLUDE USERFORM " & Element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE USERFORM " & Element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE USERFORM " & Element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Public Function ActiveProcedure() As String
    Application.VBE.ActiveCodePane.GetSelection L1&, c1&, L2&, c2&
    ActiveProcedure = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(L1&, vbext_pk_Proc)
End Function

Public Function ActiveModule() As VBComponent
    Set ActiveModule = Application.VBE.SelectedVBComponent
End Function

Public Function ActiveCodepaneWorkbook() As Workbook
    On Error GoTo ErrorHandler
    Dim WorkbookName As String
    WorkbookName = Application.VBE.SelectedVBComponent.Collection.Parent.FileName
    WorkbookName = Right(WorkbookName, Len(WorkbookName) - InStrRev(WorkbookName, "\"))
    Set ActiveCodepaneWorkbook = Workbooks(WorkbookName)
    Exit Function
ErrorHandler:
    MsgBox "doesn't work on new-unsaved workbooks"
End Function

Public Function ArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    ArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
End Function

Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    Dim i As Integer
    Dim test As Long
    On Error GoTo Catch
    Do
        i = i + 1
        test = UBound(SourceArray, i)
    Loop
Catch:
    ArrayDimensionLength = i - 1
End Function

Public Sub ArrayToRange2D(arr2d As Variant, Cell As Range)

    If ArrayDimensionLength(arr2d) = 1 Then arr2d = WorksheetFunction.Transpose(arr2d)
    Dim dif As Long
        dif = IIf(LBound(arr2d, 1) = 0, 1, 0)
    Dim rng As Range
    Set rng = Cell.Resize(UBound(arr2d, 1) + dif, UBound(arr2d, 2) + dif)

    If Application.WorksheetFunction.CountA(rng) > 0 Then
        Exit Sub
    End If

    rng.Value = arr2d
End Sub

Function AssignCPSvariables( _
                            ByRef TargetWorkbook As Workbook, _
                            ByRef module As VBComponent, _
                            ByRef procedure As String) As Boolean

    If Not AssignWorkbookVariable(TargetWorkbook) Then Exit Function
    If Not AssignModuleVariable(TargetWorkbook, module) Then Exit Function
    If Not AssignProcedureVariable(TargetWorkbook, procedure) Then Exit Function
    AssignCPSvariables = True
    
End Function

Function AssignModuleVariable( _
                             ByVal TargetWorkbook As Workbook, _
                             ByRef module As VBComponent, _
                             Optional ByVal procedure As String) As Boolean
    If procedure = "" Then
        On Error Resume Next
        Set module = ActiveModule
        On Error GoTo 0
    ElseIf module Is Nothing Then
        On Error Resume Next
        Set module = ModuleOfProcedure(TargetWorkbook, procedure)
        On Error GoTo 0
    End If
    AssignModuleVariable = Not module Is Nothing
End Function

Function AssignProcedureVariable(TargetWorkbook As Workbook, ByRef procedure As String) As Boolean
    If procedure = "" Then
        Dim cps As String
        cps = CodepaneSelection
        If Len(cps) > 0 Then
            procedure = cps
        Else
            procedure = ActiveProcedure
        End If
        If Not ProcedureExists(TargetWorkbook, procedure) Then
            Debug.Print procedure & " not found in Workbook " & TargetWorkbook.Name
        End If
    End If
    AssignProcedureVariable = Not procedure = ""
End Function

Function AssignWorkbookVariable(ByRef TargetWorkbook As Workbook) As Boolean
     If TargetWorkbook Is Nothing Then
        On Error Resume Next
        Set TargetWorkbook = ActiveCodepaneWorkbook
        On Error GoTo 0
    End If
    AssignWorkbookVariable = Not TargetWorkbook Is Nothing
End Function

Function CheckPath(Path) As String
    Dim retval
    retval = "I"
    If (retval = "I") And FileExists(Path) Then retval = "F"
    If (retval = "I") And FolderExists(Path) Then retval = "D"
    If (retval = "I") And URLExists(Path) Then retval = "U"
    CheckPath = retval
End Function

Function ClassNames(Optional TargetWorkbook As Workbook)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Set ClassNames = ComponentNames(vbext_ct_ClassModule, TargetWorkbook)
End Function

Public Function CodepaneSelection() As String
    Dim startLine As Long, StartColumn As Long, endLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, endLine, EndColumn
    If endLine - startLine = 0 Then
        CodepaneSelection = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(startLine, 1), StartColumn, EndColumn - StartColumn)
        Exit Function
    End If
    Dim Str As String
    Dim i As Long
    For i = startLine To endLine
        If Str = "" Then
            Str = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1), StartColumn)
        ElseIf i < endLine Then
            Str = Str & vbNewLine & Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
        Else
            Str = Str & vbNewLine & Left(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1), EndColumn - 1)
        End If
    Next
    CodepaneSelection = Str
End Function

Public Function CollectionContains( _
                                  Kollection As Collection, _
                                  Optional key As Variant, _
                                  Optional Item As Variant) As Boolean
    Dim strKey As String
    Dim var As Variant
    If Not IsMissing(key) Then
        strKey = CStr(key)
        On Error Resume Next
        CollectionContains = True
        var = Kollection(strKey)
        If Err.Number = 91 Then GoTo CheckForObject
        If Err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function
CheckForObject:
        If IsObject(Kollection(strKey)) Then
            CollectionContains = True
            On Error GoTo 0
            Exit Function
        End If
NotFound:
        CollectionContains = False
        On Error GoTo 0
        Exit Function
    ElseIf Not IsMissing(Item) Then
        CollectionContains = False
        For Each var In Kollection
            If var = Item Then
                CollectionContains = True
                Exit Function
            End If
        Next var
    Else
        CollectionContains = False
    End If
End Function

Public Function CollectionSort(colInput As Collection) As Collection
    Dim iCounter As Integer
    Dim iCounter2 As Integer
    Dim Temp As Variant
    Set CollectionSort = New Collection
    For iCounter = 1 To colInput.Count - 1
        For iCounter2 = iCounter + 1 To colInput.Count
            If colInput(iCounter) > colInput(iCounter2) Then
                Temp = colInput(iCounter2)
                colInput.Remove iCounter2
                colInput.Add Temp, , iCounter
            End If
        Next iCounter2
    Next iCounter
    Set CollectionSort = colInput
End Function

Function CollectionsToArray2D(collections As Collection) As Variant
    If collections.Count = 0 Then Exit Function
    Dim columnCount As Long
    columnCount = collections.Count
    Dim rowCount As Long
    rowCount = collections.Item(1).Count
    Dim var As Variant
    ReDim var(1 To rowCount, 1 To columnCount)
    Dim cols As Long
    Dim rows As Long
    For rows = 1 To rowCount
        For cols = 1 To collections.Count
            var(rows, cols) = collections(cols).Item(rows)
        Next cols
    Next rows
    CollectionsToArray2D = var
End Function




Function ComponentNames( _
                       ModuleType As vbext_ComponentType, _
                       Optional TargetWorkbook As Workbook)
    Dim coll As New Collection
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim module As VBComponent
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.Type = ModuleType Then
            coll.Add module.Name
        End If
    Next
    Set ComponentNames = coll
End Function

Function DeclarationsKeywordSubstring(Str As Variant, Optional delim As String _
                , Optional afterWord As String _
                , Optional beforeWord As String _
                , Optional counter As Integer _
                , Optional outer As Boolean _
                , Optional includeWords As Boolean) As String
    Dim i As Long
    If afterWord = "" And beforeWord = "" And counter = 0 Then
        MsgBox ("Pass at least 1 parameter betweenn -AfterWord- , -BeforeWord- , -counter-")
        Exit Function
    End If
    If TypeName(Str) = "String" Then
        If delim <> "" Then
            Str = Split(Str, delim)
            If UBound(Str) <> 0 Then
                If afterWord = "" And beforeWord = "" And counter <> 0 Then
                    If counter - 1 <= UBound(Str) Then
                        DeclarationsKeywordSubstring = Str(counter - 1)
                        Exit Function
                    End If
                End If
                For i = LBound(Str) To UBound(Str)
                    If afterWord <> "" And beforeWord = "" Then
                        If i <> 0 Then
                            If Str(i - 1) = afterWord Or Str(i - 1) = "#" & afterWord Then
                                DeclarationsKeywordSubstring = Str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord = "" And beforeWord <> "" Then
                        If i <> UBound(Str) Then
                            If Str(i + 1) = beforeWord Or Str(i + 1) = "#" & beforeWord Then
                                DeclarationsKeywordSubstring = Str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord <> "" And beforeWord <> "" Then
                        If i <> 0 And i <> UBound(Str) Then
                            If (Str(i - 1) = afterWord Or Str(i - 1) = "#" & afterWord) And (Str(i + 1) = beforeWord Or Str(i + 1) = "#" & beforeWord) Then
                                DeclarationsKeywordSubstring = Str(i)
                                Exit Function
                            End If
                        End If
                    End If
                Next i
            End If
        Else
            If InStr(1, Str, afterWord) > 0 And InStr(1, Str, beforeWord) > 0 Then
                If includeWords = False Then
                    DeclarationsKeywordSubstring = Mid(Str, InStr(1, Str, afterWord) + Len(afterWord))
                Else
                    DeclarationsKeywordSubstring = Mid(Str, InStr(1, Str, afterWord))
                End If
                If outer = True Then
                    If includeWords = False Then
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStrRev(DeclarationsKeywordSubstring, beforeWord) - 1)
                    Else
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStrRev(DeclarationsKeywordSubstring, beforeWord) + Len(beforeWord) - 1)
                    End If
                Else
                    If includeWords = False Then
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStr(1, DeclarationsKeywordSubstring, beforeWord) - 1)
                    Else
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStr(1, DeclarationsKeywordSubstring, beforeWord) + Len(beforeWord) - 1)
                    End If
                End If
                Exit Function
            End If
        End If
    Else
    End If
    DeclarationsKeywordSubstring = vbNullString
End Function

Sub DeclarationsTableCreate(TargetWorkbook As Workbook)

    DeclarationsWorksheetCreate
    
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    If Format(Now, "YYMMDDHHNN") - TargetWorksheet.Range("Z1").Value < 60 Then Exit Sub
    
    TargetWorksheet.Range("A2").CurrentRegion.Offset(1).Clear
    ArrayToRange2D CollectionsToArray2D( _
                        getDeclarations( _
                            wb:=TargetWorkbook, _
                            includeScope:=True, _
                            includeType:=True, _
                            includeKeywords:=True, _
                            includeDeclarations:=True, _
                            includeComponentName:=True, _
                            includeComponentType:=True)), _
                    TargetWorksheet.Range("A2")
                                        
    TargetWorksheet.Range("Z1").Value = Format(Now, "YYMMDDHHNN")
    
    DeclarationsTableSort
End Sub


Function DeclarationsTableKeywords() As Collection
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim lr As Long: lr = getLastRow(TargetWorksheet)
    Dim coll As New Collection
    Dim Cell As Range
    For Each Cell In TargetWorksheet.Range("C2:C" & lr)
        On Error Resume Next
        coll.Add Cell.Text, Cell.Text
        On Error GoTo 0
    Next
    Set DeclarationsTableKeywords = coll
End Function


Sub DeclarationsTableSort()

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Worksheets("Declarations_Table")
    
    Dim sort1 As String: sort1 = "B1"
    Dim sort2 As String: sort2 = "C1"
    Dim sort3 As String ': sort3 = "D1"

    With TargetWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add key:=TargetWorksheet.Range(sort1), Order:=xlAscending
        
        If Not sort2 = "" Then
            .SortFields.Add key:=TargetWorksheet.Range(sort2), Order:=xlAscending
        End If
        If Not sort3 = "" Then
            .SortFields.Add key:=TargetWorksheet.Range(sort3), Order:=xlAscending
        End If

        .SetRange TargetWorksheet.Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With
    
End Sub



Function DeclarationsWorksheetCreate() As Boolean
    If WorksheetExists("Declarations_Table", ThisWorkbook) Then Exit Function
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets.Add
    With TargetWorksheet
        .Name = "Declarations_Table"
        .Cells.VerticalAlignment = xlVAlignTop
        .Range("A1:F1").Value = Split("SCOPE,TYPE,NAME,CODE,MODULE TYPE,MODULE NAME", ",")
        .rows(1).Cells.Font.Bold = True
        .rows(1).Cells.Font.Size = 14
    End With
End Function

Sub ExportLinkedDeclaration(TargetWorkbook As Workbook, DeclarationName As String)
    DeclarationsTableCreate TargetWorkbook
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")

    Dim codeName As String
    Dim codeText As String
    Dim Cell As Range
    On Error Resume Next
    Set Cell = TargetWorksheet.Columns(3).Find(DeclarationName, LookAt:=xlWhole)
    On Error GoTo 0
    If Cell Is Nothing Then Exit Sub

    codeName = DeclarationName
    codeText = Cell.Offset(0, 1).Text
    TxtOverwrite GITHUB_LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt", codeText

End Sub



Function ExportProcedure( _
                    Optional TargetWorkbook As Workbook, _
                    Optional module As VBComponent, _
                    Optional ProcedureName As String, _
                    Optional ExportMergedTxt As Boolean) As String

    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Exit Function

    ProjetFoldersCreate

    Dim ExportedProcedures As New Collection
    On Error GoTo ErrorHandler

    ExportedProcedures.Add CStr(ProcedureName), CStr(ProcedureName)

    Dim procedure
    For Each procedure In LinkedProceduresDeep(ProcedureName, TargetWorkbook)
        ExportedProcedures.Add CStr(procedure), CStr(procedure)
    Next

    If ExportedProcedures.Count > 1 Then
    
        Dim MergedName As String
            MergedName = "Merged_" & ProcedureName
        Dim FileName As String
            FileName = GITHUB_LOCAL_LIBRARY_PROCEDURES & MergedName & ".txt"
        Dim MergedString As String

        For Each procedure In ExportedProcedures
            MergedString = MergedString & vbNewLine & ProcedureCode(TargetWorkbook, , procedure)
        Next
        ExportProcedure = MergedString
        
        If ExportMergedTxt Then
            Debug.Print "OVERWROTE " & MergedName
            TxtOverwrite FileName, MergedString
            TxtPrependContainedProcedures FileName
        End If

        For Each procedure In ExportedProcedures
            ExportTargetProcedure TargetWorkbook, , CStr(procedure)
        Next
    End If

    FollowLink GITHUB_LOCAL_LIBRARY_PROCEDURES
    
    Exit Function
ErrorHandler:
    MsgBox "An error occured in Sub ExportProcedure"
End Function

Sub ExportTargetProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional procedure As String)

    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Exit Sub

    Dim proclastmod
        proclastmod = ProcedureLastModified(TargetWorkbook, module, procedure)
    If proclastmod = 0 Then
        AddLinkedLists TargetWorkbook, module, procedure
        proclastmod = ProcedureLastModAdd(TargetWorkbook, module, procedure)
    End If

    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, CStr(procedure))
    Dim FileFullName As String
        FileFullName = GITHUB_LOCAL_LIBRARY_PROCEDURES & procedure & ".txt"
    If FileExists(FileFullName) Then
        Dim filelastmod
            filelastmod = StringLastModified(TxtRead(FileFullName))
        If proclastmod > filelastmod Then
            Debug.Print "OVERWROTE " & procedure
            TxtOverwrite FileFullName, Code
        End If
    Else
        Debug.Print "NEW " & procedure
        TxtOverwrite FileFullName, Code
    End If

    Dim Element
    For Each Element In LinkedUserforms(TargetWorkbook, module, CStr(procedure))
        TargetWorkbook.VBProject.VBComponents(Element).Export GITHUB_LOCAL_LIBRARY_USERFORMS & Element & ".frm"
    Next
    For Each Element In LinkedClasses(TargetWorkbook, module, CStr(procedure))
        TargetWorkbook.VBProject.VBComponents(Element).Export GITHUB_LOCAL_LIBRARY_CLASSES & Element & ".cls"
    Next
    For Each Element In LinkedDeclarations(TargetWorkbook, module, CStr(procedure))
        ExportLinkedDeclaration TargetWorkbook, CStr(Element)
    Next
End Sub

Public Function FileExists(ByVal FileName As String) As Boolean
    If InStr(1, FileName, "\") = 0 Then Exit Function
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileExists = (Dir(FileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Sub FoldersCreate(FolderPath As String)
    On Error Resume Next
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim ArrayElement As Variant
    individualFolders = Split(FolderPath, "\")
    For Each ArrayElement In individualFolders
        tempFolderPath = tempFolderPath & ArrayElement & "\"
        If FolderExists(tempFolderPath) = False Then
            MkDir tempFolderPath
        End If
    Next ArrayElement
End Sub

Sub FollowLink(FolderPath As String)
    If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)
    On Error Resume Next
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.document.Folder.Self.Path = FolderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub



Function FormatVBA7(Str As String) As String
    Dim selectedText
        selectedText = Str
        selectedText = Replace(selectedText, " _" & vbNewLine, "")
        selectedText = Split(selectedText, vbNewLine)
    Dim IsVba7 As String
    Dim NotVba7 As String
    Dim colIsVBA7 As New Collection
    Dim colNotVBA7 As New Collection
    Dim i As Long
    For i = LBound(selectedText) To UBound(selectedText)
        If InStr(1, selectedText(i), "PtrSafe", vbTextCompare) Then
            IsVba7 = selectedText(i)
            NotVba7 = Replace(selectedText(i), "Declare ptrsafe ", "Declare ", , , vbTextCompare)
        Else
            IsVba7 = Replace(selectedText(i), "Declare ", "Declare PtrSafe ")
            NotVba7 = selectedText(i)
        End If
        colIsVBA7.Add IsVba7
        colNotVBA7.Add NotVba7
    Next
    Set colIsVBA7 = CollectionSort(colIsVBA7)
    Set colNotVBA7 = CollectionSort(colNotVBA7)
    Dim out As String
        out = "#If VBA7 then" & vbNewLine & _
        collectionToString(colIsVBA7, vbNewLine) & vbNewLine & _
        "#Else" & vbNewLine & _
        collectionToString(colNotVBA7, vbNewLine) & vbNewLine & _
        "#End If"
    FormatVBA7 = out

End Function

Function GetMotherBoardProp() As String

    Dim strComputer As String
    Dim objSvcs As Object
    Dim objItms As Object, objItm As Object
    Dim vItem
    strComputer = "."
    Set objSvcs = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set objItms = objSvcs.execquery("Select * from Win32_BaseBoard")
    For Each objItm In objItms
        GetMotherBoardProp = objItm.SerialNumber
    Next

    Set objSvcs = Nothing
End Function

Public Function GetSheetByCodeName(wb As Workbook, codeName As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If UCase(sh.codeName) = UCase(codeName) Then Set GetSheetByCodeName = sh: Exit For
    Next sh
End Function

Sub ImportClass( _
                    Optional ClassName As String, _
                    Optional TargetWorkbook As Workbook, _
                    Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If ClassName = "" Then ClassName = CodepaneSelection
    If ClassName = "" Or InStr(1, ClassName, " ") > 0 Then Exit Sub
    Dim filePath As String
    filePath = GITHUB_LOCAL_LIBRARY_CLASSES & ClassName & ".cls"
    If CheckPath(filePath) = "I" Then
        On Error Resume Next
        Dim Code As String
        Code = TXTReadFromUrl(GITHUB_LIBRARY_CLASSES & ClassName & ".cls")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            TxtOverwrite filePath, Code
        Else
            MsgBox "File " & ClassName & ".cls not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(ClassName, TargetWorkbook) Then
        If Overwrite = True Then
            TargetWorkbook.VBProject.VBComponents.Remove TargetWorkbook.VBProject.VBComponents(ClassName)
        Else
            Exit Sub
        End If
    End If
    TargetWorkbook.VBProject.VBComponents.Import filePath
End Sub


Sub ImportDeclaration( _
                        Optional DeclarationName As String, _
                        Optional module As VBComponent, _
                        Optional TargetWorkbook As Workbook)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If DeclarationName = "" Then DeclarationName = CodepaneSelection
    If DeclarationName = "" Or InStr(1, DeclarationName, " ") > 0 Then Exit Sub
    Dim filePath As String
    filePath = GITHUB_LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt"
    Dim Code As String
    On Error Resume Next
    Code = TxtRead(filePath)
    On Error GoTo 0

    If Len(Code) = 0 Then 'CheckPath(filePath) = "I" Then
        On Error Resume Next
        Code = TXTReadFromUrl(GITHUB_LIBRARY_DECLARATIONS & DeclarationName & ".txt")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            Code = FormatVBA7(Code)
            TxtOverwrite filePath, Code
        Else
            Debug.Print "File " & DeclarationName & ".txt not found localy or online"
            Exit Sub
        End If
    Else

    End If
    If InStr(1, WorkbookCode(TargetWorkbook), Code, vbTextCompare) > 0 Then Exit Sub
    If module Is Nothing Then Set module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
    module.CodeModule.AddFromString Code

End Sub







Sub ImportProcedure( _
                    Optional procedure As String, _
                    Optional TargetWorkbook As Workbook, _
                    Optional module As VBComponent, _
                    Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If procedure = "" Then procedure = CodepaneSelection
    If procedure = "" Or InStr(1, procedure, " ") > 0 Then Exit Sub
    Dim ProcedurePath As String
        ProcedurePath = GITHUB_LOCAL_LIBRARY_PROCEDURES & procedure & ".txt"

    Dim Code As String
    On Error Resume Next
    Code = TxtRead(ProcedurePath)
    On Error GoTo 0

    If Len(Code) = 0 Then
        On Error Resume Next
        Code = TXTReadFromUrl(GITHUB_LIBRARY_PROCEDURES & procedure & ".txt")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            TxtOverwrite ProcedurePath, Code
        Else
            MsgBox "File " & procedure & ".txt not found neither localy nor online"
            Exit Sub
        End If
    End If

    Dim filelastmod
        filelastmod = StringLastModified(Code)
    Dim proclastmod

    If ProcedureExists(TargetWorkbook, procedure) = True Then
        Set module = ModuleOfProcedure(TargetWorkbook, procedure)
        proclastmod = ProcedureLastModified(TargetWorkbook, module, procedure)
        If Overwrite = True Then
            If proclastmod = 0 Or proclastmod < filelastmod Then
                ProcedureReplace module, procedure, TxtRead(ProcedurePath)
            End If
        End If
    Else
        If module Is Nothing Then
            Dim ModuleName As String
                ModuleName = StringProcedureAssignedModule(Code)
            If ModuleName = "" Then ModuleName = "vbArcImports"
            Set module = ModuleAddOrSet(TargetWorkbook, ModuleName, vbext_ct_StdModule)
        End If
        module.CodeModule.AddFromFile ProcedurePath
    End If

    ImportProcedureDependencies procedure, TargetWorkbook, module, Overwrite
    ProcedureMoveToAssignedModule TargetWorkbook, module, procedure
End Sub

Sub ImportProcedureDependencies( _
                 Optional procedure As String, _
                 Optional TargetWorkbook As Workbook, _
                 Optional module As VBComponent, _
                 Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If procedure = "" Then
        Dim cps As String
        cps = CodepaneSelection
        If Len(cps) > 0 Then
            procedure = cps
        Else
            procedure = ActiveProcedure
        End If
        If Not ProcedureExists(TargetWorkbook, procedure) Then Exit Sub
    End If
    On Error Resume Next
    If module Is Nothing Then Set module = ModuleOfProcedure(TargetWorkbook, procedure)
    If module Is Nothing Then Exit Sub
    On Error GoTo 0
    Dim var
    Dim importfile As String
    var = Split(ProcedureCode(TargetWorkbook, module, procedure), vbNewLine)
    var = Filter(var, "'@INCLUDE ")
    Dim TextLine As Variant
    For Each TextLine In var
        TextLine = Trim(TextLine)
        If TextLine Like "'@INCLUDE *" Then
            importfile = Split(TextLine, " ")(2)
            importfile = Replace(importfile, vbNewLine, "")
            If TextLine Like "'@INCLUDE PROCEDURE *" Then
                ImportProcedure importfile, TargetWorkbook, module, Overwrite
            ElseIf TextLine Like "'@INCLUDE CLASS *" Then
                ImportClass importfile, TargetWorkbook, Overwrite
            ElseIf TextLine Like "'@INCLUDE USERFORM *" Then
                ImportUserform importfile, TargetWorkbook, Overwrite
            ElseIf TextLine Like "'@INCLUDE DECLARATION *" Then
                ImportDeclaration importfile, module, TargetWorkbook
            End If
        End If
    Next
End Sub

Sub ImportUserform( _
                    Optional UserformName As String, _
                    Optional TargetWorkbook As Workbook, _
                    Optional Overwrite As Boolean)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If UserformName = "" Then UserformName = CodepaneSelection
    If UserformName = "" Or InStr(1, UserformName, " ") > 0 Then Exit Sub
    Dim FilePathFrM As String
        FilePathFrM = GITHUB_LOCAL_LIBRARY_USERFORMS & UserformName & ".frm"
    Dim FilePathFrX As String
        FilePathFrX = GITHUB_LOCAL_LIBRARY_USERFORMS & UserformName & ".frx"

    If CheckPath(FilePathFrM) = "I" Then
        On Error Resume Next
        Dim codeFrM As String
            codeFrM = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frm")
        Dim codeFrX As String
            codeFrX = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frx")
        On Error GoTo 0
        If Len(codeFrM) > 0 And Len(codeFrX) > 0 Then
            TxtOverwrite FilePathFrM, codeFrM
            TxtOverwrite FilePathFrX, codeFrX
        Else
            MsgBox "File " & UserformName & ".frm/.frx not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(UserformName, TargetWorkbook) Then
        If Overwrite = True Then
            TargetWorkbook.VBProject.VBComponents.Remove TargetWorkbook.VBProject.VBComponents(UserformName)
        Else
            Exit Sub
        End If
    End If
    TargetWorkbook.VBProject.VBComponents.Import FilePathFrM
End Sub

Function LIBRARY_FOLDER() As String
    If GetMotherBoardProp = VBARC_MOTHERBOARD Then
        LIBRARY_FOLDER = "C:\Users\acer\Documents\GitHub\VBA-Library\"
    Else
        LIBRARY_FOLDER = Environ$("USERPROFILE") & "\Documents\vbArc\Library\"
    End If
End Function

Function LastCell(rng As Range, Optional booCol As Boolean) As Range
    Dim WS As Worksheet
    Set WS = rng.Parent
    Dim Cell As Range
    If booCol = False Then
        Set Cell = WS.Cells(rows.Count, rng.Column).End(xlUp)
        If Cell.MergeCells Then Set Cell = Cells(Cell.Row + Cell.rows.Count - 1, Cell.Column)
    Else
        Set Cell = WS.Cells(rng.Row, Columns.Count).End(xlToLeft)
        If Cell.MergeCells Then Set Cell = Cells(Cell.Row, Cell.Column + Cell.Columns.Count - 1)
    End If

    Set LastCell = Cell
End Function

Public Function Len2( _
    ByVal val As Variant) _
    As Integer
    If IsArray(val) And Right(TypeName(val), 2) = "()" Then
        Len2 = UBound(val) - LBound(val) + 1
    ElseIf TypeName(val) = "String" Then
        Len2 = Len(val)
    ElseIf IsNumeric(val) Then
        Len2 = Len(CStr(val))
    Else
        Len2 = val.Count
    End If
End Function




Function LinkedClasses( _
                      TargetWorkbook As Workbook, _
                      module As VBComponent, _
                      procedure As String) As Collection

    Dim coll As New Collection
    Dim var As Variant
        var = classCallsOfModule(module)
    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, procedure)
    Dim keyword As String
    Dim ClassName As String
    Dim Element As Variant
    Dim i As Long
    On Error Resume Next
    For i = LBound(var, 1) To UBound(var, 1)
        If InStr(1, Code, var(i, 1)) > 0 Or InStr(1, Code, var(i, 2)) > 0 Then
            coll.Add var(i, 1), var(i, 1)
        End If
    Next
    For Each Element In ClassNames
        If InStr(1, Code, Element) > 0 Then
            coll.Add Element, CStr(Element)
        End If
    Next
    On Error GoTo 0
    Set LinkedClasses = coll
End Function

Function LinkedDeclarations( _
                           Optional TargetWorkbook As Workbook, _
                           Optional module As VBComponent, _
                           Optional procedure As String) As Collection

    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Stop
    
    DeclarationsTableCreate TargetWorkbook
    
    Dim TargetWorksheet As Worksheet: Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim coll As New Collection
    Dim Code As String: Code = ProcedureCode(TargetWorkbook, module, procedure)
    Dim Element
    For Each Element In DeclarationsTableKeywords
        If RegexTest(Code, "\b ?" & CStr(Element) & "\b") Then
            On Error Resume Next
            coll.Add CStr(Element), CStr(Element)
            On Error GoTo 0
        End If
    Next
    Set LinkedDeclarations = coll
End Function

Function LinkedProcedures( _
                         Optional TargetWorkbook As Workbook, _
                         Optional module As VBComponent, _
                         Optional ProcedureName As String) As Collection
    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop
    Dim Procedures As Collection
    Set Procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim coll As New Collection
    Dim procedure As Variant
    For Each procedure In Procedures
        If UCase(CStr(procedure)) <> UCase(CStr(ProcedureName)) Then
            If RegexTest(Code, "\W" & CStr(procedure) & "[.(\W]") = True Then
                coll.Add procedure, CStr(procedure)
            End If
        End If
    Next
    Set LinkedProcedures = coll
End Function

Function LinkedProceduresDeep( _
                             ProcedureName As Variant, _
                             TargetWorkbook As Workbook) As Collection

    Dim AllProcedures As Collection:       Set AllProcedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Processed As Collection:           Set Processed = New Collection
    Dim CalledProcedures As Collection:    Set CalledProcedures = New Collection

    Dim procedure As Variant
    Dim module As VBComponent

    Processed.Add CStr(ProcedureName), CStr(ProcedureName)
    On Error Resume Next
    For Each procedure In LinkedProcedures(TargetWorkbook, , CStr(ProcedureName))
    CalledProcedures.Add CStr(procedure), CStr(procedure)
    Next
    On Error GoTo 0

    Dim CalledProceduresCount As Long
        CalledProceduresCount = CalledProcedures.Count
    Dim Element
repeat:
    For Each Element In CalledProcedures
        If Not CollectionContains(Processed, , CStr(Element)) Then
            On Error Resume Next
            For Each procedure In LinkedProcedures(TargetWorkbook, , CStr(Element))
            CalledProcedures.Add CStr(procedure), CStr(procedure)
            Next
            On Error GoTo 0
            Processed.Add CStr(Element), CStr(Element)
        End If
    Next
    If CalledProcedures.Count > CalledProceduresCount Then
        CalledProceduresCount = CalledProcedures.Count
        GoTo repeat
    End If

    Set LinkedProceduresDeep = CollectionSort(CalledProcedures)
End Function


Sub LinkedProceduresMoveHere(Optional procedure As String)
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = ActiveCodepaneWorkbook
    If Not AssignProcedureVariable(TargetWorkbook, procedure) Then Exit Sub
    Dim el
    For Each el In LinkedProceduresDeep(procedure, TargetWorkbook)
        ProcedureMoveHere CStr(el)
    Next
End Sub




Function LinkedUserforms( _
                        TargetWorkbook As Workbook, _
                        module As VBComponent, _
                        procedure As String) As Collection
    Dim coll As New Collection
    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, procedure)
    Dim formName
    For Each formName In UserformNames(TargetWorkbook)
        If RegexTest(Code, "\W" & formName & "[.(\W]") = True Then coll.Add formName '& " " & "(Userform)"
    Next
    Set LinkedUserforms = coll
End Function

Function ModuleAddOrSet( _
                       TargetWorkbook As Workbook, _
                       TargetName As String, _
                       ModuleType As VBIDE.vbext_ComponentType) As VBComponent


    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim module As VBComponent
    On Error Resume Next
    Set module = TargetWorkbook.VBProject.VBComponents(TargetName)
    On Error GoTo 0
    If module Is Nothing Then
        Set module = TargetWorkbook.VBProject.VBComponents.Add(ModuleType)
        module.Name = TargetName
    End If
    Set ModuleAddOrSet = module
End Function




Function ModuleCode(module As VBComponent) As String
    With module.CodeModule
        If .CountOfLines = 0 Then ModuleCode = "": Exit Function
        ModuleCode = .Lines(1, .CountOfLines)
    End With
End Function

Public Function ModuleExists( _
                            TargetName As String, _
                            TargetWorkbook As Workbook) As Boolean
    Dim module As VBComponent
    On Error Resume Next
    Set module = TargetWorkbook.VBProject.VBComponents(TargetName)
    On Error GoTo 0
    ModuleExists = Not module Is Nothing
End Function

Public Function ModuleOfProcedure( _
                                 TargetWorkbook As Workbook, _
                                 ProcedureName As Variant) As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long, NumProc As Long
    Dim procedure As String
    Dim module As VBComponent
    For Each module In TargetWorkbook.VBProject.VBComponents
        With module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .CountOfLines
                procedure = .ProcOfLine(lineNum, ProcKind)
                If UCase(procedure) = UCase(ProcedureName) Then
                    Set ModuleOfProcedure = module
                    Exit Function
                End If
                lineNum = .ProcStartLine(procedure, ProcKind) + .ProcCountLines(procedure, ProcKind) + 1
            Loop
        End With
    Next module
End Function

Function ModuleOrSheetName(module As VBComponent) As String
    If module.Type = vbext_ct_Document Then
        If module.Name = "ThisWorkbook" Then
            ModuleOrSheetName = module.Name
        Else
            ModuleOrSheetName = GetSheetByCodeName(WorkbookOfModule(module), module.Name).Name
        End If
    Else
        ModuleName = module.Name
    End If
End Function

Function ModuleTypeToString(componentType As VBIDE.vbext_ComponentType) As String
    Select Case componentType
    Case vbext_ct_ActiveXDesigner
        ModuleTypeToString = "ActiveX Designer"
    Case vbext_ct_ClassModule
        ModuleTypeToString = "Class"
    Case vbext_ct_Document
        ModuleTypeToString = "Document"
    Case vbext_ct_MSForm
        ModuleTypeToString = "UserForm"
    Case vbext_ct_StdModule
        ModuleTypeToString = "Module"
    Case Else
        ModuleTypeToString = "Unknown Type: " & CStr(componentType)
    End Select
End Function

Function ProcedureAssignedModule( _
                                TargetWorkbook As Workbook, _
                                module As VBComponent, _
                                procedure As String) As VBComponent
        Dim ComponentName As Variant
        ComponentName = Split(ProcedureCode(TargetWorkbook, module, procedure), vbNewLine)
        ComponentName = Filter(ComponentName, "'@AssignedModule")
        If Len2(ComponentName) <> 1 Then Exit Function
        Dim UB As Long
        UB = UBound(Split(ComponentName(0), " "))
        ComponentName = Split(ComponentName(0), " ")(UB)
        Set ProcedureAssignedModule = ModuleAddOrSet(TargetWorkbook, CStr(ComponentName), vbext_ct_StdModule)
End Function

Sub ProcedureAssignedModuleAdd( _
                                Optional TargetWorkbook As Workbook, _
                                Optional module As VBComponent, _
                                Optional procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Stop
    ProcedureLinesRemove "'@AssignedModule *", TargetWorkbook, module, procedure
    module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, procedure), _
                                  "'@AssignedModule " & module.Name
End Sub

Function ProcedureBodyLineFirst( _
                               module As VBComponent, _
                               procedure As String) As Long
    ProcedureBodyLineFirst = ProcedureTitleLineFirst(module, procedure) + ProcedureTitleLineCount(module, procedure)
End Function

Function ProcedureBodyLineFirstAfterComments( _
                                            module As VBComponent, _
                                            procedure As String) As Long
    Dim N As Long
    Dim s As String
    For N = ProcedureBodyLineFirst(module, procedure) To module.CodeModule.CountOfLines
        s = Trim(module.CodeModule.Lines(N, 1))
        If s = vbNullString Then
            Exit For
        ElseIf Left(s, 1) = "'" Then
        ElseIf Left(s, 3) = "Rem" Then
        ElseIf Right(Trim(module.CodeModule.Lines(N - 1, 1)), 1) = "_" Then
        ElseIf Right(s, 1) = "_" Then
        Else
            Exit For
        End If
    Next N
    ProcedureBodyLineFirstAfterComments = N
End Function



Public Function ProcedureCode( _
                             Optional TargetWorkbook As Workbook, _
                             Optional module As VBComponent, _
                             Optional procedure As Variant, _
                             Optional IncludeHeader As Boolean = True) As String
    If Not AssignCPSvariables(TargetWorkbook, module, CStr(procedure)) Then Exit Function
    Dim lProcStart            As Long
    Dim lProcBodyStart        As Long
    Dim lProcNoLines          As Long
    Const vbext_pk_Proc = 0
    On Error GoTo Error_Handler
    lProcStart = module.CodeModule.ProcStartLine(procedure, vbext_pk_Proc)
    lProcBodyStart = module.CodeModule.ProcBodyLine(procedure, vbext_pk_Proc)
    lProcNoLines = module.CodeModule.ProcCountLines(procedure, vbext_pk_Proc)
    If IncludeHeader = True Then
        ProcedureCode = module.CodeModule.Lines(lProcStart, lProcNoLines)
    Else
        lProcNoLines = lProcNoLines - (lProcBodyStart - lProcStart)
        ProcedureCode = module.CodeModule.Lines(lProcBodyStart, lProcNoLines)
    End If
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    "Error Source: ProcedureCode" & vbCrLf & _
    "Error Description: " & err.Description & _
    Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Line No: " & Erl)
    Resume Error_Handler_Exit
End Function

Function ProcedureExists( _
                        TargetWorkbook As Workbook, _
                        ProcedureName As Variant) As Boolean
    Dim Procedures As Collection
    Set Procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim procedure As Variant
    For Each procedure In Procedures
        If UCase(CStr(procedure)) = UCase(ProcedureName) Then
            ProcedureExists = True
            Exit Function
        End If
    Next
End Function

Function ProcedureLastModAdd( _
                            Optional TargetWorkbook As Workbook, _
                            Optional module As VBComponent, _
                            Optional procedure As String, _
                            Optional ModificationDate As Double)
                       


If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Exit Function
    If ModificationDate = 0 Then ModificationDate = Format(Now, "yymmddhhnn")
    Dim LastModLine As Long
        LastModLine = ProcedureLineContaining(module, procedure, "'@LastModified *")
    If LastModLine = 0 Then GoTo PASS
    Dim LDate As Double
        LDate = Split(module.CodeModule.Lines(LastModLine, 1), " ")(1)
    ProcedureLinesRemove "'@LastModified *", TargetWorkbook, module, procedure
PASS:
    module.CodeModule.InsertLines ProcedureBodyLineFirst(module, procedure), _
                                  "'@LastModified " & ModificationDate
    
    ProcedureLastModAdd = ModificationDate
End Function

Function ProcedureLastModified( _
                            Optional TargetWorkbook As Workbook, _
                            Optional module As VBComponent, _
                            Optional procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Stop
    ProcedureLastModified = StringLastModified(ProcedureCode(TargetWorkbook, module, procedure))
End Function

Function ProcedureLinesCount( _
                            module As VBComponent, _
                            procedure As String) As Long
    ProcedureLinesCount = module.CodeModule.ProcCountLines(procedure, vbext_pk_Proc)
End Function

Public Function ProcedureLinesFirst( _
                                   module As VBComponent, _
                                   procedure As String) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
        ProcKind = vbext_pk_Proc
    ProcedureLinesFirst = module.CodeModule.ProcStartLine(procedure, ProcKind)
End Function


Public Function ProcedureLinesLast( _
                                  module As VBComponent, _
                                  procedure As String, _
                                  Optional IncludeTail As Boolean) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
        ProcKind = vbext_pk_Proc
    Dim startAt As Long
        startAt = module.CodeModule.ProcStartLine(procedure, ProcKind)
    Dim CountOf As Long
        CountOf = module.CodeModule.ProcCountLines(procedure, ProcKind)
    Dim endAt As Long
        endAt = startAt + CountOf - 1
    If Not IncludeTail Then
        Do While Not Trim(module.CodeModule.Lines(endAt, 1)) Like "End *"
            endAt = endAt - 1
        Loop
    End If
    ProcedureLinesLast = endAt
End Function

Sub ProcedureLinesRemove( _
                        myCriteria As String, _
                        Optional TargetWorkbook As Workbook, _
                        Optional module As VBComponent, _
                        Optional procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Stop

    Dim Code As String
    Dim i As Long
    For i = ProcedureLinesLast(module, procedure) To ProcedureLinesFirst(module, procedure) Step -1
        Code = Trim(module.CodeModule.Lines(i, 1))
        If Code Like myCriteria Then module.CodeModule.DeleteLines i
    Next
End Sub

Sub ProcedureLinesRemoveInclude( _
                                Optional TargetWorkbook As Workbook, _
                                Optional module As VBComponent, _
                                Optional procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Stop
    ProcedureLinesRemove "'@INCLUDE", TargetWorkbook, module, procedure
End Sub


Sub ProcedureMoveHere( _
                     Optional procedure As String)

    
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = ActiveCodepaneWorkbook
    If Not AssignProcedureVariable(TargetWorkbook, procedure) Then Exit Sub
    Dim module As VBComponent
    Set module = ModuleOfProcedure(TargetWorkbook, procedure)
    Dim s As String
        s = ProcedureCode(TargetWorkbook, module, procedure)

        If InStr(1, s, "'@AssignedModule") = 0 Then
            ProcedureAssignedModuleAdd TargetWorkbook, module, procedure
            s = ProcedureCode(TargetWorkbook, module, procedure)
        End If

    Dim sl As Long, cl As Long
        sl = ProcedureLinesFirst(module, procedure)
        cl = ProcedureLinesLast(module, procedure, False) - sl + 1
    ActiveModule.CodeModule.InsertLines ProcedureLinesLast(module, ActiveProcedure, True) + 1, s
    module.CodeModule.DeleteLines sl, cl
End Sub

Sub ProcedureMoveToAssignedModule( _
                                 Optional TargetWorkbook As Workbook, _
                                 Optional module As VBComponent, _
                                 Optional procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, procedure) Then Exit Sub
    Dim MoveToModule As VBComponent
    Set MoveToModule = ProcedureAssignedModule(TargetWorkbook, module, procedure)
    If MoveToModule Is Nothing Then Exit Sub
    ProcedureMoveToModule TargetWorkbook, module, procedure, MoveToModule
End Sub

Sub ProcedureMoveToModule( _
                         TargetWorkbook As Workbook, _
                         module As VBComponent, _
                         procedure As String, _
                         MoveToModule As VBComponent)
    Dim Code As String
        Code = ProcedureCode(TargetWorkbook, module, procedure)
    Dim startLine As Long
        startLine = ProcedureLinesFirst(module, procedure)
    Dim CountLines As Long
        CountLines = ProcedureLinesCount(module, procedure)
    MoveToModule.CodeModule.InsertLines MoveToModule.CodeModule.CountOfLines + 1, vbNewLine & Code
    module.CodeModule.DeleteLines startLine, CountLines

End Sub

Public Sub ProcedureReplace( _
                            module As VBComponent, _
                            procedure As String, _
                            Code As String)

    Dim startLine As Integer
    Dim NumLines As Integer
    With module.CodeModule
        startLine = .ProcStartLine(procedure, vbext_pk_Proc)
        NumLines = .ProcCountLines(procedure, vbext_pk_Proc)
        .DeleteLines startLine, NumLines
        .InsertLines startLine, Code
    End With
End Sub

Function ProcedureTitle( _
                       module As VBComponent, _
                       procedure As String) As String
    Dim titleLine As Long
        titleLine = ProcedureTitleLineFirst(module, procedure)
    Dim title As String
        title = module.CodeModule.Lines(titleLine, 1)
    Dim counter As Long
        counter = 1
    Do While Right(title, 1) = "_"
        counter = counter + 1
        title = module.CodeModule.Lines(titleLine, counter)
    Loop

    ProcedureTitle = title
End Function

Function ProcedureTitleLineCount( _
                                module As VBComponent, _
                                procedure As String) As Long

    ProcedureTitleLineCount = ProcedureTitleLineLast(module, procedure) - ProcedureTitleLineFirst(module, procedure) + 1
End Function



Public Function ProcedureTitleLineFirst( _
                                       module As VBComponent, _
                                       procedure As String) As Long
    ProcedureTitleLineFirst = module.CodeModule.ProcBodyLine(procedure, vbext_pk_Proc)
End Function

Function ProcedureTitleLineLast( _
                               module As VBComponent, _
                               procedure As String) As Long
    ProcedureTitleLineLast = ProcedureTitleLineFirst(module, procedure) + UBound(Split(ProcedureTitle(module, procedure), vbNewLine))
End Function

Public Function ProceduresOfModule( _
                                  module As VBComponent) As Collection
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long
    Dim coll As New Collection
    Dim procedure As String
    With module.CodeModule
        lineNum = .CountOfDeclarationLines + 1
        Do Until lineNum >= .CountOfLines
            ProcedureAs = .ProcOfLine(lineNum, ProcKind)
            coll.Add ProcedureAs
            lineNum = .ProcStartLine(ProcedureAs, ProcKind) + .ProcCountLines(ProcedureAs, ProcKind) + 1
        Loop
    End With
    Set ProceduresOfModule = coll
End Function

Function ProceduresOfTXT( _
                        Code As String) As Collection


    Code = Replace(Code, vbNewLine, vbLf)
    Dim var
        var = Split(Code, vbLf)

    Dim out
        out = ArrayAppend(Filter(var, "Sub" & Space(1), True, vbBinaryCompare), Filter(var, "Function ", True, vbBinaryCompare))
    If TypeName(out) = "Empty" Then Exit Function
    out = Filter(out, "(", True)
    out = Filter(out, "Declare", False)
    out = Filter(out, Chr(34) & "Sub", False)
    out = Filter(out, Chr(34) & "Function", False)
    out = Filter(out, "End Sub", False)
    out = Filter(out, "End Function", False)

    Dim i As Long
    For i = LBound(out) To UBound(out)
        out(i) = Left(out(i), InStr(1, out(i), "(") - 1)
        out(i) = Replace(out(i), "Private ", "")
        out(i) = Replace(out(i), "Public ", "")
        out(i) = Replace(out(i), "Sub ", "")
        out(i) = Replace(out(i), "Function ", "")
        If UBound(Split(out(i), " ")) > 0 Then
            out(i) = ""
        End If
    Next

    ArrayQuickSort out
    out = cleanArray(out)
    out = ArrayDuplicatesRemove(out)
    Set ProceduresOfTXT = ArrayToCollection(out)
End Function

Function ProceduresOfWorkbook( _
                             TargetWorkbook As Workbook, _
                             Optional ExcludeDocument As Boolean = True, _
                             Optional ExcludeClass As Boolean = True, _
                             Optional ExcludeForm As Boolean = True) As Collection
    Dim module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    For Each module In TargetWorkbook.VBProject.VBComponents
        If ExcludeClass = True And module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If ExcludeDocument = True And module.Type = vbext_ct_Document Then GoTo SKIP
        If ExcludeForm = True And module.Type = vbext_ct_MSForm Then GoTo SKIP
        With module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .CountOfLines
                ProcedureName = .ProcOfLine(lineNum, ProcKind)
                If InStr(1, ProcedureName, "_") = 0 Then
                    coll.Add ProcedureName
                End If
                lineNum = .ProcStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
SKIP:
    Next module
    Set ProceduresOfWorkbook = coll
End Function

Sub ProjetFoldersCreate()
    Dim Element
    For Each Element In vbarcFolders
        FoldersCreate CStr(Element)
    Next
End Sub

Public Function RegexTest( _
                         ByVal string1 As String, _
                         ByVal stringPattern As String, _
                         Optional ByVal globalFlag As Boolean, _
                         Optional ByVal ignoreCaseFlag As Boolean, _
                         Optional ByVal multilineFlag As Boolean) As Boolean
    Dim REGEX As Object
    Set REGEX = CreateObject("VBScript.RegExp")
    With REGEX
        .Global = globalFlag
        .IgnoreCase = ignoreCaseFlag
        .MultiLine = multilineFlag
        .Pattern = stringPattern
    End With
    RegexTest = REGEX.test(string1)
End Function


Function StringLastModified(txt As String)

    Dim Code As Variant
        Code = Filter(Split(txt, vbLf), "'@LastModified ")
    If ArrayAllocated(Code) Then
        Dim lastDate As Variant
        If Trim(Code(0)) Like "'@LastModified *" Then
            lastDate = Split(Code(0), " ")(1)
            lastDate = DateSerial(Left(lastDate, 2), Mid(lastDate, 3, 2), Mid(lastDate, 5, 2)) _
                       & " " & TimeSerial(Mid(lastDate, 7, 2), Mid(lastDate, 9, 2), 0)
            StringLastModified = Split(Code(0), " ")(1)
        End If
    Else

    End If
End Function



Function StringProcedureAssignedModule(txt As String) As String
        Dim ComponentName As Variant
        ComponentName = Split(txt, vbLf)
        ComponentName = Filter(ComponentName, "'@AssignedModule")
        If Not ArrayAllocated(ComponentName) Then Exit Function
        Dim UB As Long
            UB = UBound(Split(ComponentName(0), " "))
        ComponentName = Split(ComponentName(0), " ")(UB)
        StringProcedureAssignedModule = ComponentName
End Function



Function TXTReadFromUrl(url As String) As String
    On Error GoTo Err_GetFromWebpage
    Dim objWeb As Object
    Dim strXML As String
    Set objWeb = CreateObject("Msxml2.ServerXMLHTTP")
    objWeb.Open "GET", url, False
    objWeb.setRequestHeader "Content-Type", "text/xml"
    objWeb.setRequestHeader "Cache-Control", "no-cache"
    objWeb.setRequestHeader "Pragma", "no-cache"
    objWeb.send
    Do While objWeb.readyState <> 4
        DoEvents
    Loop
    strXML = objWeb.responseText
    TXTReadFromUrl = strXML
End_GetFromWebpage:
    Set objWeb = Nothing
    Exit Function
Err_GetFromWebpage:
    MsgBox Err.Description & " (" & Err.Number & ")"
    Resume End_GetFromWebpage
End Function

Sub TxtOverwrite(sFile As String, sText As String)
    On Error GoTo ERR_HANDLER
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open sFile For Output As #FileNumber
    Print #FileNumber, sText
    Close #FileNumber
Exit_Err_Handler:
    Exit Sub
ERR_HANDLER:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
    "Error Number: " & Err.Number & vbCrLf & _
    "Error Source: TxtOverwrite" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Sub

Sub TxtPrepend(filePath As String, txt As String)
    Dim s As String
    s = TxtRead(filePath)
    TxtOverwrite filePath, txt & vbNewLine & s
End Sub



Sub TxtPrependContainedProcedures(FileName As String)
    Dim s As String: s = TxtRead(FileName)
    Dim v As New Collection
    Set v = ProceduresOfTXT(s)
    If v.Count = 0 Then Exit Sub
    Dim Line As String: Line = String(30, "'")
    TxtPrepend FileName, _
    "'Contains the following " & "#" & v.Count & " procedures " & vbNewLine & Line & vbNewLine & _
    "'" & collectionToString(v, vbNewLine & "'") & vbNewLine & Line & vbNewLine & vbNewLine
End Sub

Function TxtRead(sPath As Variant) As String
    Dim sTXT As String
    If Dir(sPath) = "" Then
        Debug.Print "File was not found."
        Debug.Print sPath
        Exit Function
    End If
    Open sPath For Input As #1
    Do Until EOF(1)
        Line Input #1, sTXT
        TxtRead = TxtRead & sTXT & vbLf
    Loop
    Close
    If Len(TxtRead) = 0 Then
        TxtRead = ""
    Else
        TxtRead = Left(TxtRead, Len(TxtRead) - 1)
    End If
End Function

Function URLExists(url) As Boolean
    Dim Request As Object
    Dim ff As Integer
    Dim rc As Variant

    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Request
      .Open "GET", url, False
      .send
      rc = .statusText
    End With
    Set Request = Nothing
    If rc = "OK" Then URLExists = True

    Exit Function
EndNow:
End Function

Function UserformNames(TargetWorkbook As Workbook)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Set UserformNames = ComponentNames(vbext_ct_MSForm, TargetWorkbook)
End Function






Function WorkbookCode(TargetWorkbook) As String
    If TypeName(TargetWorkbook) <> "Workbook" Then Stop
    Dim module As VBComponent
    Dim txt
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.CodeModule.CountOfLines > 0 Then
            txt = txt & _
                  vbNewLine & _
                  "'" & String(10, "=") & ModuleOrSheetName(module) & " (" & module.Type & ") " & String(10, "=") & _
                  vbNewLine & _
                  ModuleCode(module)
        End If
    Next
    WorkbookCode = txt
End Function


Function WorkbookOfModule(vbComp As VBComponent) As Workbook
    Set WorkbookOfModule = WorkbookOfProject(vbComp.Collection.Parent)
End Function

Function WorkbookOfProject(vbProj As VBProject) As Workbook
    tmpStr = vbProj.FileName
    tmpStr = Right(tmpStr, Len(tmpStr) - InStrRev(tmpStr, "\"))
    Set WorkbookOfProject = Workbooks(tmpStr)
End Function



Function WorksheetExists(SheetName As String, TargetWorkbook As Workbook) As Boolean
    Dim TargetWorksheet  As Worksheet
    On Error Resume Next
    Set TargetWorksheet = TargetWorkbook.Sheets(SheetName)
    On Error GoTo 0
    WorksheetExists = Not TargetWorksheet Is Nothing
End Function

Function classCallsOfModule(module As VBComponent) As Variant


    Dim Code As Variant
    Dim Element As Variant
    Dim keyword As Variant
    Dim var As Variant
    ReDim var(1 To 2, 1 To 1)
    Dim counter As Long
    counter = 0
    If module.CodeModule.CountOfDeclarationLines > 0 Then
        Code = module.CodeModule.Lines(1, module.CodeModule.CountOfDeclarationLines)
        Code = Replace(Code, "_" & vbNewLine, "")
        Code = Split(Code, vbNewLine)
        Code = Filter(Code, " As ", , vbTextCompare)
        For Each Element In Code
            Element = Trim(Element)
            If Element Like "* As *" Then
                keyword = Split(Element, " As ")(0)
                keyword = Split(keyword, " ")(UBound(Split(keyword, " ")))
                Element = Split(Element, " As ")(1)
                Element = Replace(Element, "New ", "")
                
                For Each ClassName In ClassNames
                    If Element = ClassName Then
                        
                        ReDim Preserve var(1 To 2, 1 To counter + 1)
                        var(1, UBound(var, 2)) = Element
                        var(2, UBound(var, 2)) = keyword
                        counter = counter + 1
                    End If
                Next
            End If
        Next
        If var(1, 1) <> "" Then

            If UBound(var, 2) > 1 Then
                classCallsOfModule = WorksheetFunction.Transpose(var)
            Else
                Dim var2(1 To 1, 1 To 2)
                var2(1, 1) = var(1, 1)
                var2(1, 2) = var(2, 1)
                classCallsOfModule = var2
            End If
        End If
    End If

End Function

Function collectionToString(coll As Collection, delim As String) As String
    Dim Element
    Dim out As String
    For Each Element In coll
        out = IIf(out = "", Element, out & delim & Element)
    Next
    collectionToString = out
End Function

Function getDeclarations( _
                        wb As Workbook, _
                        Optional includeScope As Boolean, _
                        Optional includeType As Boolean, _
                        Optional includeKeywords As Boolean, _
                        Optional includeDeclarations As Boolean, _
                        Optional includeComponentName As Boolean, _
                        Optional includeComponentType As Boolean) As Collection

    Dim ComponentCollection     As New Collection
    Dim ComponentTypecollection As New Collection
    Dim DeclarationsCollection  As New Collection
    Dim KeywordsCollection      As New Collection
    Dim Output                  As New Collection
    Dim ScopeCollection         As New Collection
    Dim TypeCollection          As New Collection

    Dim Element                 As Variant
    Dim OriginalDeclarations    As Variant
    Dim Str                     As Variant
    
    Dim Tmp                     As String
    Dim Helper                  As String
    Dim i                       As Long
    
    Dim module                  As VBComponent
    For Each module In wb.VBProject.VBComponents
        If module.Type = vbext_ct_StdModule Or module.Type = vbext_ct_MSForm Then
            If module.CodeModule.CountOfDeclarationLines > 0 Then
                Str = module.CodeModule.Lines(1, module.CodeModule.CountOfDeclarationLines)
                Str = Replace(Str, "_" & vbNewLine, "")
                OriginalDeclarations = Str
                Tmp = Str
                Do While InStr(1, Str, "End Type") > 0
                    Tmp = Mid(Str, InStr(1, Str, "Type "), InStr(1, Str, "End Type") - InStr(1, Str, "Type ") + 8)
                    Str = Replace(Str, Tmp, Split(Tmp, vbNewLine)(0))
                Loop
                Do While InStr(1, Str, "End Enum") > 0
                    Tmp = Mid(Str, InStr(1, Str, "Enum "), InStr(1, Str, "End Enum") - InStr(1, Str, "Enum ") + 8)
                    Str = Replace(Str, Tmp, Split(Tmp, vbNewLine)(0))
                Loop
                Do While InStr(1, Str, "  ") > 0
                    Str = Replace(Str, "  ", " ")
                Loop
                
                Str = Split(Str, vbNewLine)
                Tmp = OriginalDeclarations
                
                For Each Element In Str
                    If Len(CStr(Element)) > 0 And Not Trim(CStr(Element)) Like "'*" And Not Trim(CStr(Element)) Like "Rem*" Then
                        If RegexTest(CStr(Element), "\b ?Enum \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(Element), " ", "Enum")
                            DeclarationsCollection.Add DeclarationsKeywordSubstring(Tmp, , "Enum " & KeywordsCollection.Item(KeywordsCollection.Count), "End Enum", , , True)
                            TypeCollection.Add "Enum"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(Element), "\b ?Type \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(Element), " ", "Type")
                            DeclarationsCollection.Add DeclarationsKeywordSubstring(Tmp, , "Type " & KeywordsCollection.Item(KeywordsCollection.Count), "End Type", , , True)
                            TypeCollection.Add "Type"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf InStr(1, CStr(Element), "Const ", vbTextCompare) > 0 Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(Element), " ", "Const")
                            DeclarationsCollection.Add CStr(Element)
                            TypeCollection.Add "Const"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(Element), "\b ?Sub \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(Element), " ", "Sub")
                            DeclarationsCollection.Add CStr(Element)
                            TypeCollection.Add "Sub"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(Element), "\b ?Function \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(Element), " ", "Function")
                            DeclarationsCollection.Add CStr(Element)
                            TypeCollection.Add "Function"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf Element Like "*(*) As *" Then
                            Helper = Left(Element, InStr(1, CStr(Element), "(") - 1)
                            Helper = Mid(Helper, InStrRev(Helper, " ") + 1)
                            KeywordsCollection.Add Helper
                            DeclarationsCollection.Add CStr(Element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf Element Like "* As *" Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(Element), " ", , "As")
                            DeclarationsCollection.Add CStr(Element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.Item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        Else
                        End If
                    End If
                Next Element
            End If
        End If
    Next module
    
    If includeScope = True Then Output.Add ScopeCollection
    If includeType = True Then Output.Add TypeCollection
    If includeKeywords = True Then Output.Add KeywordsCollection
    If includeDeclarations = True Then Output.Add DeclarationsCollection
    If includeComponentType = True Then Output.Add ComponentTypecollection
    If includeComponentName = True Then Output.Add ComponentCollection
    
    Set getDeclarations = Output
End Function

Function getLastRow(TargetSheet As Worksheet)
    Dim LastCell As Range
    Set LastCell = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    getLastRow = LastCell.Row
End Function

Function vbarcFolders() As Collection



    Dim coll As New Collection
    coll.Add GITHUB_LOCAL_LIBRARY_PROCEDURES
    coll.Add GITHUB_LOCAL_LIBRARY_CLASSES
    coll.Add GITHUB_LOCAL_LIBRARY_USERFORMS
    coll.Add GITHUB_LOCAL_LIBRARY_DECLARATIONS

    coll.Add Environ$("USERPROFILE") & "\Documents\vbArc\oleVba\"
    coll.Add Environ$("USERPROFILE") & "\Documents\vbArc\MergedTXT\"
    coll.Add Environ$("USERPROFILE") & "\Documents\vbArc\MemoryKnots\"
    coll.Add Environ$("USERPROFILE") & "\Documents\vbArc\ExportedImages\"
    Set vbarcFolders = coll
End Function

Function ProcedureLineContaining(module As VBComponent, procedure As String, This As String) As Long
    Dim i As Long
    For i = ProcedureLinesFirst(module, procedure) To ProcedureLinesLast(module, procedure)
        If module.CodeModule.Lines(i, 1) Like This Then
            ProcedureLineContaining = i
            Exit Function
        End If
    Next
End Function

Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function

Public Function CLIP(Optional StoreText As String) As String
    Dim X As Variant
    X = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                .SetData "text", X
            Case Else
                CLIP = .GetData("text")
            End Select
        End With
    End With
End Function

