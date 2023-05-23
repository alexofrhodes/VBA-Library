VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistryEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Option Compare Text
'
' Registry Class
' |* Brian J. Satola
' |* https://chejunkie.com/knowledge-base/registry-editor-class-vba/
'==========================================================================================================================
' Based on the work of Chip Pearson, www.cpearson.com, chip@cpearson.com.
'==========================================================================================================================
' This class provides several methods related to working with keys and values in the system registry.
' These routines rely on each other, so you should import this entire class module into your project
'   rather than just copying and pasting individual procedures into other modules.
' In all functions with a BaseKey parameter, the value of BaseKey must be either HKEY_CURRENT_USER (or HKCU)
'   or HKEY_LOCAL_MACHINE (or HKML). Any other value is invalid.
'--------------------------------------------------------------------------------------------------------------------------
' The registry is a tree structure; as such, the following variables are used:
'   |* BaseKey       ::   tree trunk (or specific tree [or hive] of information in your registry).
'   |* KeyName       ::   a specific branch on the trunk.
'   |* SubKeys       ::   sub-branches of a key-branch on the trunk.
'   |* ValueName     ::   a specific leaf on the branch.
'   |* ValueValue    ::   the value that is actually written on the leaf.
'
'==========================================================================================================================
' Public Subs, Functions and Properties - oh my!
'==========================================================================================================================
'   |* Get | --- About                          :: String that describes the current class.
'   |* --------- CloseRegEdit                   ::
'   |* --------- GetBaseKeyName                 ::
'   |* --------- GetBaseKeyNameShort            ::
'   |* --------- GetDataTypeName                :: Returns REG_DATA_TYPE as a string e.g. needed for WShRegKeyWrite.
'   |* Get | --- Name                           :: Returns the name# of the class.
'   |* --------- IsRegEditOpen                  ::
'   |* --------- OpenRegEdit                    :: Opens the Registry Editor (C:\Windows\regedit.exe).
'   |* --------- OpenRegEditToKey               ::
'   |* --------- RegistryCreateKey              :: Creates registry key in specified basekey (registry hive).
'   |* --------- RegistryCreateValue            :: Creates registry value in specified key of the basekey.
'   |* --------- RegistryDeleteKey              :: Deletes the registry key in basekey, including all subkeys and values.
'   |* --------- RegistryDeleteValue            :: Deletes registry value in specified key of the basekey.
'   |* --------- RegistryGetValue               :: Returns the value of the specified name in the key of the basekey.
'   |* --------- RegistryGetValueType           :: Returns the data type of the value name in the key of the basekey.
'   |* --------- RegistryKeyExists              :: Returns TRUE or FALSE indicating whether the key exists in the basekey.
'   |* --------- RegistryNumberOfSubKeys        :: Returns the number of sub-keys under the specified key in basekey.
'   |* --------- RegistryNumberOfValues         :: Returns the number of values under the specified key in basekey.
'   |* --------- RegistrySubKeyNamesToArray     :: Returns an array of names of all sub-keys under a key in basekey.
'   |* --------- RegistryUpdateValue            :: Updates the value of the specified name in the key of basekey.
'   |* --------- RegistryValueExists            :: Returns TRUE or FALSE indicating whether the value name exists.
'   |* --------- RegistryValueNamesToArray      :: Returns an array of names of all values under the key in basekey.
'   |* Get | --- Version                        :: Returns version string for the class [e.g. #.# (year)].
'
'==========================================================================================================================
' References
'==========================================================================================================================
Private m_AppErr As ApplicationError
'--Private m_SysErr As New SystemError
'
' Library API errors
'--------------------------------------------------------------------------------------------------------------------------
'   |* Get | --- SysErrNumber       ::      Returns the system error number, usually the value of Err.LastDllError
'   |* Get | --- SysErrDescription  ::      Returns the text description associated with m_RegSysErrNum,
'                                           the text returned from GetSystemErrorMessageText
'
'==========================================================================================================================
' Application constants
'==========================================================================================================================
Private Const C_NAME As String = "RegistryEditor.cls"

Private Type RegValue
    valueName   As String
    valueValue  As Variant
End Type

'==========================================================================================================================
' Private variables
'==========================================================================================================================
'--Private m_RegSysErrNum           As Long
'--Private m_RegSysErrText          As String

'==========================================================================================================================
' Error constants
'==========================================================================================================================
Const C_ERR_OFFSET = 0
Private Const C_ERR_NO_ERROR                    As Long = 0
Private Const C_ERR_INVALID_BASE_KEY            As Long = C_ERR_OFFSET + vbObjectError + 1
Private Const C_ERR_INVALID_DATA_TYPE           As Long = C_ERR_OFFSET + vbObjectError + 2
Private Const C_ERR_KEY_NOT_FOUND               As Long = C_ERR_OFFSET + vbObjectError + 3
Private Const C_ERR_VALUE_NOT_FOUND             As Long = C_ERR_OFFSET + vbObjectError + 4
Private Const C_ERR_DATA_TYPE_MISMATCH          As Long = C_ERR_OFFSET + vbObjectError + 5
Private Const C_ERR_ENTRY_LOCKED                As Long = C_ERR_OFFSET + vbObjectError + 6
Private Const C_ERR_INVALID_KEYNAME             As Long = vbObjectError + C_ERR_OFFSET + 7
Private Const C_ERR_UNABLE_TO_OPEN_KEY          As Long = vbObjectError + C_ERR_OFFSET + 8
Private Const C_ERR_UNABLE_TO_READ_KEY          As Long = vbObjectError + C_ERR_OFFSET + 9
Private Const C_ERR_UNABLE_TO_CREATE_KEY        As Long = vbObjectError + C_ERR_OFFSET + 10
Private Const C_ERR_UNABLE_TO_READ_VALUE        As Long = vbObjectError + C_ERR_OFFSET + 11
Private Const C_ERR_UNABLE_TO_UPDATE_VALUE      As Long = vbObjectError + C_ERR_OFFSET + 12
Private Const C_ERR_UNABLE_TO_CREATE_VALUE      As Long = vbObjectError + C_ERR_OFFSET + 13
Private Const C_ERR_UNABLE_TO_DELETE_KEY        As Long = vbObjectError + C_ERR_OFFSET + 14
Private Const C_ERR_UNABLE_TO_DELETE_VALUE      As Long = vbObjectError + C_ERR_OFFSET + 15
Private Const C_ERR_INVALID_PATH                As Long = vbObjectError + C_ERR_OFFSET + 16

'==========================================================================================================================
' API Constants
'==========================================================================================================================
'--Private Const HKEY_CURRENT_USER_        As Long = &H80000001
'--Private Const HKEY_LOCAL_MACHINE_       As Long = &H80000002
'--Private Const HKEY_CLASSES_ROOT_        As Long = &H80000000
'--Private Const HKEY_CURRENT_CONFIG_      As Long = &H80000005
'--Private Const HKEY_DYN_DATA_            As Long = &H80000006
'--Private Const HKEY_PERFORMANCE_DATA_    As Long = &H80000004
'--Private Const HKEY_USERS_               As Long = &H80000003
'--Private Const HKCU                      As Long = HKEY_CURRENT_USER_
'--Private Const HKLM                      As Long = HKEY_LOCAL_MACHINE_
Public Enum HKEY
    HKEY_CURRENT_USER_HKCU = &H80000001
    HKEY_LOCAL_MACHINE_HKLM = &H80000002
    HKEY_CLASSES_ROOT_HKCR = &H80000000
    HKEY_CURRENT_CONFIG_HKCC = &H80000005
    HKEY_DYN_DATA_HKDD = &H80000006
    HKEY_PERFORMANCE_DATA_HKPD = &H80000004
    HKEY_USERS_HKU = &H80000003
End Enum

Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const KEY_ALL_ACCESS            As Long = &H3F

Private Const REG_CREATED_NEW_KEY       As Long = &H1
Private Const REG_OPENED_EXISTING_KEY   As Long = &H2

Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL       As Long = &HFFFF

Private Const REG_OPTION_NON_VOLATILE   As Long = 0&
Private Const REG_OPTION_VOLATILE       As Long = &H1

Private Const ERROR_SUCCESS             As Long = 0&
Private Const ERROR_ACCESS_DENIED       As Long = 5
Private Const ERROR_INVALID_DATA        As Long = 13&
Private Const ERROR_MORE_DATA           As Long = 234        ' dderror
Private Const ERROR_NO_MORE_ITEMS       As Long = 259

Private Const S_OK                      As Long = &H0
Private Const MAX_DATA_BUFFER_SIZE      As Long = 1024
Private Const REGSTR_MAX_VALUE_LENGTH   As Long = &H100

'==========================================================================================================================
' API Types
'==========================================================================================================================
Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Boolean
End Type

Private Type FILETIME
    dwLowDateTime           As Long
    dwHighDateTime          As Long
End Type

Public Enum REG_DATA_TYPE
    REG_INVALID = -1        ' invalid
    REG_SZ = 1        ' string
    REG_EXPAND_SZ = 2        ' string
    REG_BINARY = 3        ' bianry
    REG_DWORD = 4        ' long
    REG_MULTI_SZ = 7        ' string
End Enum

Private Type ACL
    AclRevision             As Byte
    Sbz1                    As Byte
    AclSize                 As Integer
    AceCount                As Integer
    Sbz2                    As Integer
End Type

Private Type SECURITY_DESCRIPTOR
    Revision                As Byte
    Sbz1                    As Byte
    control                 As Long
    Owner                   As Long
    Group                   As Long
    Sacl                    As ACL
    Dacl                    As ACL
End Type

'==========================================================================================================================
' API Declarations
'==========================================================================================================================
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" _
(ByVal hiveKey As Long) As Long

Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
ByVal hiveKey As Long, _
ByVal lpSubKey As String, _
ByVal Reserved As Long, _
ByVal lpClass As String, _
ByVal dwOptions As Long, _
ByVal samDesired As Long, _
lpSecurityAttributes As SECURITY_ATTRIBUTES, _
phkResult As Long, _
lpdwDisposition As Long) As Long

Private Declare PtrSafe Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
ByVal hiveKey As Long, _
ByVal lpSubKey As String) As Long

Private Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
ByVal hiveKey As Long, _
ByVal lpSubKey As String, _
phkResult As Long) As Long

Private Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
ByVal hiveKey As Long, _
ByVal lpValueName As String) As Long

Private Declare PtrSafe Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
ByVal hiveKey As Long, _
ByVal dwIndex As Long, _
ByVal lpName As String, _
ByVal cbName As Long) As Long

Private Declare PtrSafe Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
ByVal hiveKey As Long, _
ByVal dwIndex As Long, _
ByVal lpName As String, _
lpcbName As Long, _
ByVal lpReserved As Long, _
ByVal lpClass As String, _
lpcbClass As Long, _
lpftLastWriteTime As FILETIME) As Long

Private Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" ( _
ByVal hiveKey As Long, _
ByVal dwIndex As Long, _
ByVal lpValueName As String, _
lpcbValueName As Long, _
ByVal lpReserved As Long, _
lpType As Long, _
lpData As Byte, _
lpcbData As Long) As Long

Private Declare PtrSafe Function RegFlushKey Lib "advapi32.dll" ( _
ByVal hiveKey As Long) As Long

Private Declare PtrSafe Function RegGetKeySecurity Lib "advapi32.dll" ( _
ByVal hiveKey As Long, _
ByVal SecurityInformation As Long, _
pSecurityDescriptor As SECURITY_DESCRIPTOR, _
lpcbSecurityDescriptor As Long) As Long

Private Declare PtrSafe Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" ( _
ByVal hiveKey As Long, _
ByVal lpClass As String, _
lpcbClass As Long, _
ByVal lpReserved As Long, _
lpcSubKeys As Long, _
lpcbMaxSubKeyLen As Long, _
lpcbMaxClassLen As Long, _
lpcValues As Long, _
lpcbMaxValueNameLen As Long, _
lpcbMaxValueLen As Long, _
lpcbSecurityDescriptor As Long, _
lpftLastWriteTime As FILETIME) As Long

Private Declare PtrSafe Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" ( _
ByVal hiveKey As Long, _
ByVal lpSubKey As String, _
ByVal lpValue As String, _
lpcbValue As Long) As Long

Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
ByVal hiveKey As Long, _
ByVal lpValueName As String, _
ByVal lpReserved As Long, _
lpType As Long, _
lpData As Any, _
lpcbData As Long) As Long

Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
ByVal hiveKey As Long, _
ByVal lpValueName As String, _
ByVal Reserved As Long, _
ByVal dwType As Long, _
lpData As Any, _
ByVal cbData As Long) As Long

Private Declare PtrSafe Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" ( _
ByVal hiveKey As Long, _
ByVal lpValueName As String, _
ByVal Reserved As Long, _
ByVal dwType As Long, _
ByVal szData As String, _
ByVal cbData As Long) As Long

Private Declare PtrSafe Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" ( _
ByVal hiveKey As Long, _
ByVal lpValueName As String, _
ByVal Reserved As Long, _
ByVal dwType As Long, _
szData As Long, _
ByVal cbData As Long) As Long

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
ByVal hiveKey As Long, _
ByVal lpSubKey As String, _
ByVal ulOptions As Long, _
ByVal samDesired As Long, _
phkResult As Long) As Long

Private Declare PtrSafe Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" ( _
ByVal hiveKey As Long, _
ByVal lpValueName As String, _
ByVal lpReserved As Long, _
ByRef lpType As Long, _
ByVal szData As String, _
ByRef lpcbData As Long) As Long


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''''ALEX additions to handle VBA Settings'''''
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Const VbaSettingsBasekey = "HKEY_CURRENT_USER\SOFTWARE\VB and VBA Program Settings"


Sub VBA_OpenSettings()
    
    Dim baseKey As String
    baseKey = VbaSettingsBasekey
    Dim wsh         As Object

    ResetErrorVariables

    '// Define open location.
    If Not RegistryUpdateValue(HKEY.HKEY_CURRENT_USER_HKCU, "Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", _
        "LastKey", "Computer\" & baseKey, createKeyIfNotExist:=False) Then
        m_AppErr.Number = C_ERR_INVALID_PATH
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        GoTo ErrHandler
    End If

    '// Option switch.
'    If closeBeforeOpening Then
        CloseRegEdit promptUserBeforeClosing:=False
'    End If

    '// Open instance of Registry Editor.
    On Error Resume Next
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run "regedit.exe -m", 1, False
    Set wsh = Nothing
    On Error GoTo 0
    Exit Sub

ErrHandler:
    m_AppErr.Source = "OpenRegEditToKey(...)"
    OpenRegEdit openToLastKey:=False
    m_AppErr.DisplayMessage

End Sub

Function VBA_SaveSetting( _
                            sAPPNAME As String, _
                            sSectionName As String, _
                            sKeyName As String, _
                            sSettingValue As String) As Boolean

On Error GoTo Error_Handler

    Call SaveSetting(sAPPNAME, sSectionName, sKeyName, sSettingValue)
    VBA_SaveSetting = True
    
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: SaveRegistrySetting" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function VBA_GetSetting(sAPPNAME As String, sSectionName As String, sKeyName As String) As String
'Returns "" if app, section or key are not found
'Always returns a VarType = vbString / TypeName = String (stored as REG_SZ), so use conversion function after retrieving as req'd
On Error GoTo Error_Handler

    VBA_GetSetting = GetSetting(sAPPNAME, sSectionName, sKeyName)
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: GetRegistrySetting" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function VBA_GetAllSettings(sAPPNAME As String, sSectionName As String) As Variant
On Error GoTo Error_Handler
    Dim aSectionSettings      As Variant
    Dim iCounter              As Long

    aSectionSettings = GetAllSettings(sAPPNAME, sSectionName)
    VBA_GetAllSettings = aSectionSettings

    If IsEmpty(aSectionSettings) = True Then
        VBA_GetAllSettings = Null
    Else
        For iCounter = LBound(VBA_GetAllSettings, 1) To UBound(VBA_GetAllSettings, 1)
            Debug.Print aSectionSettings(iCounter, 0), aSectionSettings(iCounter, 1)
        Next iCounter
    End If
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: GetAllRegistrySettings" & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function VBA_DeleteSetting(sAPPNAME As String, sSectionName As String, Optional sKeyName As String) As Boolean
    On Error GoTo Error_Handler

    If sKeyName = "" Then
        'Delete the entire section and all its keys
        Call DeleteSetting(sAPPNAME, sSectionName)
    Else
        'Delete a specific section/key
        Call DeleteSetting(sAPPNAME, sSectionName, sKeyName)
    End If
    VBA_DeleteSetting = True

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    If Err.Number <> 5 Then
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Source: DeleteRegistrySetting" & vbCrLf & _
               "Error Number: " & Err.Number & vbCrLf & _
               "Error Description: " & Err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    End If
    Resume Error_Handler_Exit
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''
''' Returns a string describing the class.
'''
Public Property Get About() As String

    About = "ChE Junkie VBA Registry class module, v" & Me.version & "." & VBA.vbCrLf
    About = About & "An extension of original work done by Chip Pearson (www.cpearson.com)." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details see:" & VBA.vbCrLf & "https://chejunkie.com/knowledge-base/registry-editor-class-vba/"

End Property

'''
''' Returns the application error object of the class.
'''
Public Property Get AppErr() As ApplicationError

    Set AppErr = m_AppErr

End Property

'''
''' Returns the *baseKey* name.
''' An empty string is returned if no match is found.
'''
Public Function GetBaseKeyName(baseKey As HKEY) As String

    Select Case baseKey
    Case HKEY.HKEY_CLASSES_ROOT_HKCR:       GetBaseKeyName = "HKEY_CLASSES_ROOT"
    Case HKEY.HKEY_CURRENT_USER_HKCU:       GetBaseKeyName = "HKEY_CURRENT_USER"
    Case HKEY.HKEY_LOCAL_MACHINE_HKLM:      GetBaseKeyName = "HKEY_LOCAL_MACHINE"
    Case HKEY.HKEY_USERS_HKU:               GetBaseKeyName = "HKEY_USERS"
    Case HKEY.HKEY_CURRENT_CONFIG_HKCC:     GetBaseKeyName = "HKEY_CURRENT_CONFIG"
    Case HKEY.HKEY_DYN_DATA_HKDD:           GetBaseKeyName = "HKEY_DYN_DATA"
    Case HKEY.HKEY_PERFORMANCE_DATA_HKPD:   GetBaseKeyName = "HKEY_PERFORMANCE_DATA"
    End Select

End Function

'''
''' Returns the abbreviated *baseKey* name, which is needed for working with WScript.Shell objects.
''' An empty string is returned if no match is found.
'''
Public Function GetBaseKeyNameShort(baseKey As HKEY) As String

    Select Case baseKey
    Case HKEY.HKEY_CLASSES_ROOT_HKCR:       GetBaseKeyNameShort = "HKCR"
    Case HKEY.HKEY_CURRENT_USER_HKCU:       GetBaseKeyNameShort = "HKCU"
    Case HKEY.HKEY_LOCAL_MACHINE_HKLM:      GetBaseKeyNameShort = "HKLM"
    Case HKEY.HKEY_USERS_HKU:               GetBaseKeyNameShort = "HKU"
    Case HKEY.HKEY_CURRENT_CONFIG_HKCC:     GetBaseKeyNameShort = "HKCC"
    Case HKEY.HKEY_DYN_DATA_HKDD:           GetBaseKeyNameShort = "HKDD"
    Case HKEY.HKEY_PERFORMANCE_DATA_HKPD:   GetBaseKeyNameShort = "HKPD"
    End Select

End Function

'''
''' Closes all instances of regedit.exe.
''' By defaulat, the user is prompted before closing open instances of the Registry Editor.
'''
Public Sub CloseRegEdit(Optional promptUserBeforeClosing As Boolean = True)

    Dim cReg                As Collection
    Dim proc                As Object
    Dim errReturnCode       As Long
    Dim response            As Integer

    If IsRegEditOpen(cReg) Then
        If promptUserBeforeClosing Then
            Select Case cReg.Count
            Case Is = 1: response = VBA.MsgBox("Are you sure that you want to close the Regestry Editor (regedit.exe)?", vbYesNo)
            Case Is > 1: response = VBA.MsgBox("Are you sure that you want to close [" & cReg.Count & "] instances of the Registry Editor (regedit.exe)?", vbYesNo)
            End Select
        End If
        If (response = vbYes) Or (promptUserBeforeClosing = False) Then
            For Each proc In cReg
                errReturnCode = proc.Terminate()
            Next proc
        End If
    End If

End Sub

'''
''' Tests whether or not the Registry Editor (regedit.exe) is open.
''' Option to pass all open instances back as a collection [cReg].
'''
Public Function IsRegEditOpen(Optional cReg As Collection) As Boolean

    Dim oServ       As Object
    Dim cProc       As Variant
    Dim oProc       As Object

    Set oServ = GetObject("winmgmts:")
    Set cProc = oServ.execquery("Select * from Win32_Process")
    Set cReg = New Collection

    For Each oProc In cProc
        If (oProc.Name = "regedit.exe") Then
            cReg.Add oProc
        End If
    Next

    On Error GoTo ErrHandler
    If (cReg.Count > 0) Then
        IsRegEditOpen = True
    End If

ErrHandler:
End Function

'''
''' Returns the name of the class.
'''
Public Property Get Name() As String

    Name = C_NAME

End Property

'''
''' Opens the Registry Editor (regedit.exe).
''' Opens to the previous key if [openToLastKey]:=True (which is the default behavior of RegEdit).
''' To kill existing regedit.exe processes before opening, set [closeBeforeOpening]:= TRUE.
'''
Public Sub OpenRegEdit(Optional openToLastKey As Boolean = True, Optional closeBeforeOpening = False)

    Dim wsh         As Object

    '// Set last position of reg.exe
    If Not openToLastKey Then
        RegistryUpdateValue HKEY.HKEY_CURRENT_USER_HKCU, "Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "LastKey", "Computer", createKeyIfNotExist:=False
    End If

    '// Option switch.
    If closeBeforeOpening Then
        CloseRegEdit promptUserBeforeClosing:=False
    End If

    '// Open instance of Registry Editor.
    On Error Resume Next
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run "regedit.exe -m", 1, False
    Set wsh = Nothing
    On Error GoTo 0

End Sub




'''
''' Opens the Registry Editor (regedit.exe) to specified locaiton.
''' To kill existing regedit.exe processes before opening, set [closeInstancesBeforeOpening]:= TRUE.
'''
Public Sub OpenRegEditToKey(baseKey As HKEY, ByVal Keyname As String, Optional closeBeforeOpening As Boolean = False)

    Dim wsh         As Object

    ResetErrorVariables

    '// Input checks.
    If IsValidBaseKey(baseKey:=baseKey) = False Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        GoTo ErrHandler
    End If
    If IsValidKeyName(Keyname:=Keyname) = False Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        GoTo ErrHandler
    End If
    If RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname) = False Then
        m_AppErr.Number = C_ERR_KEY_NOT_FOUND
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        GoTo ErrHandler
    End If

    '// Define open location.
    If Not RegistryUpdateValue(HKEY.HKEY_CURRENT_USER_HKCU, "Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", _
        "LastKey", "Computer\" & GetBaseKeyName(baseKey) & "\" & Keyname, createKeyIfNotExist:=False) Then
        m_AppErr.Number = C_ERR_INVALID_PATH
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        GoTo ErrHandler
    End If

    '// Option switch.
    If closeBeforeOpening Then
        CloseRegEdit promptUserBeforeClosing:=False
    End If

    '// Open instance of Registry Editor.
    On Error Resume Next
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run "regedit.exe -m", 1, False
    Set wsh = Nothing
    On Error GoTo 0
    Exit Sub

ErrHandler:
    m_AppErr.Source = "OpenRegEditToKey(...)"
    OpenRegEdit openToLastKey:=False
    m_AppErr.DisplayMessage

End Sub

'''
''' Returns the value of the value named *valueName* in *keyName* in *baseKey*.
''' Returns NULL if an error occurred.
'''
Public Function RegistryGetValue(baseKey As HKEY, ByVal Keyname As String, valueName As String) As Variant

    Dim hiveKey                 As Long
    Dim Result                  As Long
    Dim regDataType             As REG_DATA_TYPE
    Dim lenData                 As Long
    Dim longData                As Long
    Dim stringData              As String
    Dim intArr(0 To 1024)       As Integer
    Dim lenStringData           As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValue(...) As Variant"
        RegistryGetValue = Null
        Exit Function
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValue(...) As Variant"
        RegistryGetValue = Null
        Exit Function
    End If

    If (RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_KEY_NOT_FOUND
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValue(...) As Variant"
        RegistryGetValue = Null
        Exit Function
    End If

    regDataType = RegistryGetValueType(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName)
    hiveKey = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
    If (hiveKey = 0) Then
        '--m_RegSysErrNumber = result
        '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=result)
        m_AppErr.NumberDLL = Result
        m_AppErr.Number = C_ERR_UNABLE_TO_OPEN_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValue(...) As Variant"
        RegistryGetValue = Null
        Exit Function
    End If

    If (regDataType = REG_DWORD) Or (regDataType = REG_BINARY) Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Data is Long data-type.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Result = RegQueryValueEx(hiveKey:=hiveKey, lpValueName:=valueName, lpReserved:=0&, _
        lpType:=regDataType, lpData:=longData, lpcbData:=Len(longData))
        If (Result = ERROR_SUCCESS) Then
            RegistryGetValue = longData
            Exit Function
        Else
            '--m_RegSysErrNumber = result
            '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=result)
            m_AppErr.NumberDLL = Result
            m_AppErr.Number = C_ERR_UNABLE_TO_READ_VALUE
            m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
            m_AppErr.Source = "RegistryGetValue(...) As Variant"
            RegCloseKey hiveKey
            RegistryGetValue = Null
            Exit Function
        End If
    ElseIf (regDataType = REG_SZ) Or (regDataType = REG_EXPAND_SZ) Or (regDataType = REG_MULTI_SZ) Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Data is String data-type.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        stringData = VBA.String$(MAX_DATA_BUFFER_SIZE, vbNullChar)
        lenStringData = VBA.Len(stringData)
        Result = RegQueryValueExStr(hiveKey:=hiveKey, lpValueName:=valueName, lpReserved:=0&, _
        lpType:=regDataType, szData:=stringData, lpcbData:=lenStringData)
        If (Result <> ERROR_SUCCESS) Then
            '--m_RegSysErrNumber = result
            '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=result)
            m_AppErr.NumberDLL = Result
            m_AppErr.Number = C_ERR_UNABLE_TO_READ_VALUE
            m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
            m_AppErr.Source = "RegistryGetValue(...) As Variant"
            RegCloseKey hiveKey
            RegistryGetValue = Null
            Exit Function
        End If
        stringData = TrimToNull(stringData)
        RegistryGetValue = stringData
    Else
        m_AppErr.Number = C_ERR_INVALID_DATA_TYPE
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValue(...) As Variant"
        RegistryGetValue = Null
    End If

End Function

'''
''' Indicates whether or not the *keyName* exists in *baseKey*.
''' Returns TRUE or FALSE.
'''
''' Returns FALSE if an error occurred. See the global error values for more information.
''' If [createKeyIfNotExist]:=TRUE and the key does not exist, it will be created.
'''
Public Function RegistryKeyExists(baseKey As HKEY, ByVal Keyname As String, Optional createKeyIfNotExist As Boolean = False) As Boolean

    Dim hiveKey        As Long
    Dim Result         As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryKeyExists(...) As Boolean"
        RegistryKeyExists = False
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryKeyExists(...) As Boolean"
        RegistryKeyExists = False
    End If

    Result = RegOpenKey(hiveKey:=baseKey, lpSubKey:=Keyname, phkResult:=hiveKey)
    If (Result = ERROR_SUCCESS) Then
        RegistryKeyExists = True
    Else
        RegistryKeyExists = False
        If (createKeyIfNotExist = True) Then
            Result = RegistryCreateKey(baseKey:=baseKey, Keyname:=Keyname)
            RegistryKeyExists = CBool(Result)
        End If
    End If

    RegCloseKey hiveKey:=hiveKey

End Function

'''
''' Returns the number of sub-keys under *keyName* in *baseKey*.
'''
''' Returns -1 if an error occurred or if *keyName* does not exist.
''' Includes an option to return a list of the subkeys as [listOfSubKeyNames].
'''
Public Function RegistryNumberOfSubKeys(baseKey As HKEY, ByVal Keyname As String, Optional listOfSubKeyNames As Variant) As Long

    listOfSubKeyNames = RegistrySubKeyNamesToArray(baseKey, Keyname)

    If VBA.IsNull(listOfSubKeyNames) Then
        RegistryNumberOfSubKeys = -1
    Else
        RegistryNumberOfSubKeys = UBound(listOfSubKeyNames)
        If LBound(listOfSubKeyNames) = 0 Then RegistryNumberOfSubKeys = RegistryNumberOfSubKeys + 1
    End If

End Function

'''
''' Returns the number of values under *keyName* in *baseKey*.
'''
''' Returns -1 if an error occurred or if *keyName* does not exist.
''' Inludes the option to return a list of values as [listOfValueNames].
'''
Public Function RegistryNumberOfValues(baseKey As HKEY, ByVal Keyname As String, Optional listOfValueNames As Variant) As Long

    listOfValueNames = RegistryValueNamesToArray(baseKey, Keyname)

    If VBA.IsNull(listOfValueNames) Then
        RegistryNumberOfValues = -1
    Else
        RegistryNumberOfValues = UBound(listOfValueNames)
        If LBound(listOfValueNames) = 0 Then RegistryNumberOfValues = RegistryNumberOfValues + 1
    End If

End Function

'''
''' Returns REG_DATA_TYPE as a string.
'''
''' Required for WShRegKeyWrite.
'''
Public Function GetDataTypeName(dataType As REG_DATA_TYPE) As String

    Select Case dataType
    Case REG_INVALID:       GetDataTypeName = "REG_INVALID"
    Case REG_SZ:            GetDataTypeName = "REG_SZ"
    Case REG_EXPAND_SZ:     GetDataTypeName = "REG_EXPAND_SZ"
    Case REG_BINARY:        GetDataTypeName = "REG_BINARY"
    Case REG_DWORD:         GetDataTypeName = "REG_DWORD"
    Case REG_MULTI_SZ:      GetDataTypeName = "REG_MULTI_SZ"
    End Select

End Function

'''
''' Indicates whether or not *valueName* exists under *keyName* in *baseKey* hive.
''' Returns TRUE or FALSE.
'''
Public Function RegistryValueExists(baseKey As HKEY, ByVal Keyname As String, valueName As String, _
    Optional createKeyIfNotExist As Boolean = False, Optional CreateType As REG_DATA_TYPE = REG_DWORD) As Boolean

    Dim hiveKey        As Long
    Dim Result         As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryValueExists(...) As Boolean"
        RegistryValueExists = False
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryValueExists(...) As Boolean"
        RegistryValueExists = False
    End If

    hiveKey = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
    If (hiveKey = 0) Then
        m_AppErr.Number = C_ERR_UNABLE_TO_OPEN_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryValueExists(...) As Boolean"
        RegistryValueExists = False
    End If

    Result = RegQueryValueEx(hiveKey:=hiveKey, lpValueName:=valueName, lpReserved:=0&, lpType:=0&, lpData:=0&, lpcbData:=0&)
    If (Result = ERROR_SUCCESS) Or (Result = ERROR_MORE_DATA) Then
        RegistryValueExists = True
    Else
        If (createKeyIfNotExist = True) Then
            If (CreateType = REG_DWORD) Then
                Result = RegistryCreateValue(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName, _
                valueValue:=0&, createKeyIfNotExist:=True)
            Else
                Result = RegistryCreateValue(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName, _
                valueValue:=vbNullString, createKeyIfNotExist:=True)
            End If
            If (CBool(Result) = True) Then
                RegistryValueExists = True
            Else
                RegistryValueExists = False
            End If
        End If
    End If

    RegCloseKey hiveKey

End Function

'''
''' Returns an array of the names of all the subkeys under *keyName* in *baseKey*.
'''
''' Returns NULL if an error occurred or if *keyName* does not exist.
''' Indexes to base [0], regardless of Option Base switch.
'''
Public Function RegistrySubKeyNamesToArray(baseKey As HKEY, ByVal Keyname As String) As Variant

    Dim procHiveKeyRes            As Long
    Dim Result             As Long
    Dim ooReg           As Object

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistrySubKeyNames(...) As Variant"
        RegistrySubKeyNamesToArray = Null
        Exit Function
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistrySubKeyNames(...) As Variant"
        RegistrySubKeyNamesToArray = Null
        Exit Function
    End If

    procHiveKeyRes = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
    If (procHiveKeyRes = 0) Then
        m_AppErr.Number = C_ERR_UNABLE_TO_OPEN_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistrySubKeyNames(...) As Variant"
        RegistrySubKeyNamesToArray = Null
        RegCloseKey procHiveKeyRes
        Exit Function
    Else: RegCloseKey procHiveKeyRes
    End If

    '// Ref: http://msdn.microsoft.com/en-us/library/aa392722(v=vs.85).aspx
    On Error Resume Next
    Set ooReg = VBA.GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    ooReg.EnumKey baseKey, Keyname, RegistrySubKeyNamesToArray
    If (Err.Number <> 0) Then
        '--m_SysErr.Number = err.Number
        '--m_SysErr.Description = err.Description
        '--m_SysErr.Source = "RegistrySubKeyNames(...) As Variant"
        m_AppErr.NumberDLL = Err.LastDllError
        m_AppErr.Number = Err.Number
        m_AppErr.Description = Err.Description
        m_AppErr.Source = "RegistrySubKeyNames(...) As Variant"
    End If
    Set ooReg = Nothing
    On Error GoTo 0

End Function

'''
''' Returns an array of the names of all the values under *keyName*.
'''
''' Returns NULL if an error occurred or if the *keyName* does not exist.
''' Indexes to base [0], regardless of Option Base switch.
'''
Public Function RegistryValueNamesToArray(baseKey As HKEY, ByVal Keyname As String) As Variant

    Dim procHiveKeyRes      As Long
    Dim Result              As Long
    Dim ooReg               As Object

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryValueNamesToArray(...) As Variant"
        RegistryValueNamesToArray = Null
        Exit Function
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryValueNamesToArray(...) As Variant"
        RegistryValueNamesToArray = Null
        Exit Function
    End If

    procHiveKeyRes = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
    If (procHiveKeyRes = 0) Then
        m_AppErr.Number = C_ERR_UNABLE_TO_OPEN_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryValueNamesToArray(...) As Variant"
        RegistryValueNamesToArray = Null
        RegCloseKey procHiveKeyRes
        Exit Function
    Else: RegCloseKey procHiveKeyRes
    End If

    '// Ref: http://msdn.microsoft.com/en-us/library/xka57xy4.aspx
    Set ooReg = VBA.GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    On Error Resume Next
    ooReg.EnumValues baseKey, Keyname, RegistryValueNamesToArray
    If (Err.Number <> 0) Then
        '--m_RegSysErrNum = err.Number
        '--m_RegSysErrText = err.Description
        m_AppErr.NumberDLL = Err.LastDllError
        m_AppErr.Number = Err.Number
        m_AppErr.Description = Err.Description
        m_AppErr.Source = "RegistryValueNamesToArray(...) As Variant"
    End If
    Set ooReg = Nothing
    On Error GoTo 0

End Function

'''
''' This returns the data type of value named *valueName* in *keyName* in *baseKey*.
'''
''' The procedures in this module support only longs and strings, so the result will be REG_SZ
''' for a string, REG_DWORD for a long or REG_INVALID for any other data type.
'''
Public Function RegistryGetValueType(baseKey As HKEY, ByVal Keyname As String, valueName As String) As REG_DATA_TYPE

    Dim Result              As Long
    Dim procHiveKeyRes      As Long
    Dim dataType            As REG_DATA_TYPE

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValueType(...) As REG_DATA_TYPE"
        RegistryGetValueType = False
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValueType(...) As REG_DATA_TYPE"
        RegistryGetValueType = False
    End If

    Result = RegOpenKey(hiveKey:=baseKey, lpSubKey:=Keyname, phkResult:=procHiveKeyRes)
    If (Result <> ERROR_SUCCESS) Then
        m_AppErr.Number = C_ERR_UNABLE_TO_OPEN_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValueType(...) As REG_DATA_TYPE"
        RegistryGetValueType = REG_INVALID
        Exit Function
    End If

    Result = RegQueryValueEx(hiveKey:=procHiveKeyRes, lpValueName:=valueName, lpReserved:=0&, lpType:=dataType, lpData:=0&, lpcbData:=0&)
    If (Result <> ERROR_SUCCESS) And (Result <> ERROR_MORE_DATA) Then
        m_AppErr.Number = C_ERR_UNABLE_TO_READ_VALUE
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryGetValueType(...) As REG_DATA_TYPE"
        RegistryGetValueType = REG_INVALID
        RegCloseKey procHiveKeyRes
        Exit Function
    End If

    Select Case dataType
    Case REG_SZ
        RegistryGetValueType = REG_SZ
    Case REG_EXPAND_SZ
        RegistryGetValueType = REG_EXPAND_SZ
    Case REG_BINARY
        RegistryGetValueType = REG_BINARY
    Case REG_DWORD
        RegistryGetValueType = REG_DWORD
    Case REG_MULTI_SZ
        RegistryGetValueType = REG_MULTI_SZ
    Case Else
        RegistryGetValueType = REG_INVALID
    End Select

    RegCloseKey procHiveKeyRes

End Function

'''
''' This creates a value named *valueName* in *keyName* in *baseKey* with a value of *valueValue*.
''' If the key named by *keyName* does not exist, and [createKeyIfNotExist]:=TRUE, the key will
''' be created. If the value already exists, its value is set to th enew value if they are
''' compatible data types.
'''
Public Function RegistryCreateValue(baseKey As HKEY, ByVal Keyname As String, valueName As String, valueValue As Variant, _
    Optional createKeyIfNotExist As Boolean = False) As Boolean

    Dim procHiveKeyRes      As Long
    Dim Result              As Long
    Dim dataType            As REG_DATA_TYPE
    Dim StringValue         As String
    Dim LongValue           As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryCreateValue = False
        Exit Function
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryCreateValue = False
        Exit Function
    End If

    '** Satola B 11-Nov. 2014
    If (RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname, _
        createKeyIfNotExist:=createKeyIfNotExist) = False) Then
        m_AppErr.Number = C_ERR_KEY_NOT_FOUND
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryCreateValue = False
        Exit Function
    End If

    If (IsCompatibleValueValue(var:=valueValue) = False) Then
        m_AppErr.Number = C_ERR_INVALID_DATA_TYPE
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryCreateValue = False
        Exit Function
    End If

    If (RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname, createKeyIfNotExist:=False) = False) Then
        If (createKeyIfNotExist = True) Then
            If (RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname, createKeyIfNotExist:=True) = False) Then
                m_AppErr.Number = C_ERR_UNABLE_TO_CREATE_KEY
                m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
                RegistryCreateValue = False
                Exit Function
            End If
        Else
            m_AppErr.Number = C_ERR_KEY_NOT_FOUND
            m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
            RegistryCreateValue = False
            Exit Function
        End If
    End If

    If (RegistryValueExists(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName) = True) Then
        dataType = RegistryGetValueType(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName)
        If (dataType = REG_SZ) Then
            If (VarType(valueValue) <> vbString) Then
                m_AppErr.Number = C_ERR_DATA_TYPE_MISMATCH
                m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
                RegistryCreateValue = False
                Exit Function
            Else
                '''''''''''''''''''''''''''''
                ' ValueValue is a string. OK.
                '''''''''''''''''''''''''''''
            End If
        Else
            '''''''''''''''''''''''''
            ' ValueValue is numeric
            '''''''''''''''''''''''''
        End If
    Else
        '''''''''''''''''''''''
        ' Value does not exist.
        ' Set the DataType.
        '''''''''''''''''''''''
        If (VarType(valueValue) = vbString) Then
            dataType = REG_SZ
        Else
            dataType = REG_DWORD
        End If
    End If

    If (dataType = REG_DWORD) Then
        LongValue = VBA.CLng(valueValue)
        procHiveKeyRes = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
        If (procHiveKeyRes = 0) Then
            '--m_RegSysErrNumber = err.LastDllError
            '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=m_RegSysErrNum)
            m_AppErr.NumberDLL = Err.LastDllError
            m_AppErr.Number = C_ERR_UNABLE_TO_OPEN_KEY
            m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
            RegCloseKey procHiveKeyRes
            RegistryCreateValue = False
            Exit Function
        End If

        Result = RegSetValueExLong(hiveKey:=procHiveKeyRes, lpValueName:=valueName, Reserved:=0&, _
        dwType:=REG_DWORD, szData:=LongValue, cbData:=Len(LongValue))
        If (Result <> ERROR_SUCCESS) Then
            m_AppErr.Number = C_ERR_UNABLE_TO_UPDATE_VALUE
            m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
            RegCloseKey procHiveKeyRes
            RegistryCreateValue = False
            Exit Function
        End If
    Else
        StringValue = CStr(valueValue)
        procHiveKeyRes = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
        If (procHiveKeyRes = 0) Then
            '--m_RegSysErrNumber = err.LastDllError
            '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=m_RegSysErrNum)
            m_AppErr.NumberDLL = Err.LastDllError
            m_AppErr.Number = C_ERR_UNABLE_TO_OPEN_KEY
            m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
            RegCloseKey procHiveKeyRes
            RegistryCreateValue = False
            Exit Function
        End If
        Result = RegSetValueExStr(hiveKey:=procHiveKeyRes, lpValueName:=valueName, Reserved:=0&, _
        dwType:=REG_SZ, szData:=StringValue, cbData:=Len(StringValue))
        If (Result <> ERROR_SUCCESS) Then
            m_AppErr.Number = C_ERR_UNABLE_TO_UPDATE_VALUE
            m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
            RegistryCreateValue = False
            RegCloseKey procHiveKeyRes
            Exit Function
        End If
    End If

    RegCloseKey procHiveKeyRes
    RegistryCreateValue = True

End Function

'''
''' Creates a key named *keyName* in *baseKey*.
''' Returns True if successful or False if an error occurrs.
'''
Public Function RegistryCreateKey(baseKey As HKEY, ByVal Keyname As String) As Boolean

    Dim Result              As Long
    Dim procHiveKeyRes      As Long
    Dim dataType            As REG_DATA_TYPE
    Dim secAttrib           As SECURITY_ATTRIBUTES
    Dim disposition         As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryCreateKey = False
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryCreateKey = False
    End If

    If (RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname) = True) Then
        '''''''''''''''''''''''''''
        ' Key already exist. Return
        ' True as if we created it.
        '''''''''''''''''''''''''''
        RegistryCreateKey = True
        Exit Function
    End If

    Result = RegCreateKeyEx(hiveKey:=baseKey, lpSubKey:=Keyname, Reserved:=0&, lpClass:=vbNullString, _
    dwOptions:=REG_OPTION_NON_VOLATILE, samDesired:=KEY_ALL_ACCESS, _
    lpSecurityAttributes:=secAttrib, phkResult:=procHiveKeyRes, lpdwDisposition:=disposition)
    If (Result <> ERROR_SUCCESS) Then
        '--m_RegSysErrNumber = result
        '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=m_RegSysErrNum)
        m_AppErr.NumberDLL = Result
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryCreateKey = False
        Exit Function
    End If

    RegistryCreateKey = True

End Function

'''
''' This deletes a value named *valueName* in *keyName* in *baseKey*.
''' Returns TRUE or FALSE indicating success (returns TRUE if the value does not exist).
'''
Public Function RegistryDeleteValue(baseKey As HKEY, ByVal Keyname As String, valueName As String) As Boolean

    Dim Result              As Long
    Dim procHiveKeyRes      As Long
    Dim dataType            As REG_DATA_TYPE
    Dim secAttrib           As SECURITY_ATTRIBUTES
    Dim disposition         As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryDeleteValue = False
        Exit Function
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryDeleteValue = False
        Exit Function
    End If

    If (RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname, createKeyIfNotExist:=False) = False) Then
        m_AppErr.Number = C_ERR_KEY_NOT_FOUND
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryDeleteValue = False
        Exit Function
    End If

    procHiveKeyRes = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
    If (procHiveKeyRes = 0) Then
        RegistryDeleteValue = False
        Exit Function
    End If
    If RegistryValueExists(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName) = False Then
        RegCloseKey procHiveKeyRes
        RegistryDeleteValue = True
        Exit Function
    End If

    Result = RegDeleteValue(hiveKey:=procHiveKeyRes, lpValueName:=valueName)
    If (Result <> ERROR_SUCCESS) Then
        '--m_RegSysErrNumber = result
        '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=result)
        m_AppErr.Number = C_ERR_UNABLE_TO_DELETE_VALUE
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegCloseKey procHiveKeyRes
        RegistryDeleteValue = False
        Exit Function
    End If

    RegCloseKey procHiveKeyRes
    RegistryDeleteValue = True

End Function

'''
''' This deletes the registry key named *keyName* in *baseKey*, including all subkeys and values.
''' Returns TRUE or FALSE indicating success (returns TRUE if the key does not exist).
'''
Public Function RegistryDeleteKey(baseKey As HKEY, ByVal Keyname As String) As Boolean

    Dim Result              As Long
    Dim procHiveKeyRes      As Long
    Dim dataType            As REG_DATA_TYPE
    Dim secAttrib           As SECURITY_ATTRIBUTES
    Dim disposition         As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryDeleteKey = False
        Exit Function
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryDeleteKey = False
        Exit Function
    End If

    If (RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname, createKeyIfNotExist:=False) = False) Then
        RegistryDeleteKey = True
        Exit Function
    End If

    procHiveKeyRes = OpenRegistryKey(baseKey:=baseKey, Keyname:=Keyname)
    If (procHiveKeyRes = 0) Then
        RegistryDeleteKey = False
        Exit Function
    End If

    Result = RegDeleteKey(hiveKey:=baseKey, lpSubKey:=Keyname)
    RegCloseKey procHiveKeyRes
    If (Result <> ERROR_SUCCESS) Then
        '--m_RegSysErrNumber = result
        '--m_RegSysErrText = GetSystemErrorMessageText(result)
        m_AppErr.NumberDLL = Result
        m_AppErr.Number = C_ERR_UNABLE_TO_DELETE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        RegistryDeleteKey = False
        Exit Function
    End If

    RegistryDeleteKey = True

End Function

'''
''' Updates the value of the key, and returns TRUE or FALSE indicating success.
'''
Public Function RegistryUpdateValue(baseKey As HKEY, ByVal Keyname As String, valueName As String, NewValue As Variant, _
    Optional createKeyIfNotExist As Boolean = True) As Boolean

    Dim Result         As Boolean
    Dim hiveKey        As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey:=baseKey) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryUpdateValue(...) As Variant"
        RegistryUpdateValue = False
        Exit Function
    End If

    If (IsValidKeyName(Keyname:=Keyname) = False) Then
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryUpdateValue(...) As Variant"
        RegistryUpdateValue = False
        Exit Function
    End If

    If (IsCompatibleValueValue(var:=NewValue) = False) Then
        m_AppErr.Number = C_ERR_INVALID_DATA_TYPE
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryUpdateValue(...) As Variant"
        RegistryUpdateValue = False
        Exit Function
    End If

    Result = RegistryKeyExists(baseKey:=baseKey, Keyname:=Keyname, createKeyIfNotExist:=True)
    If (Result = False) Then
        m_AppErr.Number = C_ERR_KEY_NOT_FOUND
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryUpdateValue(...) As Variant"
        RegistryUpdateValue = False
        Exit Function
    End If

    If (VarType(NewValue) = vbString) Then
        Result = RegistryValueExists(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName, _
        createKeyIfNotExist:=createKeyIfNotExist, CreateType:=REG_DWORD)
    Else
        Result = RegistryValueExists(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName, _
        createKeyIfNotExist:=createKeyIfNotExist, CreateType:=REG_SZ)
    End If

    If (Result = False) Then
        m_AppErr.Number = C_ERR_VALUE_NOT_FOUND
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        m_AppErr.Source = "RegistryUpdateValue(...) As Variant"
        RegistryUpdateValue = False
        Exit Function
    End If

    Result = RegistryDeleteValue(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName)
    Result = RegistryCreateValue(baseKey:=baseKey, Keyname:=Keyname, valueName:=valueName, valueValue:=NewValue, createKeyIfNotExist:=True)

    RegistryUpdateValue = Result

End Function

'''
''' Opens the *keyName* in *baseKey* and returns the key handle if successful or 0 if an error occurs.
'''
Private Function OpenRegistryKey(baseKey As HKEY, ByVal Keyname As String) As Long

    Dim Result              As Long
    Dim procHiveKeyRes      As Long

    ResetErrorVariables

    If (IsValidBaseKey(baseKey) = False) Then
        ''''''''''''''''''''''''''''''''''''''
        ' Invalid Base Key. Return 0 and
        ' get out.
        ''''''''''''''''''''''''''''''''''''''
        OpenRegistryKey = 0
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        Exit Function
    End If

    Result = RegOpenKeyEx(hiveKey:=baseKey, lpSubKey:=Keyname, ulOptions:=0&, samDesired:=KEY_ALL_ACCESS, phkResult:=procHiveKeyRes)
    If (Result <> ERROR_SUCCESS) Then
        OpenRegistryKey = 0
        '--m_RegSysErrNumber = result
        '--m_RegSysErrText = GetSystemErrorMessageText(errNumber:=result)
        m_AppErr.NumberDLL = Result
        m_AppErr.Number = C_ERR_INVALID_BASE_KEY
        m_AppErr.Description = GetAppErrDescription(m_AppErr.Number)
        Exit Function
    End If

    OpenRegistryKey = procHiveKeyRes

End Function

'''
''' If [reverse] is omitted or FALSE, the function returns the portion of *text* that is to the
''' left of the first vbNullChar character. The vbNullChar is not returned. If *reverse* is TRUE,
''' the function returns the portion to the left of the last vbNullChar. The vbNullChar is not
''' returned. In either case, if vbNullChar is not found, the entire string Text is returned.
'''
Private Function TrimToNull(TEXT As String, Optional Reverse As Boolean = False) As String

    Dim pos         As Long

    If (Reverse = False) Then
        pos = VBA.InStr(1, TEXT, vbNullChar, vbTextCompare)
    Else
        pos = VBA.InStrRev(TEXT, vbNullChar, -1, vbTextCompare)
    End If
    If pos Then
        TrimToNull = VBA.Left(TEXT, pos - 1)
    Else
        TrimToNull = TEXT
    End If

End Function

'''
''' If *reverse* is FALSE, the function returns the portion of *text* that is to the left of the
''' first occurrence of *char*. If *reverse* is TRUE, the function returns the portion of *text*
''' that is to the left of the first occurrence of *char*. If *char* is not found, the entire
''' string *text* is returned.
'''
''' If [compaRemode} is vbBinaryCompare, *text* is compared in a CASE-SENSITIVE manner ("A" <> "a").
''' If [compaRemode] is any other value, text is compared in CASE-INSENSITVE mode ("A" = "a").
'''
Private Function TrimToChar(TEXT As String, Char As String, Optional ByVal Reverse As Boolean = False, _
    Optional ByVal CompaRemode As VbCompareMethod) As String

    Dim pos         As Long

    If (CompaRemode <> vbBinaryCompare) Then
        CompaRemode = vbTextCompare
    End If

    If (Reverse = False) Then
        pos = InStr(1, TEXT, Char, CompaRemode)
    Else
        pos = InStrRev(TEXT, Char, -1, CompaRemode)
    End If

    If pos Then
        TrimToChar = VBA.Left(TEXT, pos - 1)
    Else
        TrimToChar = TEXT
    End If

End Function

'''
''' Indicates whether or not *baseKey* is a valid base key (e.g. HKEY_CURRENT_USER).
'''
Private Function IsValidBaseKey(baseKey As HKEY) As Boolean

    Select Case baseKey
    Case HKEY.HKEY_CURRENT_USER_HKCU, HKEY.HKEY_LOCAL_MACHINE_HKLM, _
        HKEY.HKEY_CLASSES_ROOT_HKCR, HKEY.HKEY_CURRENT_CONFIG_HKCC, HKEY.HKEY_DYN_DATA_HKDD, _
        HKEY.HKEY_PERFORMANCE_DATA_HKPD, HKEY.HKEY_USERS_HKU
        IsValidBaseKey = True
    Case Else
        IsValidBaseKey = False
    End Select

End Function

'''
''' Resets the class error property holders.
'''
Private Sub ResetErrorVariables()

    m_AppErr.clear

End Sub

'''
''' Returns the text description of the application *errNumber*.
'''
Private Function GetAppErrDescription(errNumber As Long) As String

    Select Case errNumber
    Case C_ERR_NO_ERROR:                    GetAppErrDescription = vbNullString
    Case C_ERR_INVALID_BASE_KEY:            GetAppErrDescription = "Invalid Base Key Value."
    Case C_ERR_INVALID_DATA_TYPE:           GetAppErrDescription = "Invalid Data Type."
    Case C_ERR_KEY_NOT_FOUND:               GetAppErrDescription = "Key Not Found."
    Case C_ERR_VALUE_NOT_FOUND:             GetAppErrDescription = "Value Not Found."
    Case C_ERR_DATA_TYPE_MISMATCH:          GetAppErrDescription = "Value Data Type Mismatch."
    Case C_ERR_ENTRY_LOCKED:                GetAppErrDescription = "Registry Entry Locked."
    Case C_ERR_INVALID_KEYNAME:             GetAppErrDescription = "The Specified Key Is Invalid."
    Case C_ERR_UNABLE_TO_OPEN_KEY:          GetAppErrDescription = "Unable To Open Key."
    Case C_ERR_UNABLE_TO_READ_KEY:          GetAppErrDescription = "Unable To Read Key."
    Case C_ERR_UNABLE_TO_CREATE_KEY:        GetAppErrDescription = "Unable To Create Key."
    Case C_ERR_UNABLE_TO_READ_VALUE:        GetAppErrDescription = "Unable To Read Value."
    Case C_ERR_UNABLE_TO_UPDATE_VALUE:      GetAppErrDescription = "Unable To Update Value."
    Case C_ERR_UNABLE_TO_CREATE_VALUE:      GetAppErrDescription = "Unable To Create Value."
    Case C_ERR_UNABLE_TO_DELETE_KEY:        GetAppErrDescription = "Unable To Delete Key."
    Case C_ERR_UNABLE_TO_DELETE_VALUE:      GetAppErrDescription = "Unable To Delete Value."
    Case C_ERR_INVALID_PATH:                GetAppErrDescription = "Invalid registry path."
    Case Else
        GetAppErrDescription = "Undefined Error."
    End Select

End Function

'''
''' Indicates whether the length of *txt* is less than REGSTR_MAX_VALUE_LENGTH.
'''
Private Function IsStringValidLength(txt As String) As Boolean

    IsStringValidLength = (Len(txt) <= REGSTR_MAX_VALUE_LENGTH)

End Function

'''
''' Indicates whether or not *keyName* is valid.
''' Returns TRUE or FALSE.
'''
''' An invalid key is one whose name length is greater than REGSTR_MAX_VALUE_LENGTH, is all
''' spaces or is an empty string i.e. MUST KEEP REFERENCE TO VARIABLE.
'''
Private Function IsValidKeyName(Keyname As String) As Boolean

    IsValidKeyName = (VBA.Len(Keyname) <= REGSTR_MAX_VALUE_LENGTH) And (Len(VBA.Trim(Keyname)) > 0)

    '/* Satola B 11-Nov. 2014, checks for leading and ending slashes, and
    '   Removes them from the 'keyName' defintion if found. This eliminates
    '   a typical formatting error when calling the API functions.
    If (VBA.Mid(Keyname, 1, 1) = "\") Then
        Do While VBA.Mid(Keyname, 1, 1) = "\"
            Keyname = VBA.Mid(Keyname, 2, VBA.Len(Keyname) - 1)
        Loop
    End If
    If (VBA.Mid(Keyname, VBA.Len(Keyname), 1) = "\") Then
        Do While VBA.Mid(Keyname, VBA.Len(Keyname), 1) = "\"
            Keyname = VBA.Mid(Keyname, 1, VBA.Len(Keyname) - 1)
        Loop
    End If

End Function

'''
''' Indicates whether or not *dataType* is avalid data type (e.g. REG_DWORD).
''' Returns TRUE or FALSE.
'''
Private Function IsValidDataType(dataType As REG_DATA_TYPE) As Boolean

    Select Case dataType
    Case REG_SZ, REG_DWORD
        IsValidDataType = True
    Case Else
        IsValidDataType = False
    End Select

End Function

'''
''' Indicates if *var* can be used as a registry key value.
''' Returns TRUE or FALSE.
'''
''' Note that all numeric data types (singles, doubles, etc.) are considered values, even
''' though their values will be changed when converted to longs.
'''
Private Function IsCompatibleValueValue(var As Variant) As Boolean

    If VarType(var) >= vbArray Then
        IsCompatibleValueValue = False
        Exit Function
    End If
    If IsArray(var) = True Then
        IsCompatibleValueValue = False
        Exit Function
    End If
    If IsObject(var) = True Then
        IsCompatibleValueValue = False
        Exit Function
    End If

    Select Case VarType(var)
    Case vbBoolean, vbByte, vbCurrency, vbDate, vbDouble, vbInteger, vbLong, vbSingle, vbString
        IsCompatibleValueValue = True
    Case Else
        IsCompatibleValueValue = False
    End Select

End Function

'''
''' Returns version information for the class.
'''
Public Property Get version() As String

    '==========================================================================================================================
    ' Class changelog here.
    '==========================================================================================================================

    'Version = "Version 1.0 (2014)"  'Initial (official) release.
    'Version = "Version 1.1 (2014)"  'Satola B 10-Oct. 2014
    '   |* Minor corrections e.g. cleaned-up variable declarations and descriptions.
    '   |* The function 'GetDataTypeName' is also now accessible, so that it can be used with the
    '       Windows scripting shell functions (wsh). Using the standard API functions, however,
    '       are recommended. I only added the wsh functions because I was having trouble modifying
    '       a security key using the API functions, but the error was because I wrongly formatted the
    '       keyName string. I made the shell functions private, incase they are needed in the future.
    '   |* Because of [this] I modified the 'IsValidKeyName' function to automatically Remove leading
    '       and trailing "\" characters from KeyName input parameters (i.e. to Remove source of errors).
    'Version = "Version 2.0 (2015)"  'Satola B 08-Nov. 2015
    '   |* Additional updates to descriptions for consistency. I needed to change to a specific format
    '       so that I can programmatically export the class summary to worksheet i.e. easier maintenance.
    '   |* Changed application error variables and reformatted "GetAppErrDescription" for consistency
    '       with other classes. This includes adding "AppErrSource" property and "AppErrDisplayMessage"
    '       method, and changing the name of "GetAppErrText" to "GetAppErrDescription". [m_appErrSource]
    '       definitions were added to methods as necessary.
    '   |* Changed "GetDataTypeName" to use REG_DATA_TYPE enumerations instead of literals.
    version = "3.0 (2017)"
    '   |* Improved documentation and readability.
    '   |* Constructor statement now uses *Name* property to print to immediate.
    '   |* Completely Removed shell functions; duplicate functionality that causes confusion.
    '   |* Added *OpenRegEditToKey* so that regedit.exe can be opened to a specific location.
    '   |* Added the function *BaseKeyName* in support of method *OpenRegEditToKey*.
    '   |* Added the function *BaseKeyNameShort* for WScript.Shell, needed for DemoRegistryEditor module.
    '   |* Reformated comments to new convention (for later automation).
    '   |* Replaced application error properties and methods with a new error object/holder.
    '   |* Added opption to return the class name without extension (.cls).
    '   |* Removed the *Version* string from the get version property (more intuitive this way).
    '   |* Removed *HKEY_* properties in favour of a public enumeration: ~HKEY~.

End Property

'''
''' Class constructor.
'''
Private Sub Class_Initialize()

    Debug.Print "|* Initializing Class:= " & C_NAME

    Set m_AppErr = New ApplicationError
    m_AppErr.Initialize C_NAME

End Sub

