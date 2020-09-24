Attribute VB_Name = "ModuleReg"
Private Declare Function GetSystemDirectory Lib "kernel32" _
   Alias "GetSystemDirectoryA" _
  (ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long
   
Private Declare Function FreeLibraryRegister _
    Lib "kernel32" Alias "FreeLibrary" ( _
    ByVal hLibModule As Long) As Long

Private Declare Function GetProcAddressRegister _
    Lib "kernel32" Alias "GetProcAddress" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String) As Long

Private Declare Function CreateThreadForRegister _
    Lib "kernel32" Alias "CreateThread" ( _
    lpThreadAttributes As Any, _
    ByVal dwStackSize As Long, _
    ByVal lpStartAddress As Long, _
    ByVal lpparameter As Long, _
    ByVal dwCreationFlags As Long, _
    lpThreadID As Long) As Long

Private Declare Function GetExitCodeThread _
    Lib "kernel32" ( _
    ByVal hThread As Long, _
    lpExitCode As Long) As Long

Private Declare Sub ExitThread _
    Lib "kernel32" ( _
    ByVal xc As Long)

Public Declare Function CloseHandle _
    Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function WaitForSingleObject _
    Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long



Private Function RegX(fName$, func%) As Integer
    Dim regLib&, process&, succeed&
    Dim h1&, xc&, id&
    Dim p$
    
    Select Case func
        Case 0: p = "DllUnregisterServer"
        Case 1: p = "DllRegisterServer"
        Case Else: RegX = 0
                    Exit Function
    End Select

    regLib = LoadLibrary(fName)
    If regLib = 0 Then
        RegX = 1
        Exit Function
    End If
        
    process = GetProcAddressRegister(regLib, p)
    
    If process = 0 Then
        RegX = 2
    Else
        h1 = CreateThreadForRegister(ByVal 0&, 0&, _
            ByVal process, ByVal 0&, 0&, id)
        If h1 = 0 Then
            RegX = 3
        Else
            succeed = (WaitForSingleObject(h1, 10000) = 0)
            If succeed Then
                CloseHandle h1
                RegX = 4
            Else
                GetExitCodeThread h1, xc
                ExitThread xc
                RegX = 5
            End If
        End If
    End If

    FreeLibraryRegister regLib
End Function

Public Function GetSystemDir() As String
Dim nSize As Long
Dim tmp As String
tmp = Space$(256)
nSize = Len(tmp)
Call GetSystemDirectory(tmp, nSize)
GetSystemDir = TrimNull(tmp)
End Function
Private Function TrimNull(item As String)
Dim pos As Integer
pos = InStr(item, Chr$(0))
If pos Then
TrimNull = Left$(item, pos - 1)
Else: TrimNull = item
End If
End Function
Public Function RegGif() As Integer
On Error GoTo REGME
Dim TESTO As Object
Set TESTO = CreateObject("Gif89.Gif89.1")
Set TESTO = Nothing
RegGif = 4
Exit Function
REGME:
On Error GoTo 0
Dim Ddata() As Byte
Dim SYSP As String
SYSP = GetSystemDir & "\gif89.dll"
Ddata = LoadResData(101, "CUSTOM")
Open SYSP For Binary As #1
Put #1, , Ddata
Close #1
RegGif = RegX(SYSP, 1)
Erase Ddata
End Function
Sub main()
If RegGif <> 4 Then
MsgBox "Could not Create ActiveX!!!Sorry somethings goes wrong!", vbCritical, "Error"
End
End If
Form1.Show
End Sub
