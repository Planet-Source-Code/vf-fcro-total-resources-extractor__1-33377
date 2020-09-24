VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FORMATPIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Declare Function CreateStreamOnHGlobal Lib "ole32" _
                              (ByVal hGlobal As Long, _
                              ByVal fDeleteOnRelease As Long, _
                              ppstm As Any) As Long

Declare Function OleLoadPicture Lib "olepro32" _
                              (pStream As Any, _
                              ByVal lSize As Long, _
                              ByVal fRunmode As Long, _
                              riid As GUID, _
                              ppvObj As Any) As Long

Public Type GUID
  dwData1 As Long
  wData2 As Integer
  wData3 As Integer
  abData4(7) As Byte
End Type
Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long
Const sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Const GMEM_MOVEABLE = &H2
Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByValdwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Function GetPicture(dataXY() As Byte) As IPicture
Dim hMem  As Long
Dim lpMem  As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown
Dim ipic As IPicture
hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(dataXY) + 1)
lpMem = GlobalLock(hMem)
CopyMemory ByVal lpMem, dataXY(0), UBound(dataXY) + 1
Call GlobalUnlock(hMem)
Call CreateStreamOnHGlobal(hMem, 1, istm)
Call CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture)
Call OleLoadPicture(ByVal ObjPtr(istm), UBound(dataXY) + 1, 0, IID_IPicture, GetPicture)
Call GlobalFree(hMem)
End Function








