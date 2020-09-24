Attribute VB_Name = "Module6"

'Test za kasnije---UPDATE
 
'Dim XB() As Byte
'Dim jk As Long
'Dim HDLFL As Long
'HDLFL = BeginUpdateResource(App.Path & "\1.exe", 1)
'File nesmije biti ucitan sa loadlibrary jer ce stvoriti TEMP file!
'Ako je drugi parametar 1(TRUE) tada se brišu svi resursi
'Ako je 0 onda ne
'Open App.Path & "\2.bin" For Binary As #1
'ReDim XB(LOF(1) - 1)
'Get #1, , XB
'Close #1

'Dim res1 As Long
'res1 = UpdateResource(HDLFL, 2&, 105&, 0, ByVal 0&, 0) ****Brisanje resource-a
'res1 = UpdateResource(HDLFL, 2&, 105&, 0, ByVal VarPtr(XB(0)), CLng(UBound(XB) + 1))

'res1 = UpdateResource(HDLFL, TypePtr, TrueName, LangID, VarPtr(XB(0)), CLng(UBound(XB) + 1))
'res1 = UpdateResource(HDLFL, TypePtr, TrueName, LangID, vbNull, vbNull)
'Dim res2 As Long
'res2 = EndUpdateResource(HDLFL, 0)
