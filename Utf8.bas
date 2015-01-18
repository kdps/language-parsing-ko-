Attribute VB_Name = "Utf8"
Private Declare Function WideCharToMultiByteArray Lib "kernel32" Alias "WideCharToMultiByte" _
    (ByVal codepage As Long, _
     ByVal dwFlags As Long, _
     ByRef lpWideCharStr As Byte, _
     ByVal cchWideChar As Long, _
     ByRef lpMultiByteStr As Byte, _
     ByVal cchMultiByte As Long, _
     ByVal lpDefaultChar As Long, _
     ByVal lpUsedDefaultChar As Long) As Long
     
Private Const CP_UTF8 As Long = 65001

Public Function URLEncoder_UTF8(ByVal szString As String) As String
    Dim lngByteNum, i As Long
    Dim abytUTF16() As Byte
    Dim abytUTF8() As Byte
    Dim lngCharCount As Long
    
    abytUTF16 = szString
    lngCharCount = (UBound(abytUTF16) + 1) \ 2
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, 0, 0, 0, 0)
                    
    If lngByteNum > 0 Then
        ReDim abytUTF8(lngByteNum - 1)
        lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, _
                                         abytUTF8(0), lngByteNum, 0, 0)
    End If
    For i = 0 To UBound(abytUTF8): URLEncoder_UTF8 = URLEncoder_UTF8 & "%" & Right("0" & Hex(abytUTF8(i)), 2): Next
End Function

