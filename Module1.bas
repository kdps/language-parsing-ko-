Attribute VB_Name = "AI"
Public talk_memory As String '기억
Public talk_memory2 As String '기억(2) - 위치
Public talk_process() As String '행동
Public talk_split() As String '분할
Public talk_inv() As String '보조사
Public talk_do() As String '진행완료
Public talk_and() As String '접속부사
Public talk_ask() As String '질문
Public talk_is() As String '서술격조사
Public talk_to() As String '격조사
Public talk_in() As String '격조사2
Public mem_name As New Collection '이름
Public mem_cont As New Collection '내용
Public mem_profile As New Collection '속성
Public chk_and As Boolean '접속부사 체크
Public chk_inv As Boolean '접속부사 체크

Public Function Analys()

Dim talk_spt() As String
Dim i, z, q, w As Long
z = 1
chk_and = False
talk_memory = ""

For i = 0 To Len(frmmain.txtConvert.Text)

    '격조사 ~에서
    For q = 0 To UBound(talk_in) '검색
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_in(q))) = talk_in(q) Then '발견되면
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_in(q)) + 1), 1) = " " Then
                talk_memory2 = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                z = i + Len(talk_in(q)) + 2
            End If
        End If
    Next q
    
    '격조사 ~을,를
    For q = 0 To UBound(talk_to)
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_to(q))) = talk_to(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_to(q)) + 1), 1) = " " Then
                If chk_and = False Then
                    talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                Else
                    talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    chk_and = True
                End If
                z = i + Len(talk_to(q)) + 2
            End If
        End If
    Next q
    
    '접속부사 ~와,과
    For q = 0 To UBound(talk_and)
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_and(q))) = talk_and(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_and(q)) + 1), 1) = " " Then
                If chk_and = False Then
                    talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                Else
                    talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    chk_and = True
                End If
                z = i + Len(talk_and(q)) + 2
            End If
        End If
    Next q
    
    '서술격조사 ~은,는
    For q = 0 To UBound(talk_inv)
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_inv(q))) = talk_inv(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_inv(q)) + 1), 1) = " " Then
                talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                z = i + Len(talk_and(q)) + 2
                chk_and = True
                chk_inv = True
            End If
        End If
    Next q
    
    '보조사 ~이다
    For q = 0 To UBound(talk_is)
        If Mid(frmmain.txtConvert.Text, (i + 1), Len(talk_is(q)) + 1) = talk_is(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_is(q)) + 1), 1) = " " Or Mid(frmmain.txtConvert.Text, i + 2, 1) = "." Or i = Len(frmmain.txtConvert.Text) - Len(talk_is(q)) Then
                If chk_inv = True Then
                    If InStr(talk_memory, ",") Then
                        'talk_memory = talk_memory & "," & Mid(frmmain.txtconvert.Text, z, (i - z) + 1)
                        talk_split() = Split(talk_memory, ",")
                        For w = 1 To UBound(talk_split)
                            mem_name.Add talk_split(w)
                            mem_cont.Add Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                            frmmain.lstAnalys.AddItem talk_split(w) & "에 " & Mid(frmmain.txtConvert.Text, z, (i - z) + 1) & "를 저장했습니다."
                        Next w
                    Else
                        mem_name.Add Mid(talk_memory, 2, Len(talk_memory))
                        mem_cont.Add Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                        frmmain.lstAnalys.AddItem Mid(talk_memory, 2, Len(talk_memory)) & "에 " & Mid(frmmain.txtConvert.Text, z, (i - z) + 1) & "를 저장했습니다."
                    End If
                Else
                    frmmain.lstAnalys.AddItem Mid(talk_memory, 2, Len(talk_memory)) & "는 문법에 안맞습니다."
                End If
            End If
        End If
    Next q
    
    '저장 ~는?
    For q = 0 To UBound(talk_ask)
        If Mid(frmmain.txtConvert.Text, i + 1, 1) = talk_ask(q) Or Mid(frmmain.txtConvert.Text, i + 2, 1) = "?" Then
            If Mid(frmmain.txtConvert.Text, i + 2, 1) = " " Or i = Len(frmmain.txtConvert.Text) - Len(talk_ask(q)) Then
                If InStr(talk_memory, ",") Then
                    talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    talk_split() = Split(talk_memory, ",")
                    For w = 1 To UBound(talk_split)
                        frmmain.lstAnalys.AddItem talk_split(w) & "는 " & GetVar(Mid(frmmain.txtConvert.Text, z, (i - z) + 1))
                    Next w
                Else
                    talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    frmmain.lstAnalys.AddItem talk_memory & "는 " & GetVar(Mid(frmmain.txtConvert.Text, z, (i - z) + 1))
                End If
            End If
        End If
    Next q

    '행동 ~해줘
    For q = 0 To UBound(talk_process)
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_process(q))) = talk_process(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_process(q)) + 1), 1) = "?" Or Mid(frmmain.txtConvert.Text, i + (Len(talk_process(q)) + 1), 1) = " " Or i = Len(frmmain.txtConvert.Text) - Len(talk_process(q)) Then
                If InStr(talk_memory, ",") Then
                    talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    talk_split() = Split(talk_memory, ",")
                    For w = 1 To UBound(talk_split) - 1
                        If Not GetVar(talk_memory2) = "" Then
                            talk_memory2 = GetVar(talk_memory2)
                        End If
                        If Not GetVar(talk_split(w)) = "" Then
                            talk_split(w) = GetVar(talk_split(w))
                        End If
                        Search_Function (Mid(frmmain.txtConvert.Text, z, (i - z) + 1)), talk_memory2, talk_split(w)
                        frmmain.lstAnalys.AddItem talk_memory2 & "에서 " & talk_split(w) & "를 " & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    Next w
                Else
                    talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    frmmain.lstAnalys.AddItem GetVar(talk_memory2) & "에서1 " & talk_split(w) & "를 " & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                End If
                z = i + Len(talk_and(q)) + 2
            End If
        End If
    Next q
    
Next i

End Function

Private Function Search_Function(Functions As String, Address1 As String, Address2 As String) As String
If Functions = "검색" Then
    If InStr(Address1, "http://") Or InStr(Address1, "https://") Then
        ShellExecute 0, "open", Address1 & Address2, "", "", 0
    End If
End If
End Function

Private Function GetVar(TheVar As String) As Variant
Dim i As Integer
For i = 1 To mem_name.Count
    If mem_name(i) = TheVar Then
        GetVar = mem_cont(i)
        Exit Function
    End If
Next i
End Function

Public Function Basic_Function()

End Function

Public Function Set_Function()

ReDim talk_process(1)

talk_process(0) = "해줘"
talk_process(1) = "해"

ReDim talk_in(0)

talk_in(0) = "에서"

ReDim talk_to(1)

talk_to(0) = "을"
talk_to(1) = "를"

ReDim talk_is(0)

talk_is(0) = "이다"

ReDim talk_inv(1) '조사

talk_inv(0) = "은"
talk_inv(1) = "는"

ReDim talk_do(4) '진행완료

talk_do(0) = "do"
talk_do(1) = "했다"
talk_do(2) = "했는데"
talk_do(3) = "했음"
talk_do(4) = "했어"

ReDim talk_and(4)

talk_and(0) = ","
talk_and(1) = "또"
talk_and(2) = "그리고"
talk_and(3) = "와"
talk_and(4) = "과"

ReDim talk_ask(3) '질문

talk_ask(0) = "뭐지"
talk_ask(1) = "에 저장된 값은?"
talk_ask(2) = "뭐야"
talk_ask(3) = "는?"

End Function
