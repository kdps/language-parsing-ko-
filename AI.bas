Attribute VB_Name = "AI"
Option Explicit

Public Function Analys()

On Error Resume Next
Dim talk_spt() As String
Dim i, z, q, w As Long

z = 1

'도움말
If frmmain.txtConvert.Text = "도움말" Then
    frmmain.Alert "", 1
    frmmain.Alert "#도움말#", 1
    frmmain.Alert "연주 : (코드)를 연주해줘", 1
    frmmain.Alert "주소 추과 : (이름)의 주소는 ~다", 1
    frmmain.Alert "주소 열기 : (주소의 이름)을 열어줘", 1
    frmmain.Alert "검색엔진 추과 : (이름)의 검색엔진은 ~다", 1
    frmmain.Alert "검색엔진 사용 : (검색엔진의 이름)에서 ~를 검색해줘(다중검색 사용가능)", 1
    frmmain.Alert "프로그램 추과 : (이름의 경로는 ~다", 1
    frmmain.Alert "프로그램 실행 : (프로그램의 이름)을 열어줘(다중실행 가능)", 1
    frmmain.Alert "", 1
    Exit Function
End If

For i = 0 To Len(frmmain.txtConvert.Text) '0 ~ 전체길이 해석

    '격조사 ~에서
    For q = 0 To UBound(talk_in) '검색
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_in(q))) = talk_in(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_in(q)) + 1), 1) = " " Then
                talk_memory2 = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                z = i + Len(talk_in(q)) + 2
            End If
        End If
    Next q

    '속성 ~의
    For q = 0 To UBound(talk_of) '검색
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_of(q))) = talk_of(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_of(q)) + 1), 1) = " " Then
                talk_memory3 = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                z = i + Len(talk_of(q)) + 2
            End If
        End If
    Next q

    '을 & 를
    If ObjJosa(Mid(frmmain.txtConvert.Text, i + 1, 1)) = "를" And ((Mid(frmmain.txtConvert.Text, i + 2, 1) = "를" Or Mid(frmmain.txtConvert.Text, i + 2, 1) = "가")) Then
        If Mid(frmmain.txtConvert.Text, i + 3, 1) = " " Then
            If chk_and = False Then
                talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 2)
            Else
                talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                chk_and = True
            End If
            z = i + 3
        End If
    End If
    
    If ObjJosa(Mid(frmmain.txtConvert.Text, i + 1, 1)) = "을" And ((Mid(frmmain.txtConvert.Text, i + 2, 1) = "을" Or Mid(frmmain.txtConvert.Text, i + 2, 1) = "이")) Then
        If Mid(frmmain.txtConvert.Text, i + 3, 1) = " " Then
            If chk_and = False Then
                talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 2)
            Else
                talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                chk_and = True
            End If
            z = i + 3
        End If
    End If
    
    If Mid(frmmain.txtConvert.Text, i + 1, 1) <> "" Then
        If (Mid(frmmain.txtConvert.Text, i + 1, 1) >= "a" And Mid(frmmain.txtConvert.Text, i + 1, 1)) <= "z" Or ((Mid(frmmain.txtConvert.Text, i + 1, 1) >= "A" And Mid(frmmain.txtConvert.Text, i + 1, 1) <= "Z")) Then
            If ObjJosa(Mid(frmmain.txtConvert.Text, i + 2, 1)) = "영" Then
                If Mid(frmmain.txtConvert.Text, i + 3, 1) = " " Then
                    If chk_and = False Then
                        talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 2)
                    Else
                        talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                        chk_and = True
                    End If
                    z = i + 3
                End If
            End If
        End If
    End If
    
    '기타
    If Mid(frmmain.txtConvert.Text, i + 1, 1) <> "" Then
        If (ObjJosa(Mid(frmmain.txtConvert.Text, i + 1, 1)) = "를" And Mid(frmmain.txtConvert.Text, i + 2, 1) = "을") Or (ObjJosa(Mid(frmmain.txtConvert.Text, i + 1, 1)) = "을" And Mid(frmmain.txtConvert.Text, i + 2, 1) = "를") Then
            If Mid(frmmain.txtConvert.Text, i + 3, 1) = " " Then
                frmmain.Alert ("문법이 안맞습니다 - " & Mid(frmmain.txtConvert.Text, i + 1, 1) & ObjJosa(Mid(frmmain.txtConvert.Text, i + 1, 1))), 2
                z = i + 2 + Len(Mid(frmmain.txtConvert.Text, z, (i - z) + 1))
            End If
        End If
    End If
    
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
        
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_is(q)) + 1), 1) = " " Or i = Len(frmmain.txtConvert.Text) - Len(talk_is(q)) Then
                If chk_inv = True Then
                    If InStr(talk_memory, ",") Then
                        'talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                        talk_split() = Split(talk_memory, ",")
                        talk_search = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                        If GetVar(talk_search) <> "" Then
                            talk_search = GetVar(talk_search)
                        End If
                        For w = 1 To UBound(talk_split)
                            If talk_memory3 <> "" Then
                                ChangeVar talk_memory3, talk_search, talk_split(w)
                                frmmain.Alert talk_memory3 & "의 " & talk_split(w) & ObjJosa(talk_split(w)) & " " & talk_search & "로 저장되었습니다.", 1
                            Else
                                ChangeVar talk_memory3, talk_search, ""
                                frmmain.Alert talk_split(w) & ObjJosa(talk_split(w)) & " 저장했습니다.", 1
                            End If
                        Next w
                        talk_memory3 = ""
                    Else
                        var_name.Add Mid(talk_memory, 2, Len(talk_memory))
                        var_cont.Add Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                        var_profile.Add ""
                        frmmain.Alert Mid(talk_memory, 2, Len(talk_memory)) & ObjJosa(Mid(talk_memory, 2, Len(talk_memory))) & " 저장했습니다.", 1
                    End If
                Else
                    frmmain.Alert Mid(talk_memory, 2, Len(talk_memory)) & ObjJosa(Mid(talk_memory, 2, Len(talk_memory))) & " 문법에 안맞습니다.", 2
                End If
            End If
        End If
    Next q
    
    '저장 ~는?
    For q = 0 To UBound(talk_ask)
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_ask(q)) + 1) = talk_ask(q) Then
            If Mid(frmmain.txtConvert.Text, i + 2, 1) = " " Or i = Len(frmmain.txtConvert.Text) - Len(talk_ask(q)) Then
                If InStr(talk_memory, ",") Then
                    talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    talk_split() = Split(talk_memory, ",")
                    For w = 1 To UBound(talk_split)
                        If talk_memory3 <> "" Then
                            frmmain.Alert talk_memory3 & "의 " & ObjJosa2(talk_split(w)) & " " & GetVar(talk_split(w)), 1
                        Else
                            frmmain.Alert talk_split(w) & ObjJosa2(talk_split(w)) & " " & GetVar(talk_split(w)), 1
                        End If
                        
                    Next w
                Else
                    talk_memory = Mid(frmmain.txtConvert.Text, z, (i - z) + 1)
                    If talk_memory3 <> "" Then
                        frmmain.Alert talk_memory3 & "의 " & talk_memory & ObjJosa2(talk_memory) & " " & GetProfile(talk_memory3, talk_memory), 1
                    Else
                        frmmain.Alert talk_memory & ObjJosa2(talk_memory) & " " & GetVar(talk_memory), 1
                    End If
                End If
            End If
        End If
    Next q

    '행동 ~해줘
    For q = 0 To UBound(talk_process)
        If Mid(frmmain.txtConvert.Text, i + 1, Len(talk_process(q))) = talk_process(q) Then
            If Mid(frmmain.txtConvert.Text, i + (Len(talk_process(q)) + 1), 1) = " " Or i = Len(frmmain.txtConvert.Text) - Len(talk_process(q)) Then
                If InStr(talk_memory, ",") Then
                    talk_memory = talk_memory & "," & Mid(frmmain.txtConvert.Text, z + 1, (i - z))
                    talk_split() = Split(talk_memory, ",")
                    For w = 0 To UBound(talk_split) - 1
                    
                        If GetProfile(talk_memory3, talk_memory2) <> "" Then
                            talk_result = GetProfile(talk_memory3, talk_memory2)
                        ElseIf talk_memory3 <> "" Then
                            talk_result = talk_memory3
                        ElseIf talk_memory2 <> "" Then
                            talk_result = talk_memory2
                        End If
                        
                        If talk_split(w) <> "" Then
                            Search_Function (Mid(frmmain.txtConvert.Text, z + 1, (i - z))), talk_result, talk_split(w), talk_memory3
                        End If
                    Next w
                End If
            End If
        End If
    Next q
    
Next i

End Function

Public Function Search_Function(Functions As String, Address1 As String, Address2 As String, Address3 As String) As String
On Error Resume Next
Dim NotePlays() As String
Dim PosNote As String
Dim si As SYSTEM_INFO
Dim i, w As Long
    
If Functions = "검색" Then

    If GetProfile(Address1, Address2) <> "" Then 'Search Profile(1,2)
        ShellExecute 0, "open", GetProfile(Address1, Address2) & GetAddon(Address3), "", "", 0
        frmmain.Alert GetProfile(Address1, Address2) & GetAddon(Address3) & ObjJosa(GetProfile(Address1, Address2)) & " 검색했습니다", 0
    ElseIf GetProfile(Address1, "검색엔진") <> "" Then 'Search Profile(1,검색엔진)
        ShellExecute 0, "open", GetProfile(Address1, "검색엔진") & URLEncoder_UTF8(Address2) & GetAddon(Address3), "", "", 0
        frmmain.Alert GetProfile(Address1, "검색엔진") & URLEncoder_UTF8(Address2) & GetAddon(Address3) & ObjJosa(Address1) & " 검색했습니다", 0
    ElseIf Address1 <> "" Then 'Search 1,2
        ShellExecute 0, "open", Address1 & URLEncoder_UTF8(Address2) & GetAddon(Address3), "", "", 0
        frmmain.Alert Address1 & URLEncoder_UTF8(Address2) & GetAddon(Address3) & ObjJosa(Address1) & " 검색했습니다", 0
    Else 'Search Google
        ShellExecute 0, "open", "http://www.google.co.kr/#newwindow=1&output=search&sclient=psy-ab&q=" & URLEncoder_UTF8(Address2), "", "", 0
        frmmain.Alert "구글에서 " & Address2 & ObjJosa(Address2) & " 검색했습니다", 1
    End If
    Exit Function
    
ElseIf Functions = "열어" Or Functions = "실행" Then

    If GetProfile(Address1 & Address2, "주소") <> "" Then 'Search Profile(1,주소)
        ShellExecute 0, "open", GetProfile(Address1 & Address2, "주소") & GetAddon(Address3), "", "", 0
        frmmain.Alert GetProfile(Address1 & Address2, "주소") & GetAddon(Address3) & ObjJosa(GetProfile(Address1 & Address2, "주소")) & " 열었습니다", 0
    ElseIf GetProfile(Address1 & Address2, "경로") <> "" Then 'Run Profile(1,경로)
        ShellExecute 0, "open", GetProfile(Address1 & Address2, "경로") & GetAddon(Address3), "", "", 0
        frmmain.Alert GetProfile(Address1 & Address2, "경로") & GetAddon(Address3) & ObjJosa(GetProfile(Address1 & Address2, "경로")) & " 열었습니다", 0
    ElseIf Address1 <> "" And Not GetProfile(Address1, Address2) <> "" Then 'Search 1,2
        ShellExecute 0, "open", Address1 & Address2 & GetAddon(Address3), "", "", 0
        frmmain.Alert Address1 & Address2 & GetAddon(Address3) & ObjJosa(Address1 & Address2) & " 검색했습니다", 0
    ElseIf GetProfile(Address1, Address2) <> "" Then 'Search Profile(1,2)
        ShellExecute 0, "open", GetProfile(Address1, Address2) & GetAddon(Address3), "", "", 0
        frmmain.Alert GetProfile(Address1, Address2) & GetAddon(Address3) & ObjJosa(GetProfile(Address1, Address2)) & " 열었습니다", 0
    ElseIf Address1 & Address2 = "시디롬" Then
        mciSendString "Set CDAudio Door Open Wait", 0&, 0&, 0&
        frmmain.Alert "시디롬을 열었습니다", 0
    Else 'Search 1,2
        ShellExecute 0, "open", "http://www.google.co.kr/#newwindow=1&output=search&sclient=psy-ab&q=" & Address2, "", "", 0
        frmmain.Alert "구글에서 " & Address2 & ObjJosa(Address2) & " 검색했습니다", 0
    End If
    Exit Function

ElseIf Functions = "알려" Then

    If Address1 = "하드" And Address2 = "시리얼 넘버" Then
        frmmain.Alert "하드의 시리얼 넘버는 " & GetMainSerialNumber & " 입니다", 0
    ElseIf Address1 = "OEM" And Address2 = "ID" Then
        GetSystemInfo si
        frmmain.Alert Address1 & "의 " & Address2 & ObjJosa2(Address2) & " " & si.dwOemID & " 입니다", 0
    ElseIf Address1 = "램" And Address2 = "크기" Then
        GetSystemInfo si
        frmmain.Alert Address1 & "의 " & Address2 & ObjJosa2(Address2) & " " & si.dwPageSize & " 입니다", 0
    ElseIf Address1 = "CPU" And Address2 = "갯수" Then
        GetSystemInfo si
        frmmain.Alert Address1 & "의 " & Address2 & ObjJosa2(Address2) & " " & si.dwNumberOrfProcessors & " 입니다", 0
    ElseIf Address1 = "CPU" And Address2 = "종류" Then
        GetSystemInfo si
        frmmain.Alert Address1 & "의 " & Address2 & ObjJosa2(Address2) & " " & si.dwProcessorType & "개 입니다", 0
    End If
    Exit Function
    
ElseIf Functions = "닫아" Then

    KillAppByName Address1 & Address2
    frmmain.Alert Address1 & Address2 & ObjJosa(Address1 & Address2) & " 닫았습니다", 0
    Exit Function
    
ElseIf Functions = "꺼" Then

    If Address1 & Address2 = "컴퓨터" Then
        Shell ("shutdown -s -t 120 -c 컴퓨터를 종료합니다.")
    ElseIf Address1 & Address2 = "모니터" Then
        TurnOffMonitor
    Else
        KillAppByName Address1 & Address2
    End If
    Exit Function

ElseIf Functions = "연주" Then
    
    For i = 0 To UBound(talk_note)
        If Mid(Address2, 1, 1) = talk_note(i) Then
            Select Case i
            Case 0
                PosNote = 0
            Case 1
                PosNote = 2
            Case 2
                PosNote = 4
            Case 3
                PosNote = 5
            Case 4
                PosNote = 7
            Case 5
                PosNote = 9
            Case 6
                PosNote = 11
            End Select
        End If
    Next i
    
    If Mid(Address2, 2, 1) = "#" Then
        PosNote = PosNote + 1
    ElseIf Mid(Address2, 2, 1) = "b" Then
        PosNote = PosNote - 1
    End If
    
    PosNote = PosNote + 36
    
    'Play Note
    If Mid(Address2, 2, 1) = "#" Or Mid(Address2, 2, 1) = "b" Then
        NotePlays() = Split(CalcNote(PosNote, Mid(Address2, 3, Len(Address2) - 2)), ",")
        For w = 0 To UBound(NotePlays)
            PlayNote NotePlays(w)
        Next w
    Else
        NotePlays() = Split(CalcNote(PosNote, Mid(Address2, 2, Len(Address2) - 1)), ",")
        For w = 0 To UBound(NotePlays)
            PlayNote NotePlays(w)
        Next w
    End If
    Exit Function
End If

End Function


