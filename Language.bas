Attribute VB_Name = "Language"
Option Explicit
Public talk_memory As String '기억
Public talk_memory2 As String '기억(2) - 위치
Public talk_memory3 As String '기억(3) - 속성
Public talk_result As String '기억(4) - 결과
Public talk_search As String '내용 검색
Public talk_process() As String '행동
Public talk_split() As String '분할
Public talk_inv() As String '보조사
Public talk_do() As String '진행완료
Public talk_and() As String '접속부사
Public talk_ask() As String '질문
Public talk_is() As String '서술격조사
Public talk_to() As String '격조사
Public talk_in() As String '격조사2
Public talk_of() As String '속성
Public talk_note() As String '노트
Public var_name As New Collection '이름
Public var_cont As New Collection '내용
Public var_profile As New Collection '속성
Public var_addon As New Collection '추과
Public chk_and As Boolean '접속부사 체크
Public chk_inv As Boolean '접속부사 체크

Public Function Set_Function()

ReDim talk_process(5)
talk_process(0) = "해줘"
talk_process(1) = "싶다"
talk_process(2) = "싶어"
talk_process(4) = "라"
talk_process(5) = "줘"

ReDim talk_in(0)
talk_in(0) = "에서"

ReDim talk_of(0)
talk_of(0) = "의"

ReDim talk_to(3)
talk_to(0) = "을"
talk_to(1) = "를"
talk_to(2) = "이"
talk_to(3) = "가"

ReDim talk_is(0)
talk_is(0) = "다"

ReDim talk_inv(1) '조사
talk_inv(0) = "은"
talk_inv(1) = "는"

ReDim talk_do(4) '진행완료
talk_do(0) = "한"
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

ReDim talk_ask(4) '질문
talk_ask(0) = "뭐지"
talk_ask(1) = "에 저장된 값은?"
talk_ask(2) = "뭐야"
talk_ask(3) = "는?"
talk_ask(4) = "은?"

ReDim talk_note(6) '노트
talk_note(0) = "C"
talk_note(1) = "D"
talk_note(2) = "E"
talk_note(3) = "F"
talk_note(4) = "G"
talk_note(5) = "A"
talk_note(6) = "B"

End Function

Public Function ObjJosa(Str As String) As String
Dim LastStr As String, LastAsc As Long, JongN As Integer

LastStr = Right(Str, 1)
LastAsc = AscW(LastStr)
LastAsc = LastAsc - &HAC00

JongN = LastAsc Mod 28

If JongN = 0 And (LastStr >= "ㄱ" And LastStr <= "?") Then
ObjJosa = "를"
ElseIf (LastStr >= "ㄱ" And LastStr <= "?") Then
ObjJosa = "을"
Else
ObjJosa = "를"
End If

End Function

Public Function ObjJosa2(Str As String) As String
Dim LastStr As String, LastAsc As Long, JongN As Integer

LastStr = Right(Str, 1)
LastAsc = AscW(LastStr)
LastAsc = LastAsc - &HAC00

JongN = LastAsc Mod 28

If JongN = 0 And (LastStr >= "ㄱ" And LastStr <= "?") Then
ObjJosa2 = "는"
ElseIf (LastStr >= "ㄱ" And LastStr <= "?") Then
ObjJosa2 = "은"
Else
ObjJosa2 = "는"
End If

End Function

