Attribute VB_Name = "SysMidi"
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Boolean
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Public lNote As Long
Public rc As Long
Public hMidi As Long
Public Channel As Long
Public KeyMap(255) As Long
Public Key As Integer
Public Pitch As Long
Public Velocity As Long
Public noteLong As Long

Public Sub PlayNote(ByVal Note As Long)
On Error Resume Next
Dim midimsg As Long

midimsg = &H90 + Channel + ((Pitch + Note) * &H100) + (Velocity * &H10000)
midiOutShortMsg hMidi, midimsg

lNote = Note
End Sub

Public Sub InitializeMidi() 'Initialize Midi
On Error Resume Next
Dim x As Long

Key = 1
Pitch = 23
Velocity = 64

midiOutClose hMidi
rc = midiOutOpen(hMidi, curDevice, 0, 0, 0)

If rc = 4 Then
    frmmain.Alert "미디를 사용할 수 없습니다.", 2
End If
    
End Sub

Public Function CalcNote(chdRoot, chdType) 'Calculate Chord Note
On Error Resume Next
Dim Tention() As String
Dim Note() As String
Dim Tmp$
Dim i, z As Long

'Array Note
Select Case chdType
    Case "Maj" Or "Major" Or "M" Or "maj" Or "major"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "sus4"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
    Case "sus2"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 2 'Perfect 2
        Note(3) = chdRoot + 7 'Perfect 5
    Case "m" Or "min" Or "minor" Or "Minor"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
    Case "+" Or "aug" Or "Aug" Or "augment" Or "Augment"
        ReDim Note(3)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
    Case "Maj7" Or "Major7" Or "M7" Or "maj7" Or "major7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "Maj7b5" Or "Major7b5" Or "M7b5" Or "maj7b5" Or "major7b5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 6 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "Maj7+5" Or "Major7+5" Or "M7+5" Or "maj7+5" Or "major7+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 11 'Major 7
    Case "mM7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 11 'Major 7
    Case "m7" Or "min7" Or "minor7" Or "Minor7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "m+5" Or "min+5" Or "minor+5" Or "Minor+5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 8 'Augmented 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "+7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 8 'Minor 6
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus4"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 5 'Perfect 4
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "7sus2"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 2 'Perfect 2
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 10 'Minor 7
    Case "Dim" Or "Diminished" Or "dim" Or "diminished"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
    Case "mb5" Or "minb5" Or "minorb5" Or "Minorb5"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 10 'Major 7
    Case "Dim7" Or "Diminished 7" Or "dim7" Or "diminished 7"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 6 'Diminished 5
        Note(4) = chdRoot + 9 'Minor 7
    Case "Maj6" Or "Major6" Or "M6" Or "maj6" Or "major6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 4 'Major 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
    Case "m6" Or "min6" Or "minor6" Or "Minor6"
        ReDim Note(4)
        Note(1) = chdRoot
        Note(2) = chdRoot + 3 'Minor 3
        Note(3) = chdRoot + 7 'Perfect 5
        Note(4) = chdRoot + 9 'Major 6
End Select

'Merge Note
For z = 1 To UBound(Note)
    If Note(z) <> "" Then
        If Tmp = "" Then
            Tmp = Tmp & Note(z)
        Else
            Tmp = Tmp & "," & Note(z)
        End If
    End If
Next z

CalcNote = Tmp 'Output Temp
Tmp = "" 'Cleanup Temp

End Function

