Attribute VB_Name = "Módulo"
Public Declare Function waveOutSetVolume Lib "Winmm" (ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Public Declare Function waveOutGetVolume Lib "Winmm" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
Public Declare Function midiOutSetVolume Lib "Winmm" (ByVal mDeviceID As Integer, ByVal dmVolume As Long) As Integer
Public Declare Function midiOutGetVolume Lib "Winmm" (ByVal mDeviceID As Integer, dmVolume As Long) As Integer

Public Function GetVol(Optional Midi As Boolean = False) As Integer
    Dim v As Long
    Dim x As Long
    Dim xh As String
    If Midi Then
        Call midiOutGetVolume(0, x)
    Else
        Call waveOutGetVolume(0, x)
    End If
    xh = HexDec(Right$(Hex$(x), 4)) ', 16, 10) ') ', 16, 10))
    v = Round(Val(xh) / 655.36)
    GetVol = v
End Function

Public Sub SetVol(Volume As Integer, Optional Midi As Boolean = False)
    'v = 15
    Dim x As Long
    If Volume > 50 Then
        If Volume = 100 Then
            If Midi Then
                Call midiOutSetVolume(0, &HFFFFFFFF)
            Else
                Call waveOutSetVolume(0, &HFFFFFFFF)
            End If
        Else
            x = -((32767 / 50) * (100 - Volume))
            If Midi Then
                 Call midiOutSetVolume(0, x + (x * 65536))
            Else
                 Call waveOutSetVolume(0, x + (x * 65536))
            End If
        End If
    Else
        x = Int((32767 / 50) * Volume)
        If Midi Then
            Call midiOutSetVolume(0, x + (x * 65536))
        Else
            Call waveOutSetVolume(0, x + (x * 65536))
        End If
    End If
End Sub

Public Function HexDec(h As String) As Long
    Dim i As Integer
    Dim cnt As Long
    h = LCase(h)
    For i = 1 To Len(h)
        Select Case Mid(h, i, 1)
            Case "1": cnt = cnt + 1 * 16 ^ (Len(h) - i)
            Case "2": cnt = cnt + 2 * 16 ^ (Len(h) - i)
            Case "3": cnt = cnt + 3 * 16 ^ (Len(h) - i)
            Case "4": cnt = cnt + 4 * 16 ^ (Len(h) - i)
            Case "5": cnt = cnt + 5 * 16 ^ (Len(h) - i)
            Case "6": cnt = cnt + 6 * 16 ^ (Len(h) - i)
            Case "7": cnt = cnt + 7 * 16 ^ (Len(h) - i)
            Case "8": cnt = cnt + 8 * 16 ^ (Len(h) - i)
            Case "9": cnt = cnt + 9 * 16 ^ (Len(h) - i)
            Case "a": cnt = cnt + 10 * 16 ^ (Len(h) - i)
            Case "b": cnt = cnt + 11 * 16 ^ (Len(h) - i)
            Case "c": cnt = cnt + 12 * 16 ^ (Len(h) - i)
            Case "d": cnt = cnt + 13 * 16 ^ (Len(h) - i)
            Case "e": cnt = cnt + 14 * 16 ^ (Len(h) - i)
            Case "f": cnt = cnt + 15 * 16 ^ (Len(h) - i)
        End Select
    Next i
    HexDec = cnt
End Function

Public Function Percent(Val As Long, Percnt As Integer) As Long
    Percent = Val * (Percnt / 100)
End Function

Public Function TimeString(Seconds As Long) As String
    'convert seconds to mm:ss format
    Dim Mins As Long
    Dim Hors As Integer
    
    If Seconds < 60 Then TimeString = "0:" & Right("0" & Seconds, 2)
    If Seconds > 59 Then
        Mins = Int(Seconds / 60)
        If Mins > 59 Then
            Hors = Int(Mins / 60)
            Mins = Mins - (Hors * 60)
            Seconds = Seconds - (Mins * 60)
            TimeString = Right("0" + Str(Hors), 2) & ":" & Right("0" & Mins, 2) & ":" & Right("0" & Seconds, 2)
        Else
            Seconds = Seconds - (Mins * 60)
            TimeString = Right("0" + Str(Mins), 2) & ":" & Right("0" & Seconds, 2)
        End If
    End If
    If InStr(1, TimeString, " ") Then Mid(TimeString, InStr(1, TimeString, " "), 1) = ""
End Function

