VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_reproductor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "Taqui"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "Frm_reproductor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bt_siguiente 
      Appearance      =   0  'Flat
      Caption         =   ">>"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   1200
      Width           =   350
   End
   Begin VB.CommandButton Bt_anterior 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   1200
      Width           =   350
   End
   Begin VB.CommandButton Bt_tocar 
      Appearance      =   0  'Flat
      Caption         =   "Tocar"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Bt_silencio 
      Caption         =   "/"
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Bt_LR 
      Caption         =   "L"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Vold 
      Caption         =   "-"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Vols 
      Caption         =   "+"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton Bt_pausar 
      Appearance      =   0  'Flat
      Caption         =   "Pausar"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1560
      Width           =   855
   End
   Begin VB.Timer Tiempo 
      Enabled         =   0   'False
      Left            =   5880
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar PB_01 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   3
      Scrolling       =   1
   End
   Begin VB.CommandButton Bt_parar 
      Appearance      =   0  'Flat
      Caption         =   "Parar"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Reproducir archivo"
   End
   Begin VB.CommandButton Bt_anadir 
      Appearance      =   0  'Flat
      Caption         =   "Añadir"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Bt_salir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar PB_02 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   3
      Scrolling       =   1
   End
   Begin VB.Label Lb_04 
      Caption         =   "Lista de reproducción"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Lb_03 
      Caption         =   "Volumen"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Lb_02 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Lb_01 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Frm_reproductor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Procedimiento para pausar el programa
Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilisegundos As Long)

'procedimiento para extraer la duración de un archivo mp3
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'FunciÃ³n Api GetShortPathName para obtener _
los paths de los archivos en formato corto
Private Declare Function GetShortPathName _
    Lib "kernel32" _
    Alias "GetShortPathNameA" ( _
        ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, _
        ByVal lBuffer As Long) As Long
  
'FunciÃ³n Api mciExecute para reproducir los archivos de mÃºsica
Private Declare Function mciExecute _
    Lib "winmm.dll" ( _
        ByVal lpstrCommand As String) As Long

Dim ret As Long, path As String, s1 As Long, s2 As Long, s3 As Long, s4 As Long
Dim toca As Boolean, tocar As Boolean, durac As Long, s5 As Long
Dim duraclb As Integer, contc As Integer, contTc As Integer
Dim hor As Integer, min As Integer, seg As Integer, aux As Integer
Dim Archivos() As String
Dim cambc As Boolean, tiempod As Double, duracM As Integer
Dim volm As Double, dirar As String
Dim Midi As Boolean, Tant As Boolean


Private Sub Bt_anadir_Click()
    'Array dinámico de tipo String
    Dim flag As String
    Dim i As Integer
      
    'Flags para el commondialog para que permita selección múltiple
    flag = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    With CD1
        On Error Resume Next
        .MaxFileSize = 32000
        'Le pasamos el flag a la propiedad Flags
        .Flags = flag
        
        '.Filter = "Archivos Wav|*.wav|Archivos Mp3|*.mp3|Archivos MIDI|*.mid"
        .Filter = "Archivos Mp3|*.mp3|Archivos Mp4|*.mp4|Archivos wav|*.wav|Archivos MIDI|*.mid"
        
        .ShowOpen
        
        If .FileName = "" Then
            'Habilitar "Iniciar"
            Exit Sub
        Else
            'Guardamos en el array dinámico los archivos con la función Split _
            indicandole como separador el Chr(0)
            Archivos = Split(CD1.FileName, Chr(0))
            If UBound(Archivos) <> 0 Then
                For i = 0 To UBound(Archivos) - 1
                    If InStr(1, Archivos(0), "\") = Len(Archivos(0)) Then
                        Frm_LR.LR_03.AddItem (Archivos(0) + Archivos(i + 1))
                        durac = CLng(PlayingTime(Archivos(0) + Archivos(i + 1)))
                    Else
                        Frm_LR.LR_03.AddItem (Archivos(0) & "\" & Archivos(i + 1))
                        durac = CLng(PlayingTime(Archivos(0) & "\" & Archivos(i + 1)))
                    End If
                    Frm_LR.LR_04.AddItem (durac)
                    Frm_LR.LR_02.AddItem (TimeString(durac))
                    Frm_LR.LR_01.AddItem (Archivos(i + 1))
                Next i
            Else
                Frm_LR.LR_03.AddItem (Archivos(0))
                durac = CLng(PlayingTime(Archivos(0)))
                Frm_LR.LR_04.AddItem (durac)
                dirar = Archivos(0)
                Dim longd As Integer
                While InStr(1, dirar, "\") <> 0
                    aux = InStr(1, dirar, "\")
                    longd = Len(dirar)
                    dirar = Right(dirar, longd - aux)
                Wend
                Frm_LR.LR_01.AddItem (dirar)
                Frm_LR.LR_02.AddItem (TimeString(durac))
            End If
            contTc = Frm_LR.LR_01.ListCount
            
            If s3 = 0 Then
                'Le pasamos a la sub que obtiene con _
                el Api GetShortPathName el nombre corto del archivo
                PathCorto Frm_LR.LR_03.List(0)
                Lb_01.Caption = contc + 1 & ".- " & Frm_LR.LR_01.List(0)
    
                s3 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now)
                cambc = True
                toca = True
                Sleep 500 'Espera 0.5 segundos
                Tiempo.Enabled = True
                Tiempo.Interval = 100
                Bt_tocar.Enabled = False
                mciExecute "Play " & path
            End If
        End If
    End With
End Sub

Private Sub Bt_anterior_Click()
    
    ejecutar ("Stop ")
    PB_01.Value = 0
    Lb_02.Caption = ""
    s5 = 0
    Sleep 500
    If Tant = False Then
        Tant = True
        s3 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now)
        ejecutar ("Play ")
        toca = True
    Else
        If contc <> 0 Then
            ejecutar ("Stop ")
            PB_01.Value = 0
            Lb_02.Caption = ""
            s5 = 0
            contc = contc - 1
            PB_01.Value = 0
            PathCorto Frm_LR.LR_03.List(contc)
            Lb_01.Caption = contc + 1 & ".- " & Frm_LR.LR_01.List(contc)
            mciExecute "Play " & path
            s3 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now)
            cambc = False
            Tant = False
        End If
    End If
End Sub

Private Sub Bt_LR_Click()
    If Frm_LR.Visible = False Then
        Frm_LR.Enabled = True
        Frm_LR.Show
    ElseIf Frm_LR.Visible = True Then
        Frm_LR.Enabled = False
        Frm_LR.Hide
    End If
End Sub

Private Sub Bt_parar_Click()
    'Le pasamos el comando Stop
    ejecutar ("Stop ")
    toca = False
    PB_01.Value = 0
    Lb_02.Caption = ""
    Bt_parar.Enabled = False
    Bt_tocar.Enabled = True
    s5 = 0
    'Habilitar "Stop"
End Sub

Private Sub Bt_pausar_Click()
    ejecutar ("Pause ")
    toca = False
    s4 = s2
    Bt_pausar.Enabled = False
    Bt_tocar.Enabled = True
End Sub

Private Sub Bt_salir_Click()
    mciExecute "Close All"
    End
End Sub

'Private Sub Bt_tocar_Click()
'    Reproductor.Open ("F:\códigos\We will rock you.mp3")
'End Sub

'Sub que obtiene el path corto del archivo a reproducir
Private Sub PathCorto(Archivo As String)
Dim temp As String * 250 'Buffer
    ret = GetShortPathName(Archivo, temp, 164)
    'Sacamos los nulos al path
  path = String(255, 0)
    'Obtenemos el Path corto
      path = Replace(temp, Chr(0), "")
End Sub

Private Sub ejecutar(comando As String)
    If path = "" Then MsgBox "Error", vbCritical: Exit Sub
    'Llamamos a mciExecute pasandole un string que tiene el comando y la ruta
    mciExecute comando & path
End Sub

Private Sub Bt_siguiente_Click()
    Dim iaux As Variant
    
    Sleep 250
    s3 = s2 - Frm_LR.LR_04.List(contc)
    iaux = mciSendString("seek " & path & " to end", 0&, 0, 0)
End Sub

Private Sub Bt_silencio_Click()
    volm = 0
    PB_02.Value = volm
    SetVol Percent(100, CInt(volm * 100)), Midi
    VolDown
End Sub

Private Sub Bt_tocar_Click()
    If s3 = 0 Then
        'Array dinámico de tipo String
        Dim flag As String
        Dim i As Integer
          
        'Flags para el commondialog para que permita selección múltiple
        flag = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
        With CD1
            On Error Resume Next
            .MaxFileSize = 32000
            'Le pasamos el flag a la propiedad Flags
            .Flags = flag
            
            '.Filter = "Archivos Wav|*.wav|Archivos Mp3|*.mp3|Archivos MIDI|*.mid"
            .Filter = "Archivos Mp3|*.mp3|Archivos Mp4|*.mp4|Archivos wav|*.wav|Archivos MIDI|*.mid"
            
            .ShowOpen
            If .FileName = "" Then
                'Habilitar "Iniciar"
                Exit Sub
            Else
                'Guardamos en el array dinámico los archivos con la función Split _
                indicandole como separador el Chr(0)
                Archivos = Split(CD1.FileName, Chr(0))
                If UBound(Archivos) <> 0 Then
                    For i = 0 To UBound(Archivos) - 1
                        If InStr(1, Archivos(0), "\") = Len(Archivos(0)) Then
                            Frm_LR.LR_03.AddItem (Archivos(0) + Archivos(i + 1))
                            durac = CLng(PlayingTime(Archivos(0) + Archivos(i + 1)))
                        Else
                            Frm_LR.LR_03.AddItem (Archivos(0) & "\" & Archivos(i + 1))
                            durac = CLng(PlayingTime(Archivos(0) & "\" & Archivos(i + 1)))
                        End If
                        Frm_LR.LR_04.AddItem (durac)
                        Frm_LR.LR_02.AddItem (TimeString(durac))
                        Frm_LR.LR_01.AddItem (Archivos(i + 1))
                    Next i
                Else
                    Frm_LR.LR_03.AddItem (Archivos(0))
                    durac = CLng(PlayingTime(Archivos(0)))
                    Frm_LR.LR_04.AddItem (durac)
                    dirar = Archivos(0)
                    Dim longd As Integer
                    While InStr(1, dirar, "\") <> 0
                        aux = InStr(1, dirar, "\")
                        longd = Len(dirar)
                        dirar = Right(dirar, longd - aux)
                    Wend
                    Frm_LR.LR_01.AddItem (dirar)
                    Frm_LR.LR_02.AddItem (TimeString(durac))
                End If
                contTc = Frm_LR.LR_01.ListCount
                
                'Le pasamos a la sub que obtiene con _
                el Api GetShortPathName el nombre corto del archivo
                PathCorto Frm_LR.LR_03.List(0)
                Lb_01.Caption = contc + 1 & ".- " & Frm_LR.LR_01.List(0)
                
                s3 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now)
                cambc = True
                toca = True
                Sleep 500 'Espera 0.5 segundos
                Tiempo.Enabled = True
                Tiempo.Interval = 100
                mciExecute "Play " & path
            End If
        End With
    Else
        ejecutar ("Play ")
        toca = True
        Lb_01.Caption = contc + 1 & ".- " & Frm_LR.LR_01.List(contc)
        If Bt_pausar.Enabled = False Then s3 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now) - s4
        If Bt_parar.Enabled = False Then s3 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now)
    End If
    Bt_tocar.Enabled = False
    Bt_pausar.Enabled = True
    Bt_parar.Enabled = True
End Sub

Private Sub Form_Load()
    
    toca = False
    tocar = False
    s2 = 0
    s5 = 0
    duraclb = 0
    contc = 0
    Midi = False
    volm = 1
    PB_02.Max = 1
    PB_02.min = 0
    PB_02.Value = volm
    SetVol Percent(100, CInt(volm * 100)), Midi
    duracM = 0
    contTc = 0
    Tant = False
End Sub

Private Sub Lb_02_Click()
    duraclb = duraclb + 1
    If duraclb >= 3 Then duraclb = 0
End Sub

Private Sub PB_01_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iaux As Variant
    'PB_01.Value = CInt(x)
    'Frm_LR.LR_04.List(contc)
    s5 = x / PB_01.Width * Frm_LR.LR_04.List(contc) - s2
    's5 = x / 10 - s2
    If duraclb = 0 Then
        If Frm_LR.LR_04.List(contc) > 3600 Then
            Lb_02.Caption = TimeString(s2) & " / " & Format$(Frm_LR.LR_04.List(contc) / 86400, "hh:nn:ss")
        Else
            Lb_02.Caption = TimeString(s2) & " / " & Format$(Frm_LR.LR_04.List(contc) / 86400, "nn:ss")
        End If
    End If
    If duraclb = 1 Then Lb_02.Caption = Format(Frm_LR.LR_04.List(contc) - s2, "0.0")
    If duraclb = 2 Then Lb_02.Caption = Format(s2, "0.0")
    
    Sleep 1500 'Espera 0.5 segundos
    
    iaux = mciSendString("play " & path & " from " & s5 * 1000, 0&, 0, 0)
End Sub

Private Sub PB_02_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x > PB_02.Value Then
        volm = x / PB_02.Width
        PB_02.Value = x / PB_02.Width
        SetVol Percent(100, CInt(volm * 100)), Midi
        VolUp
    Else
        volm = x / PB_02.Width
        PB_02.Value = x / PB_02.Width
        SetVol Percent(100, CInt(volm * 100)), Midi
        VolDown
    End If
End Sub

Private Sub Tiempo_Timer()
    If toca = True Then
        s2 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now) - s3 + s5
        If Tant = True Then If s2 > 10 Then Tant = False
        s1 = s2
        'If s2 <= durac(contc) And durac(contc) <> 0 Then
        If s2 <= Frm_LR.LR_04.List(contc) And Frm_LR.LR_04.List(contc) <> 0 Then
            If duraclb = 0 Then
                'If durac(contc) > 3600 Then
                If Frm_LR.LR_04.List(contc) > 3600 Then
                    'Lb_02.Caption = TimeString(s2) & " / " & Format$(durac(contc) / 86400, "hh:nn:ss")
                    Lb_02.Caption = TimeString(s2) & " / " & Format$(Frm_LR.LR_04.List(contc) / 86400, "hh:nn:ss")
                Else
                    'Lb_02.Caption = TimeString(s2) & " / " & Format$(durac(contc) / 86400, "nn:ss")
                    Lb_02.Caption = TimeString(s2) & " / " & Format$(Frm_LR.LR_04.List(contc) / 86400, "nn:ss")
                End If
            End If
            'If duraclb = 1 Then Lb_02.Caption = Format(durac(contc) - s2, "0.0")
            If duraclb = 1 Then Lb_02.Caption = Format(Frm_LR.LR_04.List(contc) - s2, "0.0")
            If duraclb = 2 Then Lb_02.Caption = Format(s2, "0.0")
            'PB_01.Value = s1 * 100 / durac(contc)
            PB_01.Value = s1 * 100 / Frm_LR.LR_04.List(contc)
        'ElseIf s2 > durac(contc) And s2 <> 0 And durac(contc) <> 0 Then
        ElseIf s2 > Frm_LR.LR_04.List(contc) And s2 <> 0 And Frm_LR.LR_04.List(contc) <> 0 Then
            'If contc <= UBound(Archivos) Then
            If contc + 1 <= Frm_LR.LR_01.ListCount Then
                If contc + 2 <= Frm_LR.LR_01.ListCount Then
                    cambc = True
                Else
                    GoTo 1
                End If
                'Le pasamos a la sub que obtiene con _
                el Api GetShortPathName el nombre corto del archivo
                contc = contc + 1
                PB_01.Value = 0
                PathCorto Frm_LR.LR_03.List(contc)
                Lb_01.Caption = contc + 1 & ".- " & Frm_LR.LR_01.List(contc)
                mciExecute "Play " & path
                s3 = Hour(Now) * 60 * 60 + Minute(Now) * 60 + Second(Now)
                cambc = False
            Else
1:              mciExecute "Close All"
                PB_01.Value = 0
                Lb_01.Caption = ""
                Lb_02.Caption = ""
                toca = False
                Bt_tocar.Enabled = True
                Bt_pausar.Enabled = False
                Bt_parar.Enabled = False
                contc = 0
            End If
        End If
    End If
End Sub

Private Function PlayingTime(strFileToPlay As String) As String

    Dim TotalTime As String * 128
    Dim sFileName As String
    Dim sBuffer As String
    Dim lngTime As Double
    
    
    PlayingTime = 0 'Error default
    
    sBuffer = String$(260, vbNullChar)
    GetShortPathName strFileToPlay, sBuffer, 260
    strFileToPlay = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)

    If mciSendString("set " & strFileToPlay & " time format ms", TotalTime, 0&, 0) = 0 Then
        'MsgBox "Proceso duración" & mciSendString("set " & strFileToPlay & " time format ms", TotalTime, 0&, 0) & " - " & mciSendString("status " & strFileToPlay & " length", TotalTime, 128&, 0)
        Sleep 500 'Espera 0.5 segundos
        If mciSendString("status " & strFileToPlay & " length", TotalTime, 128&, 0) = 0 Then
            PlayingTime = Val(TotalTime)
            lngTime = PlayingTime / 1000
            'PlayingTime = Format$(lngTime / 86400, "nn:ss")
            tiempod = lngTime
            hor = CInt(Format$(lngTime / 86400, "hh"))
            min = CInt(Format$(lngTime / 86400, "nn"))
            seg = CInt(Format$(lngTime / 86400, "ss"))
            PlayingTime = CStr(hor * 3600 + min * 60 + seg)
        End If
    
        mciSendString "close " & strFileToPlay, TotalTime, 0&, 0&
    End If
End Function

Private Sub VolUp()
    If GetVol(Midi) < 100 Then SetVol GetVol(Midi) + 1, Midi
End Sub

Private Sub VolDown()
    If GetVol(Midi) > 0 Then SetVol GetVol(Midi) - 1, Midi
End Sub

Private Sub Vold_Click()
    If volm >= 0 Then volm = volm - 0.05
    If volm < 0 Then volm = 0
    PB_02.Value = volm
    SetVol Percent(100, CInt(volm * 100)), Midi
    VolDown
End Sub

Private Sub Vols_Click()
    If volm <= 1 Then volm = volm + 0.05
    If volm > 1 Then volm = 1
    PB_02.Value = volm
    SetVol Percent(100, CInt(volm * 100)), Midi
    VolUp
End Sub
