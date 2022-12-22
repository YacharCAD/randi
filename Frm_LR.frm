VERSION 5.00
Begin VB.Form Frm_LR 
   Caption         =   "Lista de reproducción"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox LR_02 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Bt_salir 
      Caption         =   "Salir"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox LR_01 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.ListBox LR_03 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2565
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ListBox LR_04 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2565
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Frm_LR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bt_salir_Click()
    Frm_LR.Hide
End Sub

Private Sub LR_01_DblClick()
    MsgBox LR_01.ListIndex
End Sub
