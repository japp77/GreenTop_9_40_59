VERSION 5.00
Begin VB.Form frmPesiMedi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RIEPILOGO PESI/PEZZI MEDI"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPesiMedi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPesoMedio 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   5895
   End
   Begin VB.TextBox txtPezzoMedio 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2040
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PESO MEDIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PEZZO MEDIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   6015
   End
End
Attribute VB_Name = "frmPesiMedi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

    RefreshCalcolo
    
    frmMain.SetFocus
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = Screen.Width - Me.Width - 100
    FORM_PESI_MEDI_SHOW = True
End Sub
Public Sub RefreshCalcolo()
On Error Resume Next
Dim PesoMedio As Double
Dim PezzoMedio As Double

    PesoMedio = 0
    PezzoMedio = 0
    
    If (frmMain.txtColli.Value > 0) Then
        PesoMedio = frmMain.txtPesoNetto.Value / frmMain.txtColli.Value
        PezzoMedio = frmMain.txtPezzi.Value / frmMain.txtColli.Value
    End If
    Me.txtPesoMedio.Text = FormatNumber(PesoMedio, 4)
    Me.txtPezzoMedio.Text = FormatNumber(PezzoMedio, 2)


End Sub


Private Sub Form_Unload(Cancel As Integer)
FORM_PESI_MEDI_SHOW = False

End Sub
