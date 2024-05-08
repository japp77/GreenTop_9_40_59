VERSION 5.00
Begin VB.Form frmErrore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ERRORE"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      Picture         =   "frmErrore.frx":0000
      ScaleHeight     =   2025
      ScaleWidth      =   8745
      TabIndex        =   2
      Top             =   0
      Width           =   8775
   End
   Begin VB.TextBox txtCodiceErrore 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2280
      TabIndex        =   1
      Top             =   7680
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2040
      Width           =   8775
   End
End
Attribute VB_Name = "frmErrore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
GET_SUONO_ERRORE
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
If KeyCode = vbKeyReturn Then
    If Me.txtCodiceErrore.Text = CODICE_GESTIONE_ERRORI Then
        Unload Me
    Else
        GET_SUONO_ERRORE
    End If
End If
End Sub

Private Sub Form_Load()

    
    Me.Text1.Text = ERRORE_EVASIONE
End Sub
Private Function GET_SUONO_ERRORE()

PlaySound App.Path & "\connect.wav", 0, SND_FILENAME Or SND_ASYNC


'Dim I As Integer
'    Beep 1000, 250
'    Beep 1350, 250

End Function
