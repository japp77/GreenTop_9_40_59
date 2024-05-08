VERSION 5.00
Begin VB.Form frmAttesa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   40
      Picture         =   "frmAttesa.frx":0000
      ScaleHeight     =   1815
      ScaleWidth      =   8295
      TabIndex        =   1
      Top             =   40
      Width           =   8295
      Begin VB.Label lblInfo2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ATTENDERE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1500
         Width           =   8055
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7920
      Top             =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000040C0&
      BorderWidth     =   3
      Height          =   2775
      Left            =   15
      Top             =   15
      Width           =   8350
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   8175
   End
End
Attribute VB_Name = "frmAttesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    Me.Timer1.Enabled = True
End Sub


Public Function FORM_ATTESA_UNLOAD()
    Me.Timer1.Enabled = False

    Unload Me
End Function

Private Sub Form_Unload(Cancel As Integer)
    Me.Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    DoEvents
End Sub
