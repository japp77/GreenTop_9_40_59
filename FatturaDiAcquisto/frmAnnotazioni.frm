VERSION 5.00
Begin VB.Form frmAnnotazioni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Annotazioni"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnnotazioni.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmAnnotazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyEscape) Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Text1.Text = frmMain.txtAnnotazioni.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.txtAnnotazioni.Text = Me.Text1.Text
End Sub
