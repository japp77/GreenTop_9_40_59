VERSION 5.00
Begin VB.Form frmEseguiScript 
   Caption         =   "Esegui script"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEseguiScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Esegui"
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtVista 
      Height          =   5655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   11055
   End
End
Attribute VB_Name = "frmEseguiScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
    Screen.MousePointer = 11
    CnDMT.Execute txtVista.Text
    Screen.MousePointer = 0
    
    MsgBox "Script eseguito con successo!", vbInformation, "Controllo dati"
    
Exit Sub
ERR_Command1_Click:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Command1_Click"
End Sub
