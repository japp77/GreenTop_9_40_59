VERSION 5.00
Begin VB.Form frmNuovaCartella 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inserimento nuova cartella"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtNuovaCartella 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   7215
   End
   Begin VB.DirListBox DirSelezionato 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmNuovaCartella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()
Dim F As FileSystemObject

Set F = New FileSystemObject

If Len(Trim(Me.txtNuovaCartella.Text)) > 0 Then
    If F.FolderExists(Me.DirSelezionato.Path & "\" & Me.txtNuovaCartella.Text) = True Then
        MsgBox "Cartella esistente", vbInformation, TheApp.FunctionName
        Exit Sub
    Else
        F.CreateFolder Me.DirSelezionato.Path & "\" & Me.txtNuovaCartella.Text
    End If
End If

Set F = Nothing
Unload Me
End Sub

Private Sub Form_Load()
    Me.DirSelezionato.Path = frmMain.txtPercorsoSelezionato.Text
    Me.DirSelezionato.Refresh
End Sub
