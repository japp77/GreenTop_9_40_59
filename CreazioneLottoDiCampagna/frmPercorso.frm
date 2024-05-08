VERSION 5.00
Begin VB.Form frmPercorso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleziona il percorso"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
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
   ScaleHeight     =   4980
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmPercorso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private F  As FileSystemObject

Private Sub cmdConferma_Click()
    frmMain.txtPercorsoDocumentazione.Text = Dir1.Path
    Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo ERR_Drive1_Change
    Dir1.Path = F.GetDriveName(Drive1.Drive) & "\"
    
Exit Sub
ERR_Drive1_Change:
    MsgBox Err.Description, vbCritical, "Percorso non valido"
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set F = New FileSystemObject
    
    If frmMain.txtPercorsoDocumentazione.Text = "" Then
        Me.Drive1.Drive = F.GetDrive("C:")
        Me.Dir1.Path = Me.Drive1.Drive & "\"
    Else
        Me.Drive1.Drive = F.GetDriveName(frmMain.txtPercorsoDocumentazione.Text)
        Me.Dir1.Path = F.GetFolder(frmMain.txtPercorsoDocumentazione.Text)
    End If
    

End Sub
