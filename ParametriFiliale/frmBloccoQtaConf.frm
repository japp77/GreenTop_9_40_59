VERSION 5.00
Begin VB.Form frmBloccoQtaConf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utenti blocco Q.tà conferita lavorata"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBloccoQtaConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstUtenti 
      Height          =   4335
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
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
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   3975
   End
End
Attribute VB_Name = "frmBloccoQtaConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListaUtenti()
On Error GoTo ERR_ListaUtenti
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim i As Integer

sSQL = "SELECT IDUtente, Utente "
sSQL = sSQL & "FROM Utente"

Set rs = Cn.OpenResultset(sSQL)
i = 0
While Not rs.EOF
    Me.lstUtenti.AddItem fnNotNull(rs!Utente)
    Me.lstUtenti.ItemData(i) = fnNotNullN(rs!IDUtente)
    Me.lstUtenti.Selected(i) = GET_UTENTE_SELEZIONATO(fnNotNullN(rs!IDUtente), frmMain.cboFiliale.CurrentID)
    i = i + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_ListaUtenti:
    MsgBox Err.Description, vbCritical, "ListaUtenti"
End Sub
Private Function GET_UTENTE_SELEZIONATO(IDUtente As Long, IDFiliale As Long) As Boolean
On Error GoTo ERR_GET_UTENTE_SELEZIONATO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim i As Integer

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POUtentiBloccoQtaConfInLav "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_UTENTE_SELEZIONATO = False
Else
    GET_UTENTE_SELEZIONATO = True
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_UTENTE_SELEZIONATO:
    MsgBox Err.Description, vbCritical, "GET_UTENTE_SELEZIONATO"
End Function

Private Sub cmdConferma_Click()
    Conferma
    
End Sub

Private Sub Form_Load()
    ListaUtenti
End Sub

Private Sub Conferma()
On Error GoTo ERR_CONFERMA
Dim i As Long
Dim sSQL As String

sSQL = "DELETE RV_POUtentiBloccoQtaConfInLav "
sSQL = sSQL & "WHERE IDFiliale=" & frmMain.cboFiliale.CurrentID
Cn.Execute sSQL

For i = 0 To Me.lstUtenti.ListCount - 1
    If Me.lstUtenti.Selected(i) = True Then
        sSQL = "INSERT INTO RV_POUtentiBloccoQtaConfInLav ("
        sSQL = sSQL & "IDUtente, IDFiliale) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Me.lstUtenti.ItemData(i) & ", "
        sSQL = sSQL & frmMain.cboFiliale.CurrentID
        sSQL = sSQL & ")"
        Cn.Execute sSQL
    End If
Next

Unload Me

Exit Sub

ERR_CONFERMA:
    MsgBox Err.Description, vbCritical, "CONFERMA"
End Sub

