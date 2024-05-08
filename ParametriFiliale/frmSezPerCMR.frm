VERSION 5.00
Begin VB.Form frmSezPerCMR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sezionali per CMR"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4260
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSezPerCMR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstUtenti 
      Height          =   4335
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
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
      Top             =   4560
      Width           =   3975
   End
End
Attribute VB_Name = "frmSezPerCMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConferma_Click()
    Conferma
End Sub

Private Sub Form_Load()
    GET_LISTA_SEZIONALI
End Sub
Private Sub GET_LISTA_SEZIONALI()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim i As Long

sSQL = "SELECT * FROM Sezionale "
sSQL = sSQL & " WHERE IDFiliale=" & frmMain.cboFiliale.CurrentID
sSQL = sSQL & " ORDER BY Sezionale"

Set rs = Cn.OpenResultset(sSQL)
i = 0
While Not rs.EOF
    Me.lstUtenti.AddItem fnNotNull(rs!Sezionale)
    Me.lstUtenti.ItemData(i) = fnNotNullN(rs!IDSezionale)
    Me.lstUtenti.Selected(i) = GET_SEZIONALE_SELEZIONATO(fnNotNullN(rs!IDSezionale), frmMain.cboFiliale.CurrentID)
    i = i + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_SEZIONALE_SELEZIONATO(IDSezionale As Long, IDFiliale As Long) As Boolean
On Error GoTo ERR_GET_UTENTE_SELEZIONATO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim i As Integer

sSQL = "SELECT * "
sSQL = sSQL & " FROM RV_POSezionalePerCMR "
sSQL = sSQL & " WHERE IDSezionale=" & IDSezionale
sSQL = sSQL & " AND IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_SEZIONALE_SELEZIONATO = False
Else
    GET_SEZIONALE_SELEZIONATO = True
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_UTENTE_SELEZIONATO:
    MsgBox Err.Description, vbCritical, "GET_SEZIONALE_SELEZIONATO"
End Function
Private Sub Conferma()
On Error GoTo ERR_CONFERMA
Dim sSQL As String

sSQL = "DELETE FROM RV_POSezionalePerCMR "
sSQL = sSQL & "WHERE IDFiliale=" & frmMain.cboFiliale.CurrentID
Cn.Execute sSQL

For i = 0 To Me.lstUtenti.ListCount - 1
    If Me.lstUtenti.Selected(i) = True Then
        sSQL = "INSERT INTO RV_POSezionalePerCMR ("
        sSQL = sSQL & "IDSezionale, IDFiliale) "
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
