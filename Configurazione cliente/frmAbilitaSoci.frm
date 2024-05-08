VERSION 5.00
Begin VB.Form frmAbilitaSoci 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abilita soci"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbilitaSoci.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   9480
      Width           =   6615
   End
   Begin VB.ListBox List1 
      Height          =   9285
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmAbilitaSoci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IdentificativoSocio As Long

Private Sub cmdConferma_Click()
    SALVA_CONFIGURAZIONE
End Sub

Private Sub Form_Activate()
    GET_LISTA_SOCI
End Sub

Private Sub Form_Load()
    GetParametroSocio
End Sub

Private Sub GET_LISTA_SOCI()
On Error GoTo ERR_GET_LISTA_SOCI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim i As Integer

Screen.MousePointer = 11
DoEvents

Me.List1.Clear

sSQL = "SELECT IDAnagrafica, Anagrafica, Nome, Codice "
sSQL = sSQL & "FROM IERepFornitore "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDCategoriaAnagrafica=" & IdentificativoSocio
sSQL = sSQL & " ORDER BY Anagrafica, Nome"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Me.List1.AddItem fnNotNull(rs!Anagrafica) & " " & fnNotNull(rs!Nome)
    Me.List1.ItemData(i) = fnNotNullN(rs!IDAnagrafica)
    Me.List1.Selected(i) = GET_SOCIO_ABILITATO(fnNotNullN(rs!IDAnagrafica), frmMain.CDCliente.KeyFieldID)
    i = i + 1
    DoEvents
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Screen.MousePointer = 0
Exit Sub
ERR_GET_LISTA_SOCI:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "GET_LISTA_SOCI"
End Sub
Private Sub GetParametroSocio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    IdentificativoSocio = rs!IDCategoriaAnagrafica
Else
    IdentificativoSocio = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_SOCIO_ABILITATO(idsocio As Long, idcliente As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
GET_SOCIO_ABILITATO = False
sSQL = "SELECT * FROM RV_POConfigurazioneClienteSoci "
sSQL = sSQL & "WHERE IDAnagraficaCliente=" & idcliente
sSQL = sSQL & " AND IDAnagraficaSocio=" & idsocio

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SOCIO_ABILITATO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub SALVA_CONFIGURAZIONE()
On Error GoTo ERR_CONFERMA
Dim i As Long
Dim sSQL As String

Screen.MousePointer = 11
DoEvents
sSQL = "DELETE RV_POConfigurazioneClienteSoci "
sSQL = sSQL & "WHERE IDAnagraficaCliente=" & frmMain.CDCliente.KeyFieldID
Cn.Execute sSQL

For i = 0 To Me.List1.ListCount - 1
    If Me.List1.Selected(i) = True Then
        sSQL = "INSERT INTO RV_POConfigurazioneClienteSoci ("
        sSQL = sSQL & "IDAnagraficaSocio, IDAnagraficaCliente) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Me.List1.ItemData(i) & ", "
        sSQL = sSQL & frmMain.CDCliente.KeyFieldID
        sSQL = sSQL & ")"
        Cn.Execute sSQL
        DoEvents
    End If
Next
Screen.MousePointer = 0
Unload Me

Exit Sub

ERR_CONFERMA:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "SALVA_CONFIGURAZIONE"

End Sub
