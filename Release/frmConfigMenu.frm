VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form frmConfigMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurazione del menù"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin DMTDataCmb.DMTCombo cboFiliale 
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DMTDataCmb.DMTCombo cboFilialeMenu 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Configurazione menù "
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Filiale da configurare"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblInfoDettaglio 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Filiale da dove prelevare il menù"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
End
Attribute VB_Name = "frmConfigMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboFiliale_Click()
'    If Me.cboFiliale.CurrentID = 0 Then Exit Sub
'    'Filiale
'    With Me.cboFilialeMenu
'        Set .Database = CnDMT
'        .AddFieldKey "IDFiliale"
'        .DisplayField = "Filiale"
'        .Sql = "SELECT * FROM Filiale "
'        .Sql = .Sql & "WHERE IDAttivitaAzienda=" & GET_LINK_ATTIVITA_AZIENDA(VarIDAzienda)
'        .Sql = .Sql & " AND IDFiliale<>" & Me.cboFiliale.CurrentID
'        .Fill
'    End With
End Sub

Private Sub Command1_Click()
'On Error GoTo ERR_Command1_Click

If Me.cboFiliale.CurrentID = 0 Then
    MsgBox "Inserire la filiale da configurare", vbInformation, "Configurazione menù"
    Exit Sub
End If

If Me.cboFilialeMenu.CurrentID = 0 Then
    MsgBox "Inserire la filiale da cui configurare il menù", vbInformation, "Configurazione menù"
    Exit Sub
End If

If GET_ESISTENZA_MENU_POSEIDON(Me.cboFilialeMenu.CurrentID) = False Then
    MsgBox "La filiale impostata non ha configurato un menù di GreenTop", vbInformation, "Configurazione menù"
    Exit Sub
End If

CREAZIONE_MENU Me.cboFiliale.CurrentID, Me.cboFilialeMenu.CurrentID

CREAZIONE_FUNZIONALITA Me.cboFiliale.CurrentID, Me.cboFilialeMenu.CurrentID

MsgBox "Creazione e configurazione del menù avvenuta con successo", vbInformation, "Configurazione menù"

Unload Me

Exit Sub
ERR_Command1_Click:
    MsgBox Err.Description, vbCritical, "Command1_Click"
    
    Me.lblInfo.Caption = "OPERAZIONE NON COMPLETATA"
    Me.lblInfoDettaglio.Caption = ""

End Sub
Private Function GET_LINK_VOCE_MENU_PADRE(DescrizionePadreMenu As String, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDVoceMenu FROM VoceMenu "
sSQL = sSQL & " WHERE Descrizione=" & fnNormString(DescrizionePadreMenu)
sSQL = sSQL & " AND IDFiliale=" & IDFiliale
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_VOCE_MENU_PADRE = 0
Else
    GET_LINK_VOCE_MENU_PADRE = fnNotNullN(rs!IDVoceMenu)
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_VOCE_MENU_PADRE_FUNZIONE(DescrizionePadreMenu As String, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDVoceMenu FROM VoceMenu "
sSQL = sSQL & " WHERE Descrizione=" & fnNormString(DescrizionePadreMenu)
sSQL = sSQL & " AND IDFiliale=" & IDFiliale
sSQL = sSQL & " AND IDFunzione IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_VOCE_MENU_PADRE_FUNZIONE = 0
Else
    GET_LINK_VOCE_MENU_PADRE_FUNZIONE = fnNotNullN(rs!IDVoceMenu)
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Sub INIT_CONTROLLI()

    'Filiale
    With Me.cboFiliale
        Set .Database = CnDMT
        .AddFieldKey "IDFiliale"
        .DisplayField = "Filiale"
        .Sql = "SELECT * FROM Filiale "
        .Sql = .Sql & "WHERE IDAttivitaAzienda=" & GET_LINK_ATTIVITA_AZIENDA(VarIDAzienda)
        .Fill
    End With
    
    With Me.cboFilialeMenu
        Set .Database = CnDMT
        .AddFieldKey "IDFiliale"
        .DisplayField = "Filiale"
        .Sql = "SELECT * FROM Filiale "
        '.Sql = .Sql & "WHERE IDAttivitaAzienda=" & GET_LINK_ATTIVITA_AZIENDA(VarIDAzienda)
        .Sql = .Sql & " WHERE IDFiliale<>" & Me.cboFiliale.CurrentID
        .Fill
    End With
    
End Sub
Private Function GET_LINK_ATTIVITA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda FROM AttivitaAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_MENU_POSEIDON(IDFiliale As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM VoceMenu "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND Descrizione LIKE " & fnNormString("GreenTop%")

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_MENU_POSEIDON = False
Else
    GET_ESISTENZA_MENU_POSEIDON = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREAZIONE_MENU(IDFiliale As Long, IDFilialeConfig As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim ArrayDescrizione() As String
Dim DescrizioneMenuPadre As String
Dim I As Long
Dim Unita_Progresso As Double
Dim NumeroRecord As Long

sSQL = "SELECT COUNT(IDVoceMenu) AS NumeroRecord"
sSQL = sSQL & " FROM VoceMenu "
sSQL = sSQL & "WHERE IDFiliale=" & Me.cboFilialeMenu.CurrentID
sSQL = sSQL & " AND Descrizione LIKE " & fnNormString("GreenTop%")
sSQL = sSQL & " AND IDFunzione IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing

If NumeroRecord = 0 Then Exit Sub

Me.ProgressBar1.Value = 0
Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)

''''ELIMINAZIONE DEL MENU'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM VoceMenu "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND Descrizione LIKE " & fnNormString("GreenTop%")
CnDMT.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''APERTURA RECORDSET PER SCRIVERE IL MENU'''''''''''''''''''''''''''''''''''''''''''''
Set rsNew = New ADODB.Recordset
sSQL = "SELECT * FROM VoceMenu "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


sSQL = "SELECT * FROM VoceMenu "
sSQL = sSQL & "WHERE IDFiliale=" & IDFilialeConfig
sSQL = sSQL & " AND Descrizione LIKE " & fnNormString("GreenTop%")
sSQL = sSQL & " AND IDFunzione IS NULL"
sSQL = sSQL & " ORDER BY IDVoceMenuPadre"
Set rs = CnDMT.OpenResultset(sSQL)



While Not rs.EOF
    Me.lblInfo.Caption = "Configurazione menù"
    Me.lblInfoDettaglio.Caption = fnNotNull(rs!Descrizione)
    DoEvents
    DescrizioneMenuPadre = ""
    ArrayDescrizione = Split(fnNotNull(rs!Descrizione), "\")
    For I = 1 To UBound(ArrayDescrizione) - 1
        If I > 0 Then
            DescrizioneMenuPadre = DescrizioneMenuPadre & "\"
        End If
        DescrizioneMenuPadre = DescrizioneMenuPadre & ArrayDescrizione(I)
    Next
    rsNew.AddNew
        rsNew!IDVoceMenu = fnGetNewKey("VoceMenu", "IDVoceMenu")
        If Len(DescrizioneMenuPadre) > 0 Then
            rsNew!IDVoceMenuPadre = GET_LINK_VOCE_MENU_PADRE("GreenTop Suite" & DescrizioneMenuPadre, IDFiliale)
        End If
        rsNew!IDFiliale = IDFiliale
        rsNew!VoceMenu = fnNotNull(rs!VoceMenu)
        rsNew!Progressivo = fnNotNullN(rs!Progressivo)
        rsNew!Descrizione = fnNotNull(rs!Descrizione)
    rsNew.Update
    
    If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
    End If
    DoEvents
rs.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub CREAZIONE_FUNZIONALITA(IDFiliale As Long, IDFilialeConfig As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim ArrayDescrizione() As String
Dim DescrizioneMenuPadre As String
Dim I As Long
Dim Unita_Progresso As Double
Dim NumeroRecord As Long


sSQL = "SELECT COUNT(IDVoceMenu) AS NumeroRecord"
sSQL = sSQL & " FROM VoceMenu "
sSQL = sSQL & "WHERE IDFiliale=" & IDFilialeConfig
sSQL = sSQL & " AND Descrizione LIKE " & fnNormString("GreenTop%")
sSQL = sSQL & " AND IDFunzione>0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing

If NumeroRecord = 0 Then Exit Sub

Me.ProgressBar1.Value = 0
Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)

''''APERTURA RECORDSET PER SCRIVERE IL MENU'''''''''''''''''''''''''''''''''''''''''''''
Set rsNew = New ADODB.Recordset
sSQL = "SELECT * FROM VoceMenu "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''CREAZIONE FUNZIONALITA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM VoceMenu "
sSQL = sSQL & "WHERE IDFiliale=" & IDFilialeConfig
sSQL = sSQL & " AND Descrizione LIKE " & fnNormString("GreenTop%")
sSQL = sSQL & " AND IDFunzione>0"
sSQL = sSQL & " ORDER BY IDVoceMenuPadre"
Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Me.lblInfo.Caption = "Configurazione funzionalità nel menù"
    Me.lblInfoDettaglio.Caption = GET_FUNZIONE(fnNotNullN(rs!IDFunzione))
    DoEvents
    rsNew.AddNew
        rsNew!IDVoceMenu = fnGetNewKey("VoceMenu", "IDVoceMenu")
        rsNew!IDVoceMenuPadre = GET_LINK_VOCE_MENU_PADRE_FUNZIONE(fnNotNull(rs!Descrizione), IDFiliale)
        rsNew!IDFunzione = fnNotNullN(rs!IDFunzione)
        rsNew!IDFiliale = IDFiliale
        rsNew!VoceMenu = fnNotNull(rs!VoceMenu)
        rsNew!Progressivo = fnNotNullN(rs!Progressivo)
        rsNew!Descrizione = fnNotNull(rs!Descrizione)
        rsNew!IdentificativoIconaRisorsa = fnNotNullN(rs!IdentificativoIconaRisorsa)
    rsNew.Update

    If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
    End If
    DoEvents
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
End Sub
Private Function GET_FUNZIONE(IDFunzione As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione FROM Funzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE = ""
Else
    GET_FUNZIONE = fnNotNull(rs!Funzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Form_Load()
    INIT_CONTROLLI
End Sub
