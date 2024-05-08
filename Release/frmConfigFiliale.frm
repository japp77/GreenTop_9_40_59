VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Begin VB.Form frmConfigFiliale 
   Caption         =   "Configurazione parametri filiale"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9555
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigFiliale.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9555
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraMagazzini 
      Caption         =   "Magazzini"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   4335
      Begin DMTDataCmb.DMTCombo cboMagLav 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
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
      Begin DMTDataCmb.DMTCombo cboMagVend 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
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
      Begin VB.Label Label1 
         Caption         =   "Magazzino principale o di conferimento"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Magazzino di trasformazione e vendita"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label lblNuovoMagazzino 
         Height          =   255
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "CONFERMA"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   4335
   End
   Begin VB.TextBox txtControllo 
      Height          =   6615
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   4935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin DMTDataCmb.DMTCombo cboFiliale 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
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
   Begin VB.Label Label1 
      Caption         =   "Filiale"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Width           =   4335
   End
End
Attribute VB_Name = "frmConfigFiliale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private cnConfig As ADODB.Connection
Private raParametri As ADODB.Recordset
Private LINK_PARAMETRO_FILIALE As Long
Private LINK_PARAMETRO_FILIALE_AZIENDA As Long
Private LINK_PARAMETRO_FILIALE_FARMACIA As Long
Private LINK_PARAMETRO_FILIALE_LIQ As Long

Private LINK_CLIENTE_ORDINE_PRED As Long

Private NUMERO_DOCUMENTO_ORDINE As Long

Private LINK_CLIENTE_LAV_IVGAMMA As Long
Private LINK_ORDINE_LAV_IV_GAMMA As Long
Private LINK_ORDINE_MERCE_GIACENZA As Long

Private LINK_TIPO_PEDANA_PREDEFINITA As Long

Private LINK_DEMO_CLIENTE As Long
Private LINK_DEMO_SOCIO As Long
Private LINK_DEMO_FORNITORE As Long

Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim sSQL As String
Dim NumeroOrdineCliente As String

If Me.cboFiliale.CurrentID = 0 Then Exit Sub
If Me.cboMagLav.CurrentID = 0 Then Exit Sub
If Me.cboMagVend.CurrentID = 0 Then Exit Sub

If (Me.cboMagLav.CurrentID > 0) And (Me.cboMagVend.CurrentID > 0) Then
    CREAZIONE_CAUSALI_MAGAZZINO
End If

CREAZIONE_TIPO_PRODOTTO VarIDAzienda
CREAZIONE_GRUPPO_EQUIVALENZA_ARTICOLO VarIDAzienda
CREAZIONE_SEZIONALI VarIDFiliale, VarIDAzienda
CREAZIONE_CATEGORIA_ANAGRAFICA_SOCIO

UNITA_DI_MISURA_COOP
CICLO_BIOLOGICO
STATO_LOTTO
TIPO_PRODUZIONE
TITOLO_TERRENO
STRUTTURA_SERRA
TIPO_OPERAZIONE
IDENTIFICATORI_EAN128

INSERIMENTO_PEDANA VarIDAzienda, VarIDFiliale
INSERIMENTO_IMBALLO VarIDAzienda, VarIDFiliale
INSERIMENTO_ARTICOLO VarIDAzienda, VarIDFiliale

LINK_DEMO_CLIENTE = GET_LINK_ANAGRAFICA_CLIENTE_DEMO("Cliente demo", VarIDAzienda)
LINK_DEMO_FORNITORE = GET_LINK_ANAGRAFICA_FORNITORE_DEMO("Fornitore demo", VarIDAzienda)
LINK_DEMO_SOCIO = GET_LINK_ANAGRAFICA_SOCIO_DEMO("Socio demo", VarIDAzienda)

'ORDINE MERCE IN GIACENZA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LINK_CLIENTE_ORDINE_PRED = GET_LINK_ANAGRAFICA_CLIENTE("Merce in giacenza", VarIDAzienda)
LINK_CLIENTE_LAV_IVGAMMA = GET_LINK_ANAGRAFICA_CLIENTE("Lavorazione in IV Gamma", VarIDAzienda)

If GET_ESISTENZA_ORDINE_MERCE_IN_GIACENZA(LINK_CLIENTE_ORDINE_PRED, VarIDAzienda, Me.cboFiliale.CurrentID) = False Then
    NUMERO_DOCUMENTO_ORDINE = fnNotNullN(FN_CREA_ORDINE(LINK_CLIENTE_ORDINE_PRED, VarIDAzienda, Me.cboFiliale.CurrentID))
End If

If GET_ESISTENZA_ORDINE_MERCE_IN_GIACENZA(LINK_CLIENTE_LAV_IVGAMMA, VarIDAzienda, Me.cboFiliale.CurrentID) = False Then
    LINK_ORDINE_LAV_IV_GAMMA = fnNotNullN(FN_CREA_ORDINE_LAV_IVGAMMA(LINK_CLIENTE_LAV_IVGAMMA, VarIDAzienda, Me.cboFiliale.CurrentID))
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

LINK_PARAMETRO_FILIALE = GET_LINK_PARAMETRO_FILIALE(Me.cboFiliale.CurrentID, VarIDAzienda)

If LINK_PARAMETRO_FILIALE = 0 Then Exit Sub

CREAZIONE_SEZIONALI_PER_DOCUMENTI_COOP LINK_PARAMETRO_FILIALE, VarIDAzienda, Me.cboFiliale.CurrentID
CREAZIONE_PROCESSI_PER_DOCUMENTI_COOP LINK_PARAMETRO_FILIALE, VarIDAzienda, Me.cboFiliale.CurrentID

LINK_PARAMETRO_FILIALE_AZIENDA = GET_LINK_PARAMETRO_FILIALE_AZIENDA(Me.cboFiliale.CurrentID, VarIDAzienda)

LINK_PARAMETRO_FILIALE_FARMACIA = GET_LINK_PARAMETRO_FILIALE_FARMACIA(Me.cboFiliale.CurrentID, VarIDAzienda)


LINK_PARAMETRO_FILIALE_LIQ = GET_LINK_PARAMETRO_FILIALE_LIQUIDAZIONE(Me.cboFiliale.CurrentID, VarIDAzienda)

ELIMINA_FORMULA_QUANTITA

Me.lblInfo.Caption = "OPERAZIONE COMPLETATA"

MsgBox "OPERAZIONE COMPLETATA", vbInformation, "Configurazione iniziale della filiale"
Unload Me
Exit Sub

ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdConferma_Click"
End Sub
Private Function CONNESSIONE_CONFIGURAZIONE() As Boolean
On Error GoTo ERR_CONNESSIONE_CONFIGURAZIONE

    CONNESSIONE_CONFIGURAZIONE = False
    
    If Not (cnConfig Is Nothing) Then
        cnConfig.Close
        Set cnConfig = Nothing
    End If

    Set cnConfig = New ADODB.Connection
    
    cnConfig.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0;Data source=" & App.Path & "\ConfigGreenTop.mdb"
    cnConfig.Open

    CONNESSIONE_CONFIGURAZIONE = True
    
Exit Function

ERR_CONNESSIONE_CONFIGURAZIONE:
    MsgBox Err.Description, vbCritical, "CONNESSIONE_CONFIGURAZIONE"
    CONNESSIONE_CONFIGURAZIONE = False
End Function
Private Sub Form_Load()
    If CONNESSIONE_CONFIGURAZIONE = False Then
        Me.cmdConferma.Enabled = False
        Exit Sub
    End If
    
    INIT_CONTROLLI

End Sub
Private Sub INIT_CONTROLLI()
    'Filiale
    With Me.cboFiliale
        Set .Database = CnDMT
        .AddFieldKey "IDFiliale"
        .DisplayField = "Filiale"
        .Sql = "SELECT * FROM Filiale WHERE IDAttivitaAzienda=" & GET_LINK_ATTIVITA_AZIENDA(VarIDAzienda)
        .Fill
    End With
    
    'Magazzino di lavorazione
    With Me.cboMagLav
        Set .Database = CnDMT
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .Sql = "SELECT * FROM Magazzino WHERE IDAzienda=" & VarIDAzienda
        .Fill
    End With
    
    'Magazzino di vendita
    With Me.cboMagVend
        Set .Database = CnDMT
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .Sql = "SELECT * FROM Magazzino WHERE IDAzienda=" & VarIDAzienda
        .Fill
    End With
    
    Me.cboFiliale.WriteOn VarIDFiliale
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
Private Sub CREAZIONE_CAUSALI_MAGAZZINO()
On Error GoTo ERR_CREAZIONE_CAUSALI_MAGAZZINO
Dim sSQL As String
Dim rsConfig As ADODB.Recordset
Dim Link_Funzione As Long

Me.lblInfo.Caption = "CREAZIONE CAUSALE DI MAGAZZINO"
Me.txtControllo.Text = Me.txtControllo.Text & "CAUSALI DI MAGAZZINO" & vbCrLf


sSQL = "SELECT * FROM Funzione "

Set rsConfig = New ADODB.Recordset

rsConfig.Open sSQL, cnConfig

While Not rsConfig.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsConfig!Funzione) & vbCrLf
    DoEvents
    GET_LINK_FUNZIONE rsConfig!Funzione, rsConfig!IDFunzione
    
rsConfig.MoveNext
Wend

rsConfig.Close
Set rsConfig = Nothing
Exit Sub
ERR_CREAZIONE_CAUSALI_MAGAZZINO:
    MsgBox Err.Description, vbCritical, "CREAZIONE_CAUSALI_MAGAZZINO"
    Me.txtControllo.Text = Me.txtControllo.Text & " (" & Err.Description & ")" & vbCrLf
End Sub

Private Function GET_ESISTENZA_FUNZIONE(Funzione As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE Funzione=" & fnNormString(Funzione)

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_ESISTENZA_FUNZIONE = 0
Else
    GET_ESISTENZA_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If


rs.Close
Set rs = Nothing
End Function
Private Function GET_LINK_FUNZIONE(Funzione As String, IDFunzioneConfig As Long)
Dim sSQL As String
Dim Link_Funzione As Long
Dim Link_processo_per_funzione As Long

Link_Funzione = GET_ESISTENZA_FUNZIONE(Funzione)

If Link_Funzione = 0 Then
    'CREAZIONE FUNZIONE
    Link_Funzione = fnGetNewKeyPerTipoOggetto("Funzione", "IDFunzione")
    
    sSQL = "INSERT INTO Funzione (IDFunzione, Funzione, IDTipoOggetto, "
    sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & Link_Funzione & ", "
    sSQL = sSQL & fnNormString(Funzione) & ", "
    sSQL = sSQL & 9 & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & 0 & ")"
    
    CnDMT.Execute sSQL
End If
Link_processo_per_funzione = CREA_PROCESSO_PER_FUNZIONE(Link_Funzione, IDFunzioneConfig)



End Function
Private Function GET_LINK_CAUSALE_MAGAZZINO_COMM(NomeCampo As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Funzione As String

sSQL = "SELECT " & NomeCampo & " FROM ConfigMovChiusuraComm "

Set rs = New ADODB.Recordset

rs.Open sSQL, cnConfig

If rs.EOF Then
    Funzione = ""
Else
    Funzione = fnNotNull(rs.Fields(NomeCampo).Value)
End If

rs.Close
Set rs = Nothing

If Len(Trim(Funzione)) = 0 Then
    GET_LINK_CAUSALE_MAGAZZINO_COMM = 0
    Exit Function
End If

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE Funzione=" & fnNormString(Funzione)

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_LINK_CAUSALE_MAGAZZINO_COMM = 0
Else
    GET_LINK_CAUSALE_MAGAZZINO_COMM = fnNotNullN(rs!IDFunzione)
End If
rs.Close
Set rs = Nothing
End Function

Private Function GET_LINK_FUNZIONE_MOV(Funzione As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE Funzione=" & fnNormString(Funzione)

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_LINK_FUNZIONE_MOV = 0
Else
    GET_LINK_FUNZIONE_MOV = fnNotNullN(rs!IDFunzione)
End If
rs.Close
Set rs = Nothing
End Function

Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If Not rs.EOF Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function GET_ESISTENZA_CONTATORE_PER_PROCESSO(IDContatore As Long, IDProcessoFunzione As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ContatorePerProcesso "
sSQL = sSQL & "WHERE IDContatoreArticolo=" & IDContatore
sSQL = sSQL & " AND IDProcessoPerFunzione=" & IDProcessoFunzione
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_CONTATORE_PER_PROCESSO = False
Else
    GET_ESISTENZA_CONTATORE_PER_PROCESSO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function CREA_PROCESSO_PER_FUNZIONE(IDFunzione As Long, IDFunzioneConfig As Long) As Long
Dim sSQL As String
Dim IDLocal As Long
Dim rsConfig As ADODB.Recordset

sSQL = "SELECT * FROM ProcessoPerFunzione "
sSQL = sSQL & " WHERE IDFunzione=" & IDFunzioneConfig
sSQL = sSQL & " ORDER BY Sequenza "

Set rsConfig = New ADODB.Recordset

rsConfig.Open sSQL, cnConfig

While Not rsConfig.EOF
    IDLocal = GET_LINK_PROCESSO_FUNZIONE(rsConfig!IDProcesso, IDFunzione)
    
    If IDLocal = 0 Then
        IDLocal = fnGetNewKeyPerTipoOggetto("ProcessoPerFunzione", "IDProcessoPerFunzione")
        
        sSQL = "INSERT INTO ProcessoPerFunzione ("
        sSQL = sSQL & "IDProcessoPerFunzione, IDFunzione, IDProcesso, Sequenza, "
        sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDLocal & ", "
        sSQL = sSQL & IDFunzione & ", "
        sSQL = sSQL & rsConfig!IDProcesso & ", "
        sSQL = sSQL & rsConfig!Sequenza & ", "
        sSQL = sSQL & fnNormDate(Date) & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & 0 & ")"
        
        CnDMT.Execute sSQL
    End If
    
    If ((fnNotNullN(rsConfig!IDProcesso) = 49) Or (fnNotNullN(rsConfig!IDProcesso) = 63)) Then
        CREA_PROCESSO_PER_FUNZIONE = IDLocal
        CREA_CONTATORI_PER_PROCESSO fnNotNullN(rsConfig!IDProcessoPerFunzione), IDLocal
    End If

rsConfig.MoveNext
Wend

rsConfig.Close
Set rsConfig = Nothing
End Function
Private Function GET_LINK_PROCESSO_FUNZIONE(IDProcesso As Long, IDFunzione As Long) As Long

Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDProcessoPerFunzione FROM ProcessoPerFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione
sSQL = sSQL & " AND IDProcesso=" & IDProcesso

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_LINK_PROCESSO_FUNZIONE = 0
Else
    GET_LINK_PROCESSO_FUNZIONE = fnNotNullN(rs!IDProcessoPerFunzione)
End If


rs.Close
Set rs = Nothing
End Function
Private Function GET_LINK_CONTATORE_ARTICOLO(ContatoreArticolo As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDContatoreArticolo FROM ContatoreArticolo "
sSQL = sSQL & "WHERE ContatoreArticolo=" & fnNormString(ContatoreArticolo)


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_LINK_CONTATORE_ARTICOLO = 0
Else
    GET_LINK_CONTATORE_ARTICOLO = fnNotNullN(rs!IDContatoreArticolo)
End If


rs.Close
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_CONTATORE_MAGAZZINO(IDContatoreArticolo As Long, IDMagazzino As Long) As Boolean
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM ContatoreArticoloPerMagazzino "
sSQL = sSQL & "WHERE IDContatoreArticolo=" & IDContatoreArticolo
sSQL = sSQL & " AND IDMagazzino=" & IDMagazzino

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF Then
    GET_ESISTENZA_CONTATORE_MAGAZZINO = False
Else
    GET_ESISTENZA_CONTATORE_MAGAZZINO = True
End If


rs.Close
Set rs = Nothing

End Function

Private Sub CREA_CONTATORI_PER_PROCESSO(IDProcessoPerFunzioneConfig As Long, IDProcessoPerFunzione As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim Link_Contatore As Long
Dim rsCont As ADODB.Recordset

sSQL = "SELECT * FROM ContatorePerProcesso "
sSQL = sSQL & "WHERE IDProcessoPerFunzione=" & IDProcessoPerFunzioneConfig

Set rs = New ADODB.Recordset
rs.Open sSQL, cnConfig

If Not rs.EOF Then
    sSQL = "SELECT * FROM ContatoreArticolo "
    sSQL = sSQL & " WHERE IDContatoreArticolo=" & fnNotNullN(rs!IDContatoreArticolo)
    
    Set rsCont = New ADODB.Recordset
    rsCont.Open sSQL, cnConfig
    
    If Not rsCont.EOF Then
        Link_Contatore = GET_LINK_CONTATORE_ARTICOLO(fnNotNull(rsCont!ContatoreArticolo))
        'CREAZIONE CONTATORE ARTICOLO
        If Link_Contatore = 0 Then
            Link_Contatore = fnGetNewKey("ContatoreArticolo", "IDContatoreArticolo")
            
            sSQL = "INSERT INTO ContatoreArticolo (IDContatoreArticolo, ContatoreArticolo, "
            sSQL = sSQL & "PartecipaPrezzoMedio, PartecipaCostoMedio, VariaCostoUltimo, VariaCostoPrecedente, Predefinito) "
            sSQL = sSQL & "VALUES ("
            sSQL = sSQL & Link_Contatore & ", "
            sSQL = sSQL & fnNormString(rsCont!ContatoreArticolo) & ", "
            sSQL = sSQL & fnNormString(rsCont!PartecipaPrezzoMedio) & ", "
            sSQL = sSQL & fnNormString(rsCont!PartecipaCostoMedio) & ", "
            sSQL = sSQL & fnNormBoolean(rsCont!VariaCostoUltimo) & ", "
            sSQL = sSQL & fnNormBoolean(rsCont!VariaCostoPrecedente) & ", "
            sSQL = sSQL & fnNormBoolean(rsCont!Predefinito) & ")"
            CnDMT.Execute sSQL
         End If
             If Link_Contatore = 0 Then
                'CREAZIONE CONTATORE PER PROCESSO
                sSQL = "INSERT INTO ContatorePerProcesso (IDContatoreArticolo, IDProcessoPerFunzione, "
                sSQL = sSQL & "Numero, Quantita, Valore, DataVariazione, "
                sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
                sSQL = sSQL & "VALUES ("
                sSQL = sSQL & Link_Contatore & ", "
                sSQL = sSQL & IDProcessoPerFunzione & ", "
                sSQL = sSQL & fnNormString(rs!Numero) & ", "
                sSQL = sSQL & fnNormString(rs!Quantita) & ", "
                sSQL = sSQL & fnNormString(rs!Valore) & ", "
                sSQL = sSQL & fnNormBoolean(1) & ", "
                sSQL = sSQL & fnNormDate(Date) & ", "
                sSQL = sSQL & 1 & ", "
                sSQL = sSQL & 0 & ")"
                CnDMT.Execute sSQL
            Else
                If GET_ESISTENZA_CONTATORE_PER_PROCESSO(Link_Contatore, IDProcessoPerFunzione) = False Then
                    'CREAZIONE CONTATORE PER PROCESSO
                    sSQL = "INSERT INTO ContatorePerProcesso (IDContatoreArticolo, IDProcessoPerFunzione, "
                    sSQL = sSQL & "Numero, Quantita, Valore, DataVariazione, "
                    sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
                    sSQL = sSQL & "VALUES ("
                    sSQL = sSQL & Link_Contatore & ", "
                    sSQL = sSQL & IDProcessoPerFunzione & ", "
                    sSQL = sSQL & fnNormString(rs!Numero) & ", "
                    sSQL = sSQL & fnNormString(rs!Quantita) & ", "
                    sSQL = sSQL & fnNormString(rs!Valore) & ", "
                    sSQL = sSQL & fnNormBoolean(1) & ", "
                    sSQL = sSQL & fnNormDate(Date) & ", "
                    sSQL = sSQL & 1 & ", "
                    sSQL = sSQL & 0 & ")"
                    CnDMT.Execute sSQL
                End If
            End If

        CREA_CONTATORE_PER_MAGAZZINO Link_Contatore, rsCont!IDContatoreArticolo
    End If
    
    rsCont.Close
    Set rsCont = Nothing
End If

rs.Close
Set rs = Nothing
End Sub
Private Sub CREA_CONTATORE_PER_MAGAZZINO(IDContatore As Long, IDContatoreConfig As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM ContatoreArticoloPerMagazzino "
sSQL = sSQL & "WHERE IDContatoreArticolo=" & IDContatoreConfig

Set rs = New ADODB.Recordset

rs.Open sSQL, cnConfig

If Not rs.EOF Then
    If GET_ESISTENZA_CONTATORE_MAGAZZINO(IDContatore, Me.cboMagLav.CurrentID) = False Then
        sSQL = "INSERT INTO ContatoreArticoloPerMagazzino("
        sSQL = sSQL & "IDMagazzino, IDContatoreArticolo, PartecipaGiacenza, PartecipaDisponibilita, "
        sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Me.cboMagLav.CurrentID & ", "
        sSQL = sSQL & IDContatore & ", "
        sSQL = sSQL & fnNormString(rs!PartecipaGiacenza) & ", "
        sSQL = sSQL & fnNormString(rs!PartecipaDisponibilita) & ", "
        sSQL = sSQL & fnNormDate(Date) & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & 0
        sSQL = sSQL & ")"
        CnDMT.Execute sSQL
    End If
    
    If GET_ESISTENZA_CONTATORE_MAGAZZINO(IDContatore, Me.cboMagVend.CurrentID) = False Then
        sSQL = "INSERT INTO ContatoreArticoloPerMagazzino("
        sSQL = sSQL & "IDMagazzino, IDContatoreArticolo, PartecipaGiacenza, PartecipaDisponibilita, "
        sSQL = sSQL & "DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Me.cboMagVend.CurrentID & ", "
        sSQL = sSQL & IDContatore & ", "
        sSQL = sSQL & fnNormString(rs!PartecipaGiacenza) & ", "
        sSQL = sSQL & fnNormString(rs!PartecipaDisponibilita) & ", "
        sSQL = sSQL & fnNormDate(Date) & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & 0
        sSQL = sSQL & ")"
        CnDMT.Execute sSQL
    End If
End If

rs.Close
Set rs = Nothing
End Sub
Private Sub CREAZIONE_TIPO_PRODOTTO(IDAzienda As Long)
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "TIPO PRODOTTO"
Me.txtControllo.Text = Me.txtControllo.Text & "TIPO PRODOTTO" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM TipoProdotto "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM TipoProdotto "
sSQL = sSQL & " WHERE IDAzienda=" & IDAzienda
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!TipoProdotto) & vbCrLf
    DoEvents
    If GET_ESISTENZA_TIPO_PRODOTTO(fnNotNull(rsArc!TipoProdotto), IDAzienda) = False Then
        rsNew.AddNew
            rsNew!IDTipoProdotto = fnGetNewKey("TipoProdotto", "IDTipoProdotto")
            rsNew!IDAzienda = IDAzienda
            rsNew!TipoProdotto = fnNotNull(rsArc!TipoProdotto)
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_TIPO_PRODOTTO(TipoProdotto As String, IDAzienda As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoProdotto FROM TipoProdotto "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND TipoProdotto=" & fnNormString(TipoProdotto)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_TIPO_PRODOTTO = False
Else
    GET_ESISTENZA_TIPO_PRODOTTO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREAZIONE_CATEGORIA_ANAGRAFICA_SOCIO()
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "CATEGORIA ANAGRAFICA PER SOCIO"
Me.txtControllo.Text = Me.txtControllo.Text & "CATEGORIA ANAGRAFICA PER SOCIO" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM CategoriaAnagrafica "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM CategoriaAnagrafica "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!CategoriaAnagrafica) & vbCrLf
    DoEvents
    If GET_ESISTENZA_CATEGORIA_ANAGRAFICA_SOCIO(fnNotNull(rsArc!CategoriaAnagrafica)) = False Then
        rsNew.AddNew
            rsNew!IDCategoriaAnagrafica = fnGetNewKey("CategoriaAnagrafica", "IDCategoriaAnagrafica")
            rsNew!CategoriaAnagrafica = fnNotNull(rsArc!CategoriaAnagrafica)
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_CATEGORIA_ANAGRAFICA_SOCIO(CategoriaAnagrafica As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM CategoriaAnagrafica "
sSQL = sSQL & " WHERE CategoriaAnagrafica=" & fnNormString(CategoriaAnagrafica)
sSQL = sSQL & " AND IDTipoAnagrafica IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_CATEGORIA_ANAGRAFICA_SOCIO = False
Else
    GET_ESISTENZA_CATEGORIA_ANAGRAFICA_SOCIO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function


Private Sub CREAZIONE_GRUPPO_EQUIVALENZA_ARTICOLO(IDAzienda As Long)
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset


Me.lblInfo.Caption = "GRUPPO EQUIVALENZA ARTICOLO"
Me.txtControllo.Text = Me.txtControllo.Text & "GRUPPO EQUIVALENZA ARTICOLO" & vbCrLf
DoEvents
''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM GruppoEquivalenzaArticolo "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM GruppoEquivalenzaArticolo "
sSQL = sSQL & " WHERE IDAzienda=" & IDAzienda
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!GruppoEquivalenzaArticolo) & vbCrLf
    DoEvents
    If GET_ESISTENZA_GRUPPO_EQUIVALENZA_ARTICOLO(fnNotNull(rsArc!GruppoEquivalenzaArticolo), IDAzienda) = False Then
        rsNew.AddNew
            rsNew!IDGruppoEquivalenzaArticolo = fnGetNewKey("GruppoEquivalenzaArticolo", "IDGruppoEquivalenzaArticolo")
            rsNew!IDAzienda = IDAzienda
            rsNew!GruppoEquivalenzaArticolo = fnNotNull(rsArc!GruppoEquivalenzaArticolo)
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_GRUPPO_EQUIVALENZA_ARTICOLO(GruppoEquivalenzaArticolo As String, IDAzienda As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDGruppoEquivalenzaArticolo FROM GruppoEquivalenzaArticolo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND GruppoEquivalenzaArticolo=" & fnNormString(GruppoEquivalenzaArticolo)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_GRUPPO_EQUIVALENZA_ARTICOLO = False
Else
    GET_ESISTENZA_GRUPPO_EQUIVALENZA_ARTICOLO = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREAZIONE_SEZIONALI(IDFiliale As Long, IDAzienda As Long)
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "SEZIONALI PER DOCUMENTI COOP."
Me.txtControllo.Text = Me.txtControllo.Text & "SEZIONALI PER DOCUMENTI COOP." & vbCrLf
DoEvents

''''''''RECUPERO SEZIONALI ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM Sezionale "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Sezionale "
sSQL = sSQL & " WHERE IDFiliale=" & IDFiliale
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!Sezionale) & vbCrLf
    DoEvents
    If GET_ESISTENZA_SEZIONALE(fnNotNull(rsArc!Sezionale), IDFiliale) = False Then
        rsNew.AddNew
            rsNew!IDSezionale = fnGetNewKey("Sezionale", "IDSezionale")
            rsNew!IDFiliale = VarIDFiliale
            rsNew!IDRegistroIva = fnNotNullN(rsArc!IDRegistroIva)
            rsNew!Sezionale = fnNotNull(rsArc!Sezionale)
            rsNew!DataUltimaVariazione = Date
            rsNew!IDUtenteUltimaVariazione = 1
            rsNew!VirtualDelete = 0
            rsNew!DescrizioneInFattDiff = ""
            rsNew!Prefisso = ""
        rsNew.Update
    End If
rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing

End Sub
Private Function GET_ESISTENZA_SEZIONALE(Sezionale As String, IDFiliale As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND Sezionale=" & fnNormString(Sezionale)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_SEZIONALE = False
Else
    GET_ESISTENZA_SEZIONALE = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ANAGRAFICA_CLIENTE(AnagraficaCliente As String, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim objAna As dmtRegAna.CRegAnagrafica
Dim ris As Long

Me.lblInfo.Caption = "CREAZIONE ANAGRAFICA CLIENTE PER ORDINE DI GIACENZA"
Me.txtControllo.Text = Me.txtControllo.Text & "CREAZIONE ANAGRAFICA CLIENTE PER ORDINE DI GIACENZA" & vbCrLf
DoEvents

''''''''''''Controllo dell'esistenza dell'anagrafica''''''''''''''''''''''''''''''''
sSQL = "SELECT IDAnagrafica FROM IERepCliente "
sSQL = sSQL & " WHERE IDAzienda=" & VarIDAzienda
sSQL = sSQL & " AND Anagrafica=" & fnNormString(AnagraficaCliente)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_CLIENTE = 0
Else
    GET_LINK_ANAGRAFICA_CLIENTE = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LINK_ANAGRAFICA_CLIENTE > 0 Then Exit Function

'''''''CREAZIONE DELL'ANAGRAFICA'''''''''''''''''''''''''''''''

'Crea un'istanza dell'oggetto CRegAnagrafica
Set objAna = New dmtRegAna.CRegAnagrafica

'Assegna la connessione aperta all'oggetto CReganagrafica
objAna.Connection = CnDMT

objAna.Field "Anagrafica", AnagraficaCliente, "Anagrafica" 'Richiesto

objAna.Field "DataUltimaVariazione", Date, "Anagrafica" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Anagrafica" 'Richiesto
objAna.Field "VirtualDelete", 0, "Anagrafica" 'Richiesto

'Valorizzare i campi della tabella CLIENTE, utilizzando il metodo Field
objAna.Field "IDAzienda", IDAzienda, "Cliente" 'Richiesto
objAna.Field "IDTipoAnagrafica", 2, "Cliente" 'Richiesto

objAna.Field "DataUltimaVariazione", Date, "Cliente" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Cliente" 'Richiesto
objAna.Field "VirtualDelete", 0, "Cliente" 'Richiesto

ris = objAna.Insert
If ris = 0 Then
    GET_LINK_ANAGRAFICA_CLIENTE = fnNotNullN(objAna.Field("IDAnagrafica", , "Anagrafica"))
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Private Function CREAZIONE_CLIENTE(NomeCliente As String, IDAzienda As Long)


End Function
Private Function GET_ESISTENZA_ORDINE_MERCE_IN_GIACENZA(IDCliente As Long, IDAzienda As Long, IDFiliale As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDSezionale As Long

IDSezionale = GET_LINK_SEZIONALE_ORDINE(15, IDFiliale)

sSQL = "SELECT Oggetto.IDOggetto "
sSQL = sSQL & "FROM Oggetto INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo000F ON Oggetto.IDOggetto = ValoriOggettoPerTipo000F.IDOggetto AND "
sSQL = sSQL & "Oggetto.IDTipoOggetto = ValoriOggettoPerTipo000F.IDTipoOggetto "
sSQL = sSQL & "WHERE Oggetto.IDTipoOggetto=15 "
sSQL = sSQL & " AND Oggetto.IDAzienda=" & IDAzienda
sSQL = sSQL & " AND ValoriOggettoPerTipo000F.Link_nom_anagrafica=" & IDCliente
sSQL = sSQL & " AND Oggetto.IDSezionale=" & IDSezionale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_ORDINE_MERCE_IN_GIACENZA = False
Else
    GET_ESISTENZA_ORDINE_MERCE_IN_GIACENZA = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function FN_CREA_ORDINE(IDCliente As Long, IDAzienda As Long, IDFiliale As Long) As String
On Error GoTo ERR_FN_CREA_ORDINE
Dim ObjDoc As DmtDocs.cDocument
Dim cDefault As Collection
Dim IDSezionale As Long
Dim sSQL As String

Me.lblInfo.Caption = "CREAZIONE ORDINE DI MERCE IN GIACENZA"
Me.txtControllo.Text = Me.txtControllo.Text & "CREAZIONE ORDINE DI MERCE IN GIACENZA" & vbCrLf
DoEvents

If Not (ObjDoc Is Nothing) Then
    Set ObjDoc = Nothing
End If

Set ObjDoc = New DmtDocs.cDocument
 
    With ObjDoc
    
        Set .Connection = CnDMT
        .IDAzienda = IDAzienda
        .IDAttivitaAzienda = GET_ATTIVITA_AZIENDA(IDAzienda)
        .IDFiliale = IDFiliale
        .SetTipoOggetto 15
        .IDFunzione = 128
        
        .IDEsercizio = fncEsercizio(Date, IDAzienda)
        .IDSezionale = GET_LINK_SEZIONALE_ORDINE(.IDTipoOggetto, .IDFiliale)
        .IDTipoAnagrafica = 2
        .IDUtente = 1
        .Descrizione = "Ordine da cliente"
        .DataEmissione = Date
        .Numero = 0
        If .Tables.Count = 0 Then
        'Se Tables.Count = 0 vuol dire che l'oggetto
        'DmtDocs non è mai stato inizializzato
            .Clear
            .SetTipoOggetto 15
        Else
            .ClearValues
        End If
        
        ObjDoc.Tables("ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
        
        .Field "Link_Doc_magazzino", Me.cboMagVend.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Doc_sezionale", ObjDoc.IDSezionale, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Val_valuta", 9, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_data", ObjDoc.DataEmissione, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_numero", 0, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
        .ReadDataFromCliFo IDCliente
        
        .Field "Link_Doc_pagamento", 1, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_ordine_chiuso", 0, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "RV_POOrdineCompletato", 0, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
    
        Set ObjDoc.Scadenze = Nothing
        ObjDoc.PerformDocument Nothing
        FN_CREA_ORDINE = "0"
        FN_CREA_ORDINE = ObjDoc.Insert
        
        ObjDoc.Update
        
        LINK_ORDINE_MERCE_GIACENZA = ObjDoc.IDOggetto
        
        If (ObjDoc.IDOggetto) > 0 Then
            sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
            sSQL = sSQL & "RV_PONumeroOrdinePadre=" & fnNormNumber(ObjDoc.Numero) & ", "
            sSQL = sSQL & "RV_PODataOrdinePadre=" & fnNormDate(ObjDoc.DataEmissione) & ", "
            sSQL = sSQL & "RV_PONumeroListaPrelievo=1" & ", "
            sSQL = sSQL & "RV_POIDOrdinePadre=" & ObjDoc.IDOggetto
            sSQL = sSQL & " WHERE IDOggetto=" & ObjDoc.IDOggetto
            
            CnDMT.Execute sSQL
        End If
        
    End With
    
    Set ObjDoc = Nothing
   
 Exit Function
ERR_FN_CREA_ORDINE:
    MsgBox Err.Description, vbCritical, "FN_CREA_ORDINE"
    FN_CREA_ORDINE = "0"
    LINK_ORDINE_MERCE_GIACENZA = 0
    
End Function
Public Function GET_ATTIVITA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "SELECT IDAttivitaAzienda "
    sSQL = sSQL & "FROM AttivitaAzienda "
    sSQL = sSQL & " WHERE IDAzienda = " & IDAzienda
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
    Else
        GET_ATTIVITA_AZIENDA = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End Function
Public Function fncEsercizio(dData As String, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDEsercizio, Esercizio FROM Esercizio "
    sSQL = sSQL & " WHERE IDAzienda = " & IDAzienda
    sSQL = sSQL & " AND DataInizio <= " & fnNormDate(dData)
    sSQL = sSQL & " AND DataFine >= " & fnNormDate(dData)
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fncEsercizio = fnNotNullN(rs!IDEsercizio)
    Else
        fncEsercizio = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Function fncTrovaDocumento(NumeroDoc As String, DataOrdine As String, IDAzienda As Long) As Long
On Error GoTo ERR_fncTrovaDocumento
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto From Oggetto "
sSQL = sSQL & " WHERE IDTipoOggetto=15"
sSQL = sSQL & " AND Numero=" & fnNormString(NumeroDoc)
sSQL = sSQL & " AND DataEmissione=" & fnNormDate(DataOrdine)
sSQL = sSQL & " AND IDAzienda=" & IDAzienda
sSQL = sSQL & "ORDER BY IDOggetto DESC"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaDocumento = fnNotNullN(rs!IDOggetto)
Else
    fncTrovaDocumento = 0
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_fncTrovaDocumento:
    MsgBox Err.Description, vbCritical, "Impossibile stampare"
    fncTrovaDocumento = 0
End Function
Public Function GET_LINK_SEZIONALE_ORDINE(IDTipoOggetto As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDSezionale "
    sSQL = sSQL & " FROM DefaultFilialePerTipoOggetto "
    sSQL = sSQL & " WHERE IDFiliale = " & IDFiliale
    sSQL = sSQL & " AND IDTipoOggetto = " & IDTipoOggetto
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_LINK_SEZIONALE_ORDINE = fnNotNullN(rs!IDSezionale)
    Else
        GET_LINK_SEZIONALE_ORDINE = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End Function
Private Function GET_LINK_PARAMETRO_FILIALE(IDFiliale As Long, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_PARAMETRO_FILIALE
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim AvviaImpostazioniTabellari As Boolean
Dim IDParametro As Long


sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Me.lblInfo.Caption = "CREAZIONE PARAMETRI FILIALE"
Me.txtControllo.Text = Me.txtControllo.Text & Me.lblInfo.Caption & vbCrLf
DoEvents
AvviaImpostazioniTabellari = False
If rs.EOF Then
    rs.AddNew
        rs!IDRV_POSchemaCoop = fnGetNewKey("RV_POSchemaCoop", "IDRV_POSchemaCoop")
        rs!IDAzienda = IDAzienda
        rs!IDFiliale = IDFiliale
        rs!IDUtente = 0
        rs!IDMagazzino_Carico = Me.cboMagLav.CurrentID
        rs!IDMagazzino_Vendita = Me.cboMagVend.CurrentID
        rs!IDCausale_Carico_Mag_Carico = 98
        rs!IDCausale_Scarico_Mag_Carico = 99
        rs!IDCausale_Carico_Mag_Vendita = 98
        rs!IDCausale_Scarico_Mag_Vendita = 99
        rs!IDTipoImballo = GET_LINK_TIPO_PRODOTTO(VarIDAzienda, 6)
        rs!IDTipoGrezzo = GET_LINK_TIPO_PRODOTTO(VarIDAzienda, 1)
        rs!IDTipoLavorato = GET_LINK_TIPO_PRODOTTO(VarIDAzienda, 2)
        rs!IDTipoCaloPeso = GET_LINK_TIPO_PRODOTTO(VarIDAzienda, 4)
        rs!IDTipoScarto = GET_LINK_TIPO_PRODOTTO(VarIDAzienda, 3)
        rs!IDTipoAumentoPeso = GET_LINK_TIPO_PRODOTTO(VarIDAzienda, 5)
        rs!IDCausaleCaloPeso = GET_LINK_FUNZIONE_PER_PARAMETRI("IDCausaleCaloPeso", VarIDAzienda)
        rs!IDCausaleScarto = GET_LINK_FUNZIONE_PER_PARAMETRI("IDCausaleScarto", VarIDAzienda)
        rs!IDCausaleAumentoPeso = GET_LINK_FUNZIONE_PER_PARAMETRI("IDCausaleAumentoPeso", VarIDAzienda)
        rs!IDCausaleAumentoPesoCarico = GET_LINK_FUNZIONE_PER_PARAMETRI("IDCausaleAumentoPesoCarico", VarIDAzienda)
        rs!IDCausaleCaloPesoCarico = GET_LINK_FUNZIONE_PER_PARAMETRI("IDCausaleCaloPesoCarico", VarIDAzienda)
        rs!IDCausaleScartoCarico = GET_LINK_FUNZIONE_PER_PARAMETRI("IDCausaleScartoCarico", VarIDAzienda)
        rs!AttivazioneNuovoMetodoCalcolo = 1
        rs!GestioneConferimento = False
        rs!IDTipoGestioneArticoliVendita = 0
        rs!IDTipoArrotondamento = 0
        rs!IDRV_POTipoPesoArticolo = 0
        rs!IDRV_POTipoComportamentoLavorazione = 3
        rs!IDRV_POTipoSceltaArticoloLottoCampagna = 0
        rs!IDRV_POTipoPesoArticolo = 0
        rs!LottoCampagnaObbligatorio = 0
        rs!IDTipoArrotondamentoConferimento = 0
        rs!StampaEtichetteNuove = 1
        rs!IDCategoriaAnagrafica = GET_LINK_CAT_ANA_SOCIO
        
        If NUMERO_DOCUMENTO_ORDINE > 0 Then
            rs!IDClienteCoop = LINK_CLIENTE_ORDINE_PRED
            rs!DataOrdineCoop = Date
            rs!NumeroOrdineCoop = NUMERO_DOCUMENTO_ORDINE
            rs!NumeroListaOrdineCoop = 1
        End If
        
        rs!VisAndamentoOrdDaLav = 1
        rs!VisAndamentoOrdDaOrd = 1
        rs!PrezziArticoloDaOrdine = 1
        rs!PrezziImballoDaOrdine = 1
        rs!PrezzoInclusoImballoDaOrdine = 1
        rs!IvaBloccata = 1
        rs!PedanaAutomatica = 1
        rs!OrdineAutomatico = 1
        rs!VisualizzaTotaliOrdinePrep = 1
        rs!VisualizzaTotaliOrdineSmist = 1
        
        rs!IDTipoCorpoFatturaSocio = 2
        rs!RiportaDescrizioneRifLiq = 1
        rs!IDListinoImballiDefault = GET_LINK_LISTINO_AZIENDA(VarIDAzienda)
        
        rs!IDOrdineLavorazioneIVGamma = LINK_ORDINE_LAV_IV_GAMMA
        rs!IDOrdineIVGamma = GET_LINK_ORDINE_GIACENZA(LINK_CLIENTE_ORDINE_PRED, NUMERO_DOCUMENTO_ORDINE, Date, IDAzienda)
        rs!IDTipoPedanaDefault = LINK_TIPO_PEDANA_PREDEFINITA
    rs.Update
    
    AvviaImpostazioniTabellari = True
    
End If

GET_LINK_PARAMETRO_FILIALE = fnNotNullN(rs!IDRV_POSchemaCoop)
IDParametro = fnNotNullN(rs!IDRV_POSchemaCoop)
rs.Close
Set rs = Nothing

If AvviaImpostazioniTabellari = True Then
    INSERIMENTO_OPERAZIONI_PER_DOC IDFiliale, IDAzienda, IDParametro
    INSERIMENTO_NOTE_PER_DOC_TESTA IDAzienda, IDFiliale, IDParametro
    CREAZIONE_LOTTO_CONF_VEND IDAzienda, IDFiliale
    CREAZIONE_LOTTO_CAMPAGNA IDAzienda, IDFiliale
    
End If

Exit Function

ERR_GET_LINK_PARAMETRO_FILIALE:
    MsgBox Err.Description, vbCritical, "GET_LINK_PARAMETRO_FILIALE"
    GET_LINK_PARAMETRO_FILIALE = 0
    Me.txtControllo.Text = Me.txtControllo.Text & " (" & Err.Description & ")" & vbCrLf
    
End Function
Private Function GET_LINK_TIPO_PRODOTTO(IDAzienda As Long, IDTipoProdottoConfig As Long) As Long
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim TipoProdotto As String

TipoProdotto = ""

sSQL = "SELECT TipoProdotto FROM TipoProdotto "
sSQL = sSQL & "WHERE IDTipoProdotto=" & IDTipoProdottoConfig
'sSQL = sSQL & " AND IDAzienda=" & IDAzienda
Set rsArc = New ADODB.Recordset

rsArc.Open sSQL, cnConfig

If rsArc.EOF Then
    TipoProdotto = ""
Else
    TipoProdotto = fnNotNull(rsArc!TipoProdotto)
End If

rsArc.Close
Set rsArc = Nothing

If Len(TipoProdotto) = 0 Then
    GET_LINK_TIPO_PRODOTTO = 0
    Exit Function
End If

sSQL = "SELECT IDTipoProdotto FROM TipoProdotto "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND TipoProdotto=" & fnNormString(TipoProdotto)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_TIPO_PRODOTTO = 0
Else
    GET_LINK_TIPO_PRODOTTO = fnNotNullN(rs!IDTipoProdotto)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_FUNZIONE_PER_PARAMETRI(NomeCampo As String, IDAzienda As Long) As Long
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim Funzione As String
Dim IDFunzioneLocal As Long


Funzione = ""
IDFunzioneLocal = 0
GET_LINK_FUNZIONE_PER_PARAMETRI = 0

sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
Set rsArc = New ADODB.Recordset

rsArc.Open sSQL, cnConfig

If rsArc.EOF Then
    IDFunzioneLocal = 0
Else
    IDFunzioneLocal = fnNotNullN(rsArc(NomeCampo).Value)
End If

rsArc.Close
Set rsArc = Nothing

If IDFunzioneLocal = 0 Then Exit Function

Funzione = GET_DESCRIZIONE_FUNZIONE_LOCAL(IDFunzioneLocal)

If Len(Funzione) = 0 Then Exit Function

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE IDTipoOggetto=9"
sSQL = sSQL & " AND Funzione=" & fnNormString(Funzione)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FUNZIONE_PER_PARAMETRI = 0
Else
    GET_LINK_FUNZIONE_PER_PARAMETRI = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_FUNZIONE_LOCAL(IDFunzioneLocal As Long) As String
Dim sSQL As String
Dim rsArc As ADODB.Recordset

sSQL = "SELECT Funzione FROM Funzione"
sSQL = sSQL & " WHERE IDFunzione=" & IDFunzioneLocal

Set rsArc = New ADODB.Recordset

rsArc.Open sSQL, cnConfig

If rsArc.EOF Then
    GET_DESCRIZIONE_FUNZIONE_LOCAL = ""
Else
    GET_DESCRIZIONE_FUNZIONE_LOCAL = fnNotNull(rsArc!Funzione)
End If

rsArc.Close
Set rsArc = Nothing

End Function
Private Function GET_LINK_CAT_ANA_SOCIO() As Long
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim CategoriaAnagrafica As String

CategoriaAnagrafica = ""

sSQL = "SELECT CategoriaAnagrafica FROM CategoriaAnagrafica "

Set rsArc = New ADODB.Recordset

rsArc.Open sSQL, cnConfig

If rsArc.EOF Then
    CategoriaAnagrafica = ""
Else
    CategoriaAnagrafica = fnNotNull(rsArc!CategoriaAnagrafica)
End If

rsArc.Close
Set rsArc = Nothing

If Len(CategoriaAnagrafica) = 0 Then
    GET_LINK_CAT_ANA_SOCIO = 0
    Exit Function
End If

sSQL = "SELECT IDCategoriaAnagrafica FROM CategoriaAnagrafica "
sSQL = sSQL & "WHERE CategoriaAnagrafica=" & fnNormString(CategoriaAnagrafica)
sSQL = sSQL & " AND IDTipoAnagrafica IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CAT_ANA_SOCIO = 0
Else
    GET_LINK_CAT_ANA_SOCIO = fnNotNullN(rs!IDCategoriaAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub CREAZIONE_SEZIONALI_PER_DOCUMENTI_COOP(IDParametriAzienda As Long, IDAzienda As Long, IDFiliale As Long)
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As ADODB.Recordset

Me.lblInfo.Caption = "CONFIGURAZIONE SEZIONALI PER DOCUMENTI"
Me.txtControllo.Text = Me.txtControllo.Text & Me.lblInfo.Caption & vbCrLf
DoEvents


sSQL = "SELECT * FROM RV_POSezionalePerDocumento "
sSQL = sSQL & "WHERE IDRV_POSchemaCoop=" & IDParametriAzienda
Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


sSQL = "SELECT * FROM RV_POSezionalePerDocumento "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
    
While Not rsArc.EOF
    If GET_ESISTENZA_SEZIONALE_PER_DOCUMENTO_COOP(fnNotNullN(rsArc!IDDocumentoCoop), IDParametriAzienda) = False Then
        rs.AddNew
            rs!IDRV_POSezionalePerDocumento = fnGetNewKey("RV_POSezionalePerDocumento", "IDRV_POSezionalePerDocumento")
            rs!IDRV_POSchemaCoop = IDParametriAzienda
            rs!IDDocumentoCoop = fnNotNullN(rsArc!IDDocumentoCoop)
            rs!Predefinito = True
            rs!IDSezionale = GET_LINK_SEZIONALE_PER_DOCUMENTI(fnNotNull(rsArc!IDSezionale), Me.cboFiliale.CurrentID)
        rs.Update
    End If
rsArc.MoveNext
Wend

rsArc.Close
Set rsArc = Nothing
End Sub
Private Function GET_ESISTENZA_SEZIONALE_PER_DOCUMENTO_COOP(IDTipoDocumentoCoop As Long, IDParametriFiliale As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POSezionalePerDocumento "
sSQL = sSQL & "WHERE IDRV_POSchemaCoop=" & IDParametriFiliale
sSQL = sSQL & " AND IDDocumentoCoop=" & IDTipoDocumentoCoop

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_SEZIONALE_PER_DOCUMENTO_COOP = False
Else
    GET_ESISTENZA_SEZIONALE_PER_DOCUMENTO_COOP = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_SEZIONALE_PER_DOCUMENTI(IDSezionaleLocal As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim Sezionale As String

Sezionale = GET_DESCRIZIONE_SEZIONALE_LOCAL(IDSezionaleLocal)

If Len(Sezionale) = 0 Then Exit Function

sSQL = "SELECT IDSezionale FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND Sezionale=" & fnNormString(Sezionale)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_SEZIONALE_PER_DOCUMENTI = 0
Else
    GET_LINK_SEZIONALE_PER_DOCUMENTI = fnNotNullN(rs!IDSezionale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_SEZIONALE_LOCAL(IDSezionaleLocal As Long) As String
Dim sSQL As String
Dim rsArc As ADODB.Recordset

sSQL = "SELECT Sezionale FROM Sezionale"
sSQL = sSQL & " WHERE IDSezionale=" & IDSezionaleLocal

Set rsArc = New ADODB.Recordset

rsArc.Open sSQL, cnConfig

If rsArc.EOF Then
    GET_DESCRIZIONE_SEZIONALE_LOCAL = ""
Else
    GET_DESCRIZIONE_SEZIONALE_LOCAL = fnNotNull(rsArc!Sezionale)
End If

rsArc.Close
Set rsArc = Nothing

End Function
Private Sub CREAZIONE_PROCESSI_PER_DOCUMENTI_COOP(IDParametriAzienda As Long, IDAzienda As Long, IDFiliale As Long)
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As ADODB.Recordset

Me.lblInfo.Caption = "CONFIGURAZIONE PROCESSI DI MAGAZZINO PER DOCUMENTI"
Me.txtControllo.Text = Me.txtControllo.Text & Me.lblInfo.Caption & vbCrLf
DoEvents

sSQL = "SELECT * FROM RV_POProcessiDocumentoCoop "
sSQL = sSQL & "WHERE IDRV_POSchemaCoop=" & IDParametriAzienda
Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POProcessiDocumentoCoop "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
    
While Not rsArc.EOF
    'If fnNotNullN(rsArc!IDRV_POProcessiDocumentoCoop) = 14 Then
    '    MsgBox "STOP"
    'End If


    If GET_ESISTENZA_PROCESSO_PER_DOCUMENTO_COOP(fnNotNullN(rsArc!IDDocumentoCoop), IDParametriAzienda, fnNotNullN(rsArc!IDTipoProcessoCoop)) = False Then
        rs.AddNew
            
            rs!IDRV_POProcessiDocumentoCoop = fnGetNewKey("RV_POProcessiDocumentoCoop", "IDRV_POProcessiDocumentoCoop")
            rs!IDRV_POSchemaCoop = IDParametriAzienda
            rs!IDDocumentoCoop = fnNotNullN(rsArc!IDDocumentoCoop)
            rs!IDTipoProcessoCoop = fnNotNullN(rsArc!IDTipoProcessoCoop)
            rs!IDTipoMagazzino = fnNotNullN(rsArc!IDTipoMagazzino)
            rs!IDFunzione = GET_LINK_FUNZIONE_PER_PROCESSI(fnNotNullN(rsArc!IDFunzione), VarIDAzienda)
            
            
        rs.Update
    End If
rsArc.MoveNext
Wend

rsArc.Close
Set rsArc = Nothing
End Sub
Private Function GET_ESISTENZA_PROCESSO_PER_DOCUMENTO_COOP(IDTipoDocumentoCoop As Long, IDParametriFiliale As Long, IDTipoProcessoCoop) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POProcessiDocumentoCoop "
sSQL = sSQL & "WHERE IDRV_POSchemaCoop=" & IDParametriFiliale
sSQL = sSQL & " AND IDDocumentoCoop=" & IDTipoDocumentoCoop
sSQL = sSQL & " AND IDTipoProcessoCoop=" & IDTipoProcessoCoop

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_PROCESSO_PER_DOCUMENTO_COOP = False
Else
    GET_ESISTENZA_PROCESSO_PER_DOCUMENTO_COOP = True
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_LINK_FUNZIONE_PER_PROCESSI(IDFunzioneLocal As Long, IDAzienda As Long) As Long
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim Funzione As String


GET_LINK_FUNZIONE_PER_PROCESSI = 0

Funzione = GET_DESCRIZIONE_FUNZIONE_LOCAL(IDFunzioneLocal)

If Len(Funzione) = 0 Then Exit Function

sSQL = "SELECT IDFunzione FROM Funzione "
sSQL = sSQL & "WHERE IDTipoOggetto=9"
sSQL = sSQL & " AND Funzione=" & fnNormString(Funzione)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FUNZIONE_PER_PROCESSI = 0
Else
    GET_LINK_FUNZIONE_PER_PROCESSI = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_PARAMETRO_FILIALE_AZIENDA(IDFiliale As Long, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_PARAMETRO_FILIALE
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim AvviaTabelle As Boolean

AvviaTabelle = False


sSQL = "SELECT * FROM RV_PO01_ParametriFiliale "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Me.lblInfo.Caption = "CREAZIONE PARAMETRI FILIALE PER GreenTop Azienda"
Me.txtControllo.Text = Me.txtControllo.Text & Me.lblInfo.Caption & vbCrLf
DoEvents

If rs.EOF Then
    
    rs.AddNew
        rs!IDRV_PO01_ParametriFiliale = fnGetNewKey("RV_PO01_ParametriFiliale", "IDRV_PO01_ParametriFiliale")
        rs!IDAzienda = IDAzienda
        rs!IDFiliale = IDFiliale
        rs!IDCategoriaAnagrafica = GET_LINK_CAT_ANA_SOCIO
        rs!NumeroLottoDiCampagna = 1
        
        rs!IDMagazzinoRiferimento = Me.cboMagLav.CurrentID
        rs!IDSezionaleProtSemp = GET_LINK_SEZIONALE("Passaporto", IDFiliale)
        rs!IDSezionaleProtDett = GET_LINK_SEZIONALE("Passaporto", IDFiliale)
        rs!IDRV_PO01_TipoPassaporto = 1
        rs!ContLottoPeriodoCamp = 1 'Indica che la numerazione del lotto avviene per periodo di campagna
        
        
        rs!IDTipoSchedaAgrofarmaci = GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO(VarIDAzienda, 2)
        rs!GeneraLottoAgrofarmaci = True
        rs!IDRV_PO01_TipoCR_Agrofarmaci = 1
        
        rs!IDTipoSchedaFertilizzanti = GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO(VarIDAzienda, 5)
        rs!GeneraLottoFertilizzanti = True
        rs!IDRV_PO01_TipoCR_Fertilizzanti = 1
        
        rs!IDTipoSchedaSementi = GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO(VarIDAzienda, 3)
        rs!GeneraLottoSementi = True
        rs!IDRV_PO01_TipoCR_Sementi = 1
        
        rs!IDTipoSchedaTrappole = GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO(VarIDAzienda, 4)
        rs!GeneraLottoTrappole = True
        rs!IDRV_PO01_TipoCR_Trappole = 1
        
        rs!IDTipoSchedaAltreOperazioni = GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO(VarIDAzienda, 6)
        rs!GeneraLottoAltreOperazioni = False
        rs!IDRV_PO01_TipoCR_AltreOperazioni = 0
        
        rs!IDTipoSchedaVendita = GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO(VarIDAzienda, 1)
        rs!IDRV_PO01_TipoCR_Vendita = 2
        
    rs.Update
    AvviaTabelle = True
End If

GET_LINK_PARAMETRO_FILIALE_AZIENDA = fnNotNullN(rs!IDRV_PO01_ParametriFiliale)

rs.Close
Set rs = Nothing


If AvviaTabelle = True Then
    GET_FAMIGLIA_PRODOTTI
    GET_VARIETA_PRODOTTI
End If
Exit Function

ERR_GET_LINK_PARAMETRO_FILIALE:
    MsgBox Err.Description, vbCritical, "GET_LINK_PARAMETRO_FILIALE_AZIENDA"
    GET_LINK_PARAMETRO_FILIALE_AZIENDA = 0
    Me.txtControllo.Text = Me.txtControllo.Text & " (" & Err.Description & ")" & vbCrLf
    
End Function
Private Function GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO(IDAzienda As Long, IDTipoGruppoEquivalenza As Long) As Long
Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rs As DmtOleDbLib.adoResultset
Dim GruppoEquivalenza As String

GruppoEquivalenza = ""

sSQL = "SELECT GruppoEquivalenzaArticolo FROM GruppoEquivalenzaArticolo "
sSQL = sSQL & "WHERE IDGruppoEquivalenzaArticolo=" & IDTipoGruppoEquivalenza
Set rsArc = New ADODB.Recordset

rsArc.Open sSQL, cnConfig

If rsArc.EOF Then
    GruppoEquivalenza = ""
Else
    GruppoEquivalenza = fnNotNull(rsArc!GruppoEquivalenzaArticolo)
End If

rsArc.Close
Set rsArc = Nothing

If Len(GruppoEquivalenza) = 0 Then
    GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO = 0
    Exit Function
End If

sSQL = "SELECT IDGruppoEquivalenzaArticolo FROM GruppoEquivalenzaArticolo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND GruppoEquivalenzaArticolo=" & fnNormString(GruppoEquivalenza)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO = 0
Else
    GET_LINK_GRUPPO_EQUIVALENZA_ARTICOLO = fnNotNullN(rs!IDGruppoEquivalenzaArticolo)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub INSERIMENTO_OPERAZIONI_PER_DOC(IDFiliale As Long, IDAzienda As Long, IDSchemaCoop As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim I As Long

sSQL = "DELETE FROM RV_POOperazionePerDoc "
sSQL = sSQL & "WHERE IDRV_POSchemaCoop=" & IDSchemaCoop
CnDMT.Execute sSQL

sSQL = "SELECT * FROM RV_POOperazionePerDoc "

Set rs = New ADODB.Recordset
rs.Open sSQL, cnConfig


Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    rsNew.AddNew
        For I = 0 To rs.Fields.Count - 1
            Select Case rs.Fields(I).Name
                Case "IDRV_POOperazionePerDoc"
                    rsNew.Fields(rs.Fields(I).Name).Value = fnGetNewKey("RV_POOperazionePerDoc", "IDRV_POOperazionePerDoc")
                Case "IDRV_POSchemaCoop"
                    rsNew.Fields(rs.Fields(I).Name).Value = IDSchemaCoop
                Case Else
                    rsNew.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
            End Select
        Next
    rsNew.Update
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.Close
Set rs = Nothing
End Sub
Private Function GET_LINK_SEZIONALE(Sezionale As String, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale FROM Sezionale "
sSQL = sSQL & " WHERE Sezionale=" & fnNormString(Sezionale)
sSQL = sSQL & " AND IDFiliale=" & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_SEZIONALE = 0
Else
    GET_LINK_SEZIONALE = fnNotNullN(rs!IDSezionale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub INSERIMENTO_NOTE_PER_DOC(IDAzienda As Long, IDFiliale As Long, IDSchemaCoop As Long, IDTipoOggetto As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM RV_PONoteDocumentiCoop "
sSQL = sSQL & " WHERE IDRV_POSchemaCoop=" & IDSchemaCoop
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    
        rs!IDRV_PONoteDocumentiCoop = fnGetNewKey("RV_PONoteDocumentiCoop", "IDRV_PONoteDocumentiCoop")
        rs!IDRV_POSchemaCoop = IDSchemaCoop
        rs!Annotazioni1 = ""
        rs!Annotazioni2 = ""
        rs!Annotazioni3 = ""
        rs!IDTipoOggetto = IDTipoOggetto
        rs!Annotazioni4 = ""
        rs!Annotazioni5 = ""
        If IDTipoOggetto <> 4 Then
            rs!StampaLotti = True
        Else
            rs!StampaLotti = False
        End If
        rs!NonStampaNumeroProtICE = True
        rs!NonStampaImporti = False
        rs!IDFiliale = IDFiliale
        rs!IDAzienda = IDAzienda
        rs!TipoModalitaReport = 3
        rs!NonStampaImportiRighe = False
        rs!NonStampareDescrizioneArticolo = False
        rs!StampaDescrizioneRidottaArticolo = False
        rs!Annotazioni6 = ""
        rs!Annotazioni7 = ""
        rs!Annotazioni8 = ""
        rs!Annotazioni9 = ""
        rs!Annotazioni10 = ""
        rs!NonStampareRifOrdCliente = 1
        rs!StampaNoteCliente = 0
        rs!NonStampareImballi = 1
        rs!StampaVarietaArticolo = 0
        rs!NonStampareImballiCMR = 1
        rs!StampaLottiCMR = 0
        rs!StampaImbSuRigaArt = 0
        rs!StampaAccontiCommissioni = 0
        rs!GrandezzaCarattere = 0
        rs!TipoPiedePagina = 0
        rs!NonStampareImportoUnitario = 0
        rs!StampaNomeRagioneSociale = 0
        rs!DescrizioneAggiuntivaOggetto = ""
        
        
        
    rs.Update
End If
rs.Close
Set rs = Nothing


End Sub
Private Sub INSERIMENTO_NOTE_PER_DOC_TESTA(IDAzienda As Long, IDFiliale As Long, IDSchemaCoop As Long)
    INSERIMENTO_NOTE_PER_DOC IDAzienda, IDFiliale, IDSchemaCoop, 114
    INSERIMENTO_NOTE_PER_DOC IDAzienda, IDFiliale, IDSchemaCoop, 2
    INSERIMENTO_NOTE_PER_DOC IDAzienda, IDFiliale, IDSchemaCoop, 8
    INSERIMENTO_NOTE_PER_DOC IDAzienda, IDFiliale, IDSchemaCoop, 4
    INSERIMENTO_NOTE_PER_DOC IDAzienda, IDFiliale, IDSchemaCoop, fnGetTipoOggetto("RV_POCaricoMerceL")
    INSERIMENTO_NOTE_PER_DOC IDAzienda, IDFiliale, IDSchemaCoop, fnGetTipoOggetto("RV_POFattAcqL")
    
    
    
End Sub
Private Function GET_LINK_LISTINO_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset
Dim Link_Listino_Imballo As Long

GET_LINK_LISTINO_AZIENDA = 0


'''''''''''''''''''''''LISTINO AZIENDA'''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDListinoDiBase "
sSQL = sSQL & "FROM ConfigurazioneVendite "
sSQL = sSQL & " WHERE IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LISTINO_AZIENDA = 0
Else
    GET_LINK_LISTINO_AZIENDA = fnNotNullN(rs!IDListinoDiBase)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Private Sub CREAZIONE_LOTTO_CONF_VEND(IDAzienda As Long, IDFiliale As Long)
Dim sSQL As String
Dim rsNew As ADODB.Recordset

Dim Link_Testa_Lotto As Long


Link_Testa_Lotto = 0
sSQL = "SELECT * FROM RV_POLottoCostruzioneTesta "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDFiliale=" & IDFiliale

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    Link_Testa_Lotto = fnGetNewKey("RV_POLottoCostruzioneTesta", "IDRV_POLottoCostruzioneTesta")
    rsNew.AddNew
        rsNew!IDRV_POLottoCostruzioneTesta = Link_Testa_Lotto
        rsNew!IDFiliale = IDFiliale
        rsNew!IDAzienda = IDAzienda
        rsNew!IDSocio = 0
        rsNew!PosConferimento = 0
        rsNew!PosVendita = 0
        rsNew!SenzaCodiceRiferimento_Conf = 0
        rsNew!SenzaCodiceRiferimento_Vend = 0
    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing


If Link_Testa_Lotto > 0 Then
    sSQL = "DELETE FROM RV_POLottoCostruzioneRighe "
    sSQL = sSQL & "WHERE IDRV_POLottoCostruzioneTesta=" & Link_Testa_Lotto
    CnDMT.Execute sSQL

    sSQL = "SELECT * FROM RV_POLottoCostruzioneRighe "
    sSQL = sSQL & "WHERE IDRV_POLottoCostruzioneTesta=" & Link_Testa_Lotto
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_POLottoCostruzioneRighe = fnGetNewKey("RV_POLottoCostruzioneRighe", "IDRV_POLottoCostruzioneRighe")
            rsNew!IDRV_POLottoCostruzioneTesta = Link_Testa_Lotto
            rsNew!IDRV_POLottoComp = 23
            rsNew!Posizione = 1
            rsNew!Lunghezza = 3
            rsNew!TipoStringaLotto = 1
            rsNew!TipoLotto = 1
            rsNew!Testo = ""
            rsNew!SXDX = 1
        rsNew.Update
        rsNew.AddNew
            rsNew!IDRV_POLottoCostruzioneRighe = fnGetNewKey("RV_POLottoCostruzioneRighe", "IDRV_POLottoCostruzioneRighe")
            rsNew!IDRV_POLottoCostruzioneTesta = Link_Testa_Lotto
            rsNew!IDRV_POLottoComp = 14
            rsNew!Posizione = 1
            rsNew!Lunghezza = 1
            rsNew!TipoStringaLotto = 1
            rsNew!TipoLotto = 1
            rsNew!Testo = ""
            rsNew!SXDX = 1
        rsNew.Update
         rsNew.AddNew
            rsNew!IDRV_POLottoCostruzioneRighe = fnGetNewKey("RV_POLottoCostruzioneRighe", "IDRV_POLottoCostruzioneRighe")
            rsNew!IDRV_POLottoCostruzioneTesta = Link_Testa_Lotto
            rsNew!IDRV_POLottoComp = 23
            rsNew!Posizione = 1
            rsNew!Lunghezza = 3
            rsNew!TipoStringaLotto = 1
            rsNew!TipoLotto = 2
            rsNew!Testo = ""
            rsNew!SXDX = 1
            
        rsNew.Update
        rsNew.AddNew
            rsNew!IDRV_POLottoCostruzioneRighe = fnGetNewKey("RV_POLottoCostruzioneRighe", "IDRV_POLottoCostruzioneRighe")
            rsNew!IDRV_POLottoCostruzioneTesta = Link_Testa_Lotto
            rsNew!IDRV_POLottoComp = 14
            rsNew!Posizione = 1
            rsNew!Lunghezza = 1
            rsNew!TipoStringaLotto = 1
            rsNew!TipoLotto = 2
            rsNew!Testo = ""
            rsNew!SXDX = 1
        rsNew.Update
    End If
    
    rsNew.Close
    Set rsNew = Nothing
End If




End Sub

Private Sub CREAZIONE_LOTTO_CAMPAGNA(IDAzienda As Long, IDFiliale As Long)
Dim sSQL As String
Dim rsNew As ADODB.Recordset

Dim Link_Testa_Lotto As Long

Link_Testa_Lotto = 0
sSQL = "SELECT * FROM RV_PO01_ConfigLottoCampagna "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDFiliale=" & IDFiliale

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    Link_Testa_Lotto = fnGetNewKey("RV_PO01_ConfigLottoCampagna", "IDRV_PO01_ConfigLottoCampagna")
    rsNew.AddNew
        rsNew!IDRV_PO01_ConfigLottoCampagna = Link_Testa_Lotto
        rsNew!IDFiliale = IDFiliale
        rsNew!IDAzienda = IDAzienda
        rsNew!UtilizzaProgressivoLotto = 1
        rsNew!IDRV_PO01_LatoProgressivoLotto = 1

    rsNew.Update
End If

rsNew.Close
Set rsNew = Nothing


If Link_Testa_Lotto > 0 Then
    sSQL = "DELETE FROM RV_PO01_ConfigLottoCampagnaRighe "
    sSQL = sSQL & "WHERE IDRV_PO01_ConfigLottoCampagna=" & Link_Testa_Lotto
    CnDMT.Execute sSQL

    sSQL = "SELECT * FROM RV_PO01_ConfigLottoCampagnaRighe "
    sSQL = sSQL & "WHERE IDRV_PO01_ConfigLottoCampagna=" & Link_Testa_Lotto
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_ConfigLottoCampagnaRighe = fnGetNewKey("RV_PO01_ConfigLottoCampagnaRighe", "IDRV_PO01_ConfigLottoCampagnaRighe")
            rsNew!IDRV_PO01_ConfigLottoCampagna = Link_Testa_Lotto
            rsNew!IDRV_PO01_CampiLottoCampagna = 2
            rsNew!Lunghezza = 5
            rsNew!Progressivo = 1
            rsNew!Testo = ""
            
        rsNew.Update
        rsNew.AddNew
            rsNew!IDRV_PO01_ConfigLottoCampagnaRighe = fnGetNewKey("RV_PO01_ConfigLottoCampagnaRighe", "IDRV_PO01_ConfigLottoCampagnaRighe")
            rsNew!IDRV_PO01_ConfigLottoCampagna = Link_Testa_Lotto
            rsNew!IDRV_PO01_CampiLottoCampagna = 1
            rsNew!Lunghezza = 1
            rsNew!Progressivo = 2
            rsNew!Testo = "-"
            
        rsNew.Update
        rsNew.AddNew
            rsNew!IDRV_PO01_ConfigLottoCampagnaRighe = fnGetNewKey("RV_PO01_ConfigLottoCampagnaRighe", "IDRV_PO01_ConfigLottoCampagnaRighe")
            rsNew!IDRV_PO01_ConfigLottoCampagna = Link_Testa_Lotto
            rsNew!IDRV_PO01_CampiLottoCampagna = 36
            rsNew!Lunghezza = 20
            rsNew!Progressivo = 3
            rsNew!Testo = ""
        rsNew.Update
        rsNew.AddNew
            rsNew!IDRV_PO01_ConfigLottoCampagnaRighe = fnGetNewKey("RV_PO01_ConfigLottoCampagnaRighe", "IDRV_PO01_ConfigLottoCampagnaRighe")
            rsNew!IDRV_PO01_ConfigLottoCampagna = Link_Testa_Lotto
            rsNew!IDRV_PO01_CampiLottoCampagna = 1
            rsNew!Lunghezza = 1
            rsNew!Progressivo = 4
            rsNew!Testo = "-"
            
        rsNew.Update
        rsNew.AddNew
            rsNew!IDRV_PO01_ConfigLottoCampagnaRighe = fnGetNewKey("RV_PO01_ConfigLottoCampagnaRighe", "IDRV_PO01_ConfigLottoCampagnaRighe")
            rsNew!IDRV_PO01_ConfigLottoCampagna = Link_Testa_Lotto
            rsNew!IDRV_PO01_CampiLottoCampagna = 10
            rsNew!Lunghezza = 2
            rsNew!Progressivo = 5
            rsNew!Testo = ""
            
        rsNew.Update
        rsNew.AddNew
            rsNew!IDRV_PO01_ConfigLottoCampagnaRighe = fnGetNewKey("RV_PO01_ConfigLottoCampagnaRighe", "IDRV_PO01_ConfigLottoCampagnaRighe")
            rsNew!IDRV_PO01_ConfigLottoCampagna = Link_Testa_Lotto
            rsNew!IDRV_PO01_CampiLottoCampagna = 1
            rsNew!Lunghezza = 1
            rsNew!Progressivo = 5
            rsNew!Testo = "-"
            
        rsNew.Update
    End If
    
    rsNew.Close
    Set rsNew = Nothing
End If




End Sub

Private Sub UNITA_DI_MISURA_COOP()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "UNITA DI MISURA"
Me.txtControllo.Text = Me.txtControllo.Text & "UNITA DI MISURA" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM UnitaDiMisura "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM UnitaDiMisura "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!UnitaDiMisura) & vbCrLf
    DoEvents
    rsNew.Filter = "UnitaDiMisura=" & fnNormString(rsArc!UnitaDiMisura)
    If rsNew.EOF Then
        rsNew.AddNew
        rsNew!IDUnitaDiMisura = fnGetNewKey("UnitaDiMisura", "IDUnitaDiMisura")
        rsNew!UnitaDiMisura = fnNotNull(rsArc!UnitaDiMisura)
        rsNew!DescrizioneFattura = Trim(fnNotNull(rsArc!DescrizioneFattura))

    End If
    rsNew!RV_POIDUnitaDiMisuraCoop = fnNotNullN(rsArc!RV_POIDUnitaDiMisuraCoop)
    
    rsNew.Update
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "UNITA_DI_MISURA_COOP"
    
End Sub
Private Function GET_ESISTENZA_UNITA_DI_MISURA(UnitaDiMisura As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisura FROM UnitaDiMisura "
sSQL = sSQL & " WHERE UnitaDiMisura=" & fnNormString(UnitaDiMisura)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_UNITA_DI_MISURA = False
Else
    GET_ESISTENZA_UNITA_DI_MISURA = True
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_PARAMETRO_FILIALE_FARMACIA(IDFiliale As Long, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_PARAMETRO_FILIALE
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim AvviaCorpoParametri As Boolean
Dim Link_Parametro As Long

AvviaCorpoParametri = False


sSQL = "SELECT * FROM RV_PO10_ParametriAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Me.lblInfo.Caption = "CREAZIONE PARAMETRI FILIALE PER GreenTop Azienda"
Me.txtControllo.Text = Me.txtControllo.Text & Me.lblInfo.Caption & vbCrLf
DoEvents

If rs.EOF Then
    AvviaCorpoParametri = True
    rs.AddNew
        rs!IDRV_PO10_ParametriAzienda = fnGetNewKey("RV_PO10_ParametriAzienda", "IDRV_PO10_ParametriAzienda")
        rs!IDAzienda = IDAzienda
        rs!IDAnagrafica = GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda)
        
        Link_Parametro = rs!IDRV_PO10_ParametriAzienda
    rs.Update
End If

GET_LINK_PARAMETRO_FILIALE_FARMACIA = fnNotNullN(rs!IDRV_PO10_ParametriAzienda)

rs.Close
Set rs = Nothing

If AvviaCorpoParametri = True Then
    'PROTOCOLLO
    CREA_PROTOCOLLO_FARMACIA Link_Parametro, IDFiliale
    'DOCUMENTI
    CREA_DOCUMENTI_ELA_FARMACIA Link_Parametro, IDFiliale
End If

Exit Function

ERR_GET_LINK_PARAMETRO_FILIALE:
    MsgBox Err.Description, vbCritical, "GET_LINK_PARAMETRO_FILIALE_FARMACIA"
    GET_LINK_PARAMETRO_FILIALE_FARMACIA = 0
    Me.txtControllo.Text = Me.txtControllo.Text & " (" & Err.Description & ")" & vbCrLf
    
End Function
Private Function GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM Azienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_AZIENDA = 0
Else
    GET_LINK_ANAGRAFICA_AZIENDA = fnNotNull(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREA_PROTOCOLLO_FARMACIA(IDSchema As Long, IDFiliale As Long)
Dim sSQL As String

Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "CREA PROTOCOLLO FARMACIA"
Me.txtControllo.Text = Me.txtControllo.Text & "CREA PROTOCOLLO FARMACIA" & vbCrLf
DoEvents

sSQL = "SELECT * FROM RV_PO10_ProtocolloPresidio "
sSQL = sSQL & "WHERE IDRV_PO10_ParametriAzienda=" & IDSchema

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
If rsNew.EOF Then

    rsNew.AddNew
        rsNew!IDRV_PO10_ProtocolloPresidio = fnGetNewKey("RV_PO10_ProtocolloPresidio", "IDRV_PO10_ProtocolloPresidio")
        rsNew!IDRV_PO10_ParametriAzienda = IDSchema
        rsNew!IDSezionale = GET_LINK_SEZIONALE("Passaporto", IDFiliale)
        rsNew!Predefinito = True
        rsNew!Descrizione = ""
    rsNew.Update
End If
rsNew.Close
Set rsNew = Nothing


End Sub
Private Sub CREA_DOCUMENTI_ELA_FARMACIA(IDSchema As Long, IDFiliale As Long)
Dim sSQL As String

Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "CREA DOCUMENTI DA ELABORARE FARMACIA"
Me.txtControllo.Text = Me.txtControllo.Text & "CREA DOCUMENTI DA ELABORARE FARMACIA" & vbCrLf
DoEvents

sSQL = "SELECT * FROM RV_PO10_DocumentiDaEla "
sSQL = sSQL & "WHERE IDRV_PO10_ParametriAzienda=" & IDSchema

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
If rsNew.EOF Then

    rsNew.AddNew
        rsNew!IDRV_PO10_DocumentiDaEla = fnGetNewKey("RV_PO10_DocumentiDaEla", "IDRV_PO10_DocumentiDaEla")
        rsNew!IDRV_PO10_ParametriAzienda = IDSchema
        rsNew!IDTipoOggetto = 2
        rsNew!DaElaborare = True
        rsNew!IDReportPerTipoOggetto = GET_LINK_REPORT_TIPO_OGGETTO("RV_PO10_PresidiDDT.rpt", "RV_PO10_Protocollatura")
    rsNew.Update
    
    
    rsNew.AddNew
        rsNew!IDRV_PO10_DocumentiDaEla = fnGetNewKey("RV_PO10_DocumentiDaEla", "IDRV_PO10_DocumentiDaEla")
        rsNew!IDRV_PO10_ParametriAzienda = IDSchema
        rsNew!IDTipoOggetto = 114
        rsNew!DaElaborare = True
        rsNew!IDReportPerTipoOggetto = GET_LINK_REPORT_TIPO_OGGETTO("RV_PO10_PresidiFA.rpt", "RV_PO10_Protocollatura")
    rsNew.Update
End If
rsNew.Close
Set rsNew = Nothing


End Sub
Private Function GET_LINK_REPORT_TIPO_OGGETTO(Report As String, Gestore As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ReportTipoOggetto.IDReportTipoOggetto, ReportTipoOggetto.ReportTipoOggetto "
sSQL = sSQL & "FROM ReportTipoOggetto INNER JOIN "
sSQL = sSQL & "TipoOggetto ON ReportTipoOggetto.IDTipoOggetto = TipoOggetto.IDTipoOggetto INNER JOIN "
sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
sSQL = sSQL & "WHERE Gestore=" & fnNormString(Gestore)
sSQL = sSQL & " AND ReportTipoOggetto=" & fnNormString(Report)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REPORT_TIPO_OGGETTO = 0
Else
    GET_LINK_REPORT_TIPO_OGGETTO = fnNotNullN(rs!IDReportTipoOggetto)
    
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_PARAMETRO_FILIALE_LIQUIDAZIONE(IDFiliale As Long, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_PARAMETRO_FILIALE
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim AvviaCorpoParametri As Boolean
Dim Link_Parametro As Long

AvviaCorpoParametri = False


sSQL = "SELECT * FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale


Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Me.lblInfo.Caption = "CREAZIONE PARAMETRI FILIALE PER LIQUIDAZIONE"
Me.txtControllo.Text = Me.txtControllo.Text & Me.lblInfo.Caption & vbCrLf
DoEvents

If rs.EOF Then
    AvviaCorpoParametri = True
    rs.AddNew
        rs!IDRV_POCalcoloLiqTesta = fnGetNewKey("RV_POCalcoloLiqTesta", "IDRV_POCalcoloLiqTesta")
        rs!IDFiliale = IDFiliale
        rs!IDSocio = 0
        rs!AddebitoImballo = 0
        rs!IDListinoImballo = 0
        rs!IDTipoImportoArticolo = 3
        rs!IDTipoImportoDocumento = 1
        rs!IDTipoQuantita = 4
        rs!ArticoliDiQuadratura = False
        rs!IDTipoLiquidazione = 1
        rs!IDTipoPrezzoMedio = 1
        rs!IDTipoCalcoloPrezzoMedioNC = 1
        rs!IDTipoCalcoloPrezzoMedioND = 1
        rs!IDRV_POTipoLiqConf = 0
        rs!PrezziMediInCampionatura = 0
        rs!AggiornaPMNonBloccatiCamp = 0
        rs!CalcoloTrattenuteSuCamp = 0
        rs!IDRV_POTipoConfLiquidazioneNuovo = 0
        rs!IDRV_POTipoConfLiquidazioneChiuso = 0
        rs!CalcolaPrezzoMedioIVGamma = 0
        
        
        
        
        
        Link_Parametro = rs!IDRV_POCalcoloLiqTesta
    rs.Update
End If

GET_LINK_PARAMETRO_FILIALE_LIQUIDAZIONE = fnNotNullN(rs!IDRV_POCalcoloLiqTesta)

rs.Close
Set rs = Nothing

If AvviaCorpoParametri = True Then
    CREA_RIGHE_LIQUIDAZIONE Link_Parametro
End If

Exit Function

ERR_GET_LINK_PARAMETRO_FILIALE:
    MsgBox Err.Description, vbCritical, "GET_LINK_PARAMETRO_FILIALE_LIQUIDAZIONE"
    GET_LINK_PARAMETRO_FILIALE_LIQUIDAZIONE = 0
    Me.txtControllo.Text = Me.txtControllo.Text & " (" & Err.Description & ")" & vbCrLf
    
End Function

Private Sub CREA_RIGHE_LIQUIDAZIONE(IDSchema As Long)
Dim sSQL As String

Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "CONFIGURAZIONE LIQUIDAZIONE"
Me.txtControllo.Text = Me.txtControllo.Text & "CONFIGURAZIONE LIQUIDAZIONE" & vbCrLf
DoEvents

sSQL = "SELECT * FROM RV_POCalcoloLiqRighe "
sSQL = sSQL & "WHERE IDRV_POCalcoloLiqTesta=" & IDSchema

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rsNew.EOF Then
    rsNew.AddNew
        rsNew!IDRV_POCalcoloLiqRighe = fnGetNewKey("RV_POCalcoloLiqRighe", "IDRV_POCalcoloLiqRighe")
        rsNew!IDRV_POCalcoloLiqTesta = IDSchema
        rsNew!IDRV_POTipoTrattenuta = 6
        rsNew!Posizione = 1
    rsNew.Update
End If
rsNew.Close
Set rsNew = Nothing


End Sub
Private Function FN_CREA_ORDINE_LAV_IVGAMMA(IDCliente As Long, IDAzienda As Long, IDFiliale As Long) As Long

On Error GoTo ERR_FN_CREA_ORDINE
Dim ObjDoc As DmtDocs.cDocument
Dim cDefault As Collection
Dim IDSezionale As Long
Dim NumeroDocumento As String

Me.lblInfo.Caption = "CREAZIONE ORDINE DI LAVORAZIONE IV GAMMA"
Me.txtControllo.Text = Me.txtControllo.Text & "CREAZIONE ORDINE DI LAVORAZIONE IV GAMMA" & vbCrLf
DoEvents

If Not (ObjDoc Is Nothing) Then
    Set ObjDoc = Nothing
End If

Set ObjDoc = New DmtDocs.cDocument
 
    With ObjDoc
    
        Set .Connection = CnDMT
        .IDAzienda = IDAzienda
        .IDAttivitaAzienda = GET_ATTIVITA_AZIENDA(IDAzienda)
        .IDFiliale = IDFiliale
        .SetTipoOggetto 15
        .IDFunzione = 128
        
        .IDEsercizio = fncEsercizio(Date, IDAzienda)
        .IDSezionale = GET_LINK_SEZIONALE_ORDINE(.IDTipoOggetto, .IDFiliale)
        .IDTipoAnagrafica = 2
        .IDUtente = 1
        .Descrizione = "Ordine da cliente"
        .DataEmissione = Date
        .Numero = 0
        If .Tables.Count = 0 Then
        'Se Tables.Count = 0 vuol dire che l'oggetto
        'DmtDocs non è mai stato inizializzato
            .Clear
            .SetTipoOggetto 15
        Else
            .ClearValues
        End If
        
        ObjDoc.Tables("ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
        
        .Field "Link_Doc_magazzino", Me.cboMagVend.CurrentID, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Doc_sezionale", ObjDoc.IDSezionale, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Link_Val_valuta", 9, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_data", ObjDoc.DataEmissione, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_numero", 0, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
        .ReadDataFromCliFo IDCliente
        
        .Field "Link_Doc_pagamento", 1, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        .Field "Doc_ordine_chiuso", 0, "ValoriOggettoPerTipo" & fnGetHex(ObjDoc.IDTipoOggetto)
        
    
        Set ObjDoc.Scadenze = Nothing
        ObjDoc.PerformDocument Nothing
        NumeroDocumento = "0"
        NumeroDocumento = ObjDoc.Insert
        
        ObjDoc.Update
        
        FN_CREA_ORDINE_LAV_IVGAMMA = ObjDoc.IDOggetto
        
        
        If (ObjDoc.IDOggetto) > 0 Then
            sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
            sSQL = sSQL & "RV_PONumeroOrdinePadre=" & fnNormNumber(ObjDoc.Numero) & ", "
            sSQL = sSQL & "RV_PODataOrdinePadre=" & fnNormDate(ObjDoc.DataEmissione) & ", "
            sSQL = sSQL & "RV_PONumeroListaPrelievo=1" & ", "
            sSQL = sSQL & "RV_POIDOrdinePadre=" & ObjDoc.IDOggetto
            sSQL = sSQL & " WHERE IDOggetto=" & ObjDoc.IDOggetto
            
            CnDMT.Execute sSQL
        End If
        
    End With
    
    Set ObjDoc = Nothing
   
 Exit Function
ERR_FN_CREA_ORDINE:
    MsgBox Err.Description, vbCritical, "FN_CREA_ORDINE_LAV_IVGAMMA"
    FN_CREA_ORDINE_LAV_IVGAMMA = 0
End Function

Public Function GET_LINK_ORDINE_GIACENZA(IDCliente As Long, NumeroOrdine As Long, DataOrdine As String, IDAzienda As Long) As Long
On errro GoTo ERR_GET_LINK_ORDINE_GIACENZA
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ValoriOggettoPerTipo000F.IDOggetto "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F INNER JOIN "
sSQL = sSQL & "Oggetto ON ValoriOggettoPerTipo000F.IDOggetto = Oggetto.IDOggetto "
sSQL = sSQL & "AND ValoriOggettoPerTipo000F.IDTipoOggetto = Oggetto.IDTipoOggetto "
sSQL = sSQL & "WHERE Doc_ordine_chiuso = 0 "
sSQL = sSQL & " AND IDAzienda=" & IDAzienda
    
    'If Me.cdCliente.KeyFieldID > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Link_nom_anagrafica=" & IDCliente
    'End If
    
    'If Me.txtNumeroOrdine.Value > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Doc_numero=" & NumeroOrdine
    'End If
    
    'If Me.txtDataOrdine.Value > 0 Then
        sSQL_WHERE = sSQL_WHERE & " AND Doc_data=" & fnNormDate(DataOrdine)
    'End If
    
    sSQL = sSQL & sSQL_WHERE & " ORDER BY Doc_data DESC, Doc_numero DESC"

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_LINK_ORDINE_GIACENZA = 0
Else
    GET_LINK_ORDINE_GIACENZA = fnNotNullN(rs!IDOggetto)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_ORDINE_GIACENZA:
    MsgBox Err.Description, vbCritical, "GET_LINK_ORDINE_GIACENZA"
    GET_LINK_ORDINE_GIACENZA = 0
End Function
Private Sub GET_FAMIGLIA_PRODOTTI()
On Error GoTo ERR_GET_FAMIGLIA_PRODOTTI
Dim sSQL As String
Dim RigaFile As String
Dim ArrayFile() As String
Dim rsNew As ADODB.Recordset
Dim Descrizione As String
Dim IDRiga As Long

sSQL = "SELECT * FROM RV_PO01_FamigliaProdotti "

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Open App.Path & "\ProdottiFamiglia.txt" For Input As #1

Do Until EOF(1)
    
    Input #1, RigaFile
    
    ArrayFile = Split(RigaFile, "-")
    IDRiga = ArrayFile(0)
    
    Descrizione = Trim(ArrayFile(1))
    
    rsNew.Filter = "IDRV_PO01_FamigliaProdotti=" & IDRiga
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_FamigliaProdotti = IDRiga
            rsNew!FamigliaProdotti = Descrizione
            rsNew!CodiceImEx = IDRiga
            rsNew!ResaMinima = 0
            rsNew!ResaMassima = 0
        rsNew.Update
    End If
    rsNew.Filter = vbNullString
Loop

Close #1

rsNew.Close
Set rs = Nothing
Exit Sub
ERR_GET_FAMIGLIA_PRODOTTI:
    MsgBox Err.Description, vbCritical, "GET_FAMIGLIA_PRODOTTI"
End Sub
Private Sub GET_VARIETA_PRODOTTI()
On Error GoTo ERR_GET_VARIETA_PRODOTTI
Dim sSQL As String
Dim RigaFile As String
Dim ArrayFile() As String
Dim rsNew As ADODB.Recordset
Dim Descrizione As String
Dim IDRiga As Long
Dim CodiceEx As String

Dim IDFamiglia As Long

sSQL = "SELECT * FROM RV_PO01_Varieta "

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Open App.Path & "\ProdottiVarieta.txt" For Input As #1

Do Until EOF(1)
    
    Input #1, RigaFile
    
    ArrayFile = Split(RigaFile, "-")
    IDFamiglia = ArrayFile(0)
    CodiceEx = ArrayFile(1)
    Descrizione = Trim(ArrayFile(2))
    
    rsNew.Filter = "IDRV_PO01_FamigliaProdotti=" & IDFamiglia
    rsNew.Filter = rsNew.Filter & " AND Varieta=" & fnNormString(Descrizione)
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_Varieta = fnGetNewKey("RV_PO01_Varieta", "IDRV_PO01_Varieta")
            rsNew!IDRV_PO01_FamigliaProdotti = IDFamiglia
            rsNew!Varieta = Descrizione
            rsNew!CodiceImEx = IDFamiglia & "-" & CodiceEx
        rsNew.Update
    End If
    
    rsNew.Filter = vbNullString
    
    
Loop

Close #1

rsNew.Close
Set rs = Nothing

Exit Sub
ERR_GET_VARIETA_PRODOTTI:
MsgBox Err.Description, vbCritical, "GET_VARIETA_PRODOTTI"
End Sub
Private Sub INSERIMENTO_PEDANA(IDAzienda As Long, IDFiliale As Long)
On Error GoTo ERR_AVVIA_INSERIMENTO_ARTICOLO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Dim LINK_LISTINO As Long
Dim LINK_UM As Long
Dim LINK_IVA As Long
Dim LINK_TIPO_PRODOTTO As Long
Dim LINK_PEDANA As Long
Dim CODICE_PEDANA As String
Dim DESCRIZIONE_PEDANA As String


LINK_TIPO_PEDANA_PREDEFINITA = 0
LINK_PEDANA = 0


sSQL = "SELECT * FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=0"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic



sSQL = "SELECT * FROM Articolo "
sSQL = sSQL & "WHERE Tipo=0"
Set rs = New ADODB.Recordset
rs.Open sSQL, cnConfig

While Not rs.EOF
    If GET_CONTROLLO_ESISTENZA_ARTICOLO(rs!Codice, IDAzienda) = False Then
        LINK_TIPO_PRODOTTO = GET_LINK_TIPO_PRODOTTO_ARTICOLO(IDAzienda, fnNotNull(rs!TipoProdotto))
        LINK_UM = GET_LINK_UM_ARTICOLO(fnNotNull(rs!UnitaDiMisura))
        LINK_IVA = GET_LINK_IVA_ARTICOLO(fnNotNullN(rs!Iva))
        
        rsNew.AddNew
            rsNew!IDArticolo = fnGetNewKey("Articolo", "IDArticolo")
            rsNew!IDAzienda = IDAzienda
            rsNew!IDIvaAcquisto = LINK_IVA
            rsNew!IDIvaVendita = LINK_IVA
            rsNew!CodiceArticolo = fnNotNull(rs!Codice)
            rsNew!Articolo = fnNotNull(rs!Articolo)
            rsNew!IDUnitaDiMisuraAcquisto = LINK_UM
            rsNew!IDUnitaDiMisuraVendita = LINK_UM
            rsNew!IDTipoProdotto = LINK_TIPO_PRODOTTO
            rsNew!IDUtenteUltimaVariazione = 1
            rsNew!VirtualDelete = 0
            rsNew!DataUltimaVariazione = Date
            rsNew!PesoNetto = fnNotNullN(rs!Peso)
            
        rsNew.Update
        LINK_PEDANA = rsNew!IDArticolo
        CODICE_PEDANA = rsNew!CodiceArticolo
        DESCRIZIONE_PEDANA = rsNew!Articolo
        
    End If
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.Close
Set rs = Nothing

If LINK_PEDANA > 0 Then
    sSQL = "SELECT * FROM RV_POTipoPedana "
    sSQL = sSQL & "WHERE CodiceTipoPedana = " & fnNormString(CODICE_PEDANA)
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_POTipoPedana = fnGetNewKey("RV_POTipoPedana", "IDRV_POTipoPedana")
            rsNew!CodiceTipoPedana = CODICE_PEDANA
            rsNew!TipoPedana = DESCRIZIONE_PEDANA
            rsNew!IDArticoloImballo = LINK_PEDANA
            rsNew!IDFiliale = IDFiliale
            rsNew!IDAzienda = IDAzienda
        rsNew.Update
        LINK_TIPO_PEDANA_PREDEFINITA = rsNew!IDRV_POTipoPedana
    End If
    
    rsNew.Close
    Set rsNew = Nothing
    
    
    
End If

Exit Sub
ERR_AVVIA_INSERIMENTO_ARTICOLO:
    MsgBox Err.Description, vbCritical, "AVVIA_INSERIMENTO_PEDANA"
    
End Sub

Private Function GET_CONTROLLO_ESISTENZA_ARTICOLO(CodiceArticolo As String, IDAzienda As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDArticolo FROM Articolo "
sSQL = sSQL & "WHERE CodiceArticolo=" & fnNormString(CodiceArticolo)
sSQL = sSQL & " AND IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ESISTENZA_ARTICOLO = False
Else
    GET_CONTROLLO_ESISTENZA_ARTICOLO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_UM_ARTICOLO(UM As String) As Long
Dim sSQL As String

Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDUnitaDiMisura FROM UnitaDiMisura "
sSQL = sSQL & "WHERE UnitaDiMisura=" & fnNormString(UM)


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM_ARTICOLO = 0
Else
    GET_LINK_UM_ARTICOLO = fnNotNullN(rs!IDUnitaDiMisura)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_IVA_ARTICOLO(AliquotaIva As Double) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDIva FROM Iva "
sSQL = sSQL & "WHERE AliquotaIva=" & fnNormNumber(AliquotaIva)
sSQL = sSQL & " AND IDIvaDetraibile IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_ARTICOLO = 0
Else
    GET_LINK_IVA_ARTICOLO = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_TIPO_PRODOTTO_ARTICOLO(IDAzienda As Long, TipoProdotto As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDTipoProdotto FROM TipoProdotto "
sSQL = sSQL & "WHERE TipoProdotto=" & fnNormString(TipoProdotto)
sSQL = sSQL & " AND IDAzienda=" & IDAzienda


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_TIPO_PRODOTTO_ARTICOLO = 0
Else
    GET_LINK_TIPO_PRODOTTO_ARTICOLO = fnNotNullN(rs!IDTipoProdotto)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub INSERIMENTO_IMBALLO(IDAzienda As Long, IDFiliale As Long)
On Error GoTo ERR_AVVIA_INSERIMENTO_ARTICOLO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Dim LINK_LISTINO As Long
Dim LINK_UM As Long
Dim LINK_IVA As Long
Dim LINK_TIPO_PRODOTTO As Long



sSQL = "SELECT * FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=0"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic



sSQL = "SELECT * FROM Articolo "
sSQL = sSQL & "WHERE Tipo=1"
Set rs = New ADODB.Recordset
rs.Open sSQL, cnConfig

While Not rs.EOF
    If GET_CONTROLLO_ESISTENZA_ARTICOLO(rs!Codice, IDAzienda) = False Then
        LINK_TIPO_PRODOTTO = GET_LINK_TIPO_PRODOTTO_ARTICOLO(IDAzienda, fnNotNull(rs!TipoProdotto))
        LINK_UM = GET_LINK_UM_ARTICOLO(fnNotNull(rs!UnitaDiMisura))
        LINK_IVA = GET_LINK_IVA_ARTICOLO(fnNotNullN(rs!Iva))
        
        rsNew.AddNew
            rsNew!IDArticolo = fnGetNewKey("Articolo", "IDArticolo")
            rsNew!IDAzienda = IDAzienda
            rsNew!IDIvaAcquisto = LINK_IVA
            rsNew!IDIvaVendita = LINK_IVA
            rsNew!CodiceArticolo = fnNotNull(rs!Codice)
            rsNew!Articolo = fnNotNull(rs!Articolo)
            rsNew!IDUnitaDiMisuraAcquisto = LINK_UM
            rsNew!IDUnitaDiMisuraVendita = LINK_UM
            rsNew!IDTipoProdotto = LINK_TIPO_PRODOTTO
            rsNew!IDUtenteUltimaVariazione = 1
            rsNew!VirtualDelete = 0
            rsNew!DataUltimaVariazione = Date
            rsNew!tara = fnNotNullN(rs!Peso)
            
        rsNew.Update

        
    End If
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.Close
Set rs = Nothing

Exit Sub
ERR_AVVIA_INSERIMENTO_ARTICOLO:
    MsgBox Err.Description, vbCritical, "AVVIA_INSERIMENTO_IMBALLO"
    
End Sub

Private Sub INSERIMENTO_ARTICOLO(IDAzienda As Long, IDFiliale As Long)
On Error GoTo ERR_AVVIA_INSERIMENTO_ARTICOLO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Dim LINK_LISTINO As Long
Dim LINK_UM As Long
Dim LINK_IVA As Long
Dim LINK_TIPO_PRODOTTO As Long
Dim LINK_ARTICOLO_IMBALLO As Long



sSQL = "SELECT * FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=0"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic



sSQL = "SELECT * FROM Articolo "
sSQL = sSQL & "WHERE Tipo=2"
Set rs = New ADODB.Recordset
rs.Open sSQL, cnConfig

While Not rs.EOF
    If GET_CONTROLLO_ESISTENZA_ARTICOLO(rs!Codice, IDAzienda) = False Then
        LINK_TIPO_PRODOTTO = GET_LINK_TIPO_PRODOTTO_ARTICOLO(IDAzienda, fnNotNull(rs!TipoProdotto))
        LINK_UM = GET_LINK_UM_ARTICOLO(fnNotNull(rs!UnitaDiMisura))
        LINK_IVA = GET_LINK_IVA_ARTICOLO(fnNotNullN(rs!Iva))
        LINK_ARTICOLO_IMBALLO = GET_LINK_ARTICOLO_IMBALLO(IDAzienda, fnNotNull(rs!CodiceArticoloImballo))
        rsNew.AddNew
            rsNew!IDArticolo = fnGetNewKey("Articolo", "IDArticolo")
            rsNew!IDAzienda = IDAzienda
            rsNew!IDIvaAcquisto = LINK_IVA
            rsNew!IDIvaVendita = LINK_IVA
            rsNew!CodiceArticolo = fnNotNull(rs!Codice)
            rsNew!Articolo = fnNotNull(rs!Articolo)
            rsNew!IDUnitaDiMisuraAcquisto = LINK_UM
            rsNew!IDUnitaDiMisuraVendita = LINK_UM
            rsNew!IDTipoProdotto = LINK_TIPO_PRODOTTO
            rsNew!IDUtenteUltimaVariazione = 1
            rsNew!VirtualDelete = 0
            rsNew!DataUltimaVariazione = Date
            rsNew!PesoNetto = fnNotNullN(rs!Peso)
            rsNew!RV_POIDImballoConferimento = LINK_ARTICOLO_IMBALLO
            rsNew!RV_POIDImballoVendita = LINK_ARTICOLO_IMBALLO
        rsNew.Update

        
    End If
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.Close
Set rs = Nothing

Exit Sub
ERR_AVVIA_INSERIMENTO_ARTICOLO:
    MsgBox Err.Description, vbCritical, "AVVIA_INSERIMENTO_ARTICOLO"
    
End Sub

Private Function GET_LINK_ARTICOLO_IMBALLO(IDAzienda As Long, CodiceImballo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDArticolo FROM Articolo "
sSQL = sSQL & "WHERE CodiceArticolo=" & fnNormString(CodiceImballo)
sSQL = sSQL & " AND IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ARTICOLO_IMBALLO = 0
Else
    GET_LINK_ARTICOLO_IMBALLO = fnNotNullN(rs!IDArticolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ANAGRAFICA_CLIENTE_DEMO(AnagraficaCliente As String, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_ANAGRAFICA_CLIENTE_DEMO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim objAna As dmtRegAna.CRegAnagrafica
Dim ris As Long

Me.lblInfo.Caption = "CREAZIONE ANAGRAFICA CLIENTE DEMO"
Me.txtControllo.Text = Me.txtControllo.Text & "CREAZIONE ANAGRAFICA CLIENTE DEMO" & vbCrLf
DoEvents

''''''''''''Controllo dell'esistenza dell'anagrafica''''''''''''''''''''''''''''''''
sSQL = "SELECT IDAnagrafica FROM IERepCliente "
sSQL = sSQL & " WHERE IDAzienda=" & VarIDAzienda
sSQL = sSQL & " AND Anagrafica=" & fnNormString(AnagraficaCliente)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_CLIENTE_DEMO = 0
Else
    GET_LINK_ANAGRAFICA_CLIENTE_DEMO = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LINK_ANAGRAFICA_CLIENTE_DEMO > 0 Then Exit Function

'''''''CREAZIONE DELL'ANAGRAFICA'''''''''''''''''''''''''''''''

'Crea un'istanza dell'oggetto CRegAnagrafica
Set objAna = New dmtRegAna.CRegAnagrafica

'Assegna la connessione aperta all'oggetto CReganagrafica
objAna.Connection = CnDMT

objAna.Field "Anagrafica", AnagraficaCliente, "Anagrafica" 'Richiesto

objAna.Field "DataUltimaVariazione", Date, "Anagrafica" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Anagrafica" 'Richiesto
objAna.Field "VirtualDelete", 0, "Anagrafica" 'Richiesto
objAna.Field "PartitaIva", "09748321008", "Anagrafica" 'Richiesto
objAna.Field "CodiceFiscale", "09748321008", "Anagrafica" 'Richiesto
objAna.Field "Indirizzo", "Via orvieto 12", "Anagrafica" 'Richiesto
objAna.Field "IDComune", 5898, "Anagrafica" 'Richiesto
objAna.Field "IDNazione", 110, "Anagrafica" 'Richiesto

'Valorizzare i campi della tabella CLIENTE, utilizzando il metodo Field
objAna.Field "IDAzienda", IDAzienda, "Cliente" 'Richiesto
objAna.Field "IDTipoAnagrafica", 2, "Cliente" 'Richiesto
objAna.Field "DataUltimaVariazione", Date, "Cliente" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Cliente" 'Richiesto
objAna.Field "VirtualDelete", 0, "Cliente" 'Richiesto

ris = objAna.Insert
If ris = 0 Then
    GET_LINK_ANAGRAFICA_CLIENTE_DEMO = fnNotNullN(objAna.Field("IDAnagrafica", , "Anagrafica"))
End If

Exit Function

ERR_GET_LINK_ANAGRAFICA_CLIENTE_DEMO:
    MsgBox Err.Description, vbCritical, "GET_LINK_ANAGRAFICA_CLIENTE_DEMO"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Private Function GET_LINK_ANAGRAFICA_FORNITORE_DEMO(AnagraficaCliente As String, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_ANAGRAFICA_FORNITORE_DEMO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim objAna As dmtRegAna.CRegAnagrafica
Dim ris As Long

Me.lblInfo.Caption = "CREAZIONE ANAGRAFICA FORNITORE DEMO"
Me.txtControllo.Text = Me.txtControllo.Text & "CREAZIONE ANAGRAFICA FORNITORE DEMO" & vbCrLf
DoEvents

''''''''''''Controllo dell'esistenza dell'anagrafica''''''''''''''''''''''''''''''''
sSQL = "SELECT IDAnagrafica FROM IERepCliente "
sSQL = sSQL & " WHERE IDAzienda=" & VarIDAzienda
sSQL = sSQL & " AND Anagrafica=" & fnNormString(AnagraficaCliente)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_FORNITORE_DEMO = 0
Else
    GET_LINK_ANAGRAFICA_FORNITORE_DEMO = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LINK_ANAGRAFICA_FORNITORE_DEMO > 0 Then Exit Function

'''''''CREAZIONE DELL'ANAGRAFICA'''''''''''''''''''''''''''''''

'Crea un'istanza dell'oggetto CRegAnagrafica
Set objAna = New dmtRegAna.CRegAnagrafica

'Assegna la connessione aperta all'oggetto CReganagrafica
objAna.Connection = CnDMT

objAna.Field "Anagrafica", AnagraficaCliente, "Anagrafica" 'Richiesto

objAna.Field "DataUltimaVariazione", Date, "Anagrafica" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Anagrafica" 'Richiesto
objAna.Field "VirtualDelete", 0, "Anagrafica" 'Richiesto
objAna.Field "PartitaIva", "09745805004", "Anagrafica" 'Richiesto
objAna.Field "CodiceFiscale", "09745805004", "Anagrafica" 'Richiesto
objAna.Field "Indirizzo", "Via vinci 30A", "Anagrafica" 'Richiesto
objAna.Field "IDComune", 5898, "Anagrafica" 'Richiesto
objAna.Field "IDNazione", 110, "Anagrafica" 'Richiesto

'Valorizzare i campi della tabella CLIENTE, utilizzando il metodo Field
objAna.Field "IDAzienda", IDAzienda, "Fornitore" 'Richiesto
objAna.Field "IDTipoAnagrafica", 3, "Fornitore" 'Richiesto



objAna.Field "DataUltimaVariazione", Date, "Fornitore" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Fornitore" 'Richiesto
objAna.Field "VirtualDelete", 0, "Fornitore" 'Richiesto

ris = objAna.Insert
If ris = 0 Then
    GET_LINK_ANAGRAFICA_FORNITORE_DEMO = fnNotNullN(objAna.Field("IDAnagrafica", , "Anagrafica"))
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Function
ERR_GET_LINK_ANAGRAFICA_FORNITORE_DEMO:
    MsgBox Err.Description, vbCritical, "GET_LINK_ANAGRAFICA_FORNITORE_DEMO"
    GET_LINK_ANAGRAFICA_FORNITORE_DEMO = 0
End Function


Private Function GET_LINK_ANAGRAFICA_SOCIO_DEMO(AnagraficaCliente As String, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_ANAGRAFICA_SOCIO_DEMO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim objAna As dmtRegAna.CRegAnagrafica
Dim ris As Long

Me.lblInfo.Caption = "CREAZIONE ANAGRAFICA SOCIO DEMO"
Me.txtControllo.Text = Me.txtControllo.Text & "CREAZIONE ANAGRAFICA SOCIO DEMO" & vbCrLf
DoEvents

''''''''''''Controllo dell'esistenza dell'anagrafica''''''''''''''''''''''''''''''''
sSQL = "SELECT IDAnagrafica FROM IERepCliente "
sSQL = sSQL & " WHERE IDAzienda=" & VarIDAzienda
sSQL = sSQL & " AND Anagrafica=" & fnNormString(AnagraficaCliente)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_SOCIO_DEMO = 0
Else
    GET_LINK_ANAGRAFICA_SOCIO_DEMO = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LINK_ANAGRAFICA_SOCIO_DEMO > 0 Then Exit Function

'''''''CREAZIONE DELL'ANAGRAFICA'''''''''''''''''''''''''''''''

'Crea un'istanza dell'oggetto CRegAnagrafica
Set objAna = New dmtRegAna.CRegAnagrafica

'Assegna la connessione aperta all'oggetto CReganagrafica
objAna.Connection = CnDMT

objAna.Field "Anagrafica", AnagraficaCliente, "Anagrafica" 'Richiesto

objAna.Field "DataUltimaVariazione", Date, "Anagrafica" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Anagrafica" 'Richiesto
objAna.Field "VirtualDelete", 0, "Anagrafica" 'Richiesto
objAna.Field "PartitaIva", "05875972311", "Anagrafica" 'Richiesto
objAna.Field "CodiceFiscale", "05875972311", "Anagrafica" 'Richiesto
objAna.Field "Indirizzo", "Via della terra rossa 56", "Anagrafica" 'Richiesto
objAna.Field "IDComune", 5415, "Anagrafica" 'Richiesto
objAna.Field "IDNazione", 110, "Anagrafica" 'Richiesto
objAna.Field "IDCategoriaAnagrafica", GET_LINK_CATEGORIA_ANAGRAFICA_SOCIO("Socio"), "Anagrafica" 'Richiesto

'Valorizzare i campi della tabella CLIENTE, utilizzando il metodo Field
objAna.Field "IDAzienda", IDAzienda, "Fornitore" 'Richiesto
objAna.Field "IDTipoAnagrafica", 3, "Fornitore" 'Richiesto


objAna.Field "DataUltimaVariazione", Date, "Fornitore" 'Richiesto
objAna.Field "IDUtenteUltimaVariazione", 1, "Fornitore" 'Richiesto
objAna.Field "VirtualDelete", 0, "Fornitore" 'Richiesto

ris = objAna.Insert
If ris = 0 Then
    GET_LINK_ANAGRAFICA_SOCIO_DEMO = fnNotNullN(objAna.Field("IDAnagrafica", , "Anagrafica"))
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Function
ERR_GET_LINK_ANAGRAFICA_SOCIO_DEMO:
    MsgBox Err.Description, vbCritical, "GET_LINK_ANAGRAFICA_SOCIO_DEMO"
    GET_LINK_ANAGRAFICA_SOCIO_DEMO = 0
End Function


Private Function GET_LINK_CATEGORIA_ANAGRAFICA_SOCIO(CategoriaAnagrafica As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM CategoriaAnagrafica "
sSQL = sSQL & " WHERE CategoriaAnagrafica=" & fnNormString(CategoriaAnagrafica)
sSQL = sSQL & " AND IDTipoAnagrafica IS NULL"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CATEGORIA_ANAGRAFICA_SOCIO = 0
Else
    GET_LINK_CATEGORIA_ANAGRAFICA_SOCIO = fnNotNullN(rs!IDCategoriaAnagrafica)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Sub CICLO_BIOLOGICO()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "CICLO BIOLOGICO"
Me.txtControllo.Text = Me.txtControllo.Text & "CICLO BIOLOGICO" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM CicloBiologico "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PO01_CicloBiologico "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!CicloBiologico) & vbCrLf
    DoEvents
    rsNew.Filter = "CicloBiologico=" & fnNormString(rsArc!CicloBiologico)
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_CicloBiologico = fnGetNewKey("RV_PO01_CicloBiologico", "IDRV_PO01_CicloBiologico")
            rsNew!CicloBiologico = fnNotNull(rsArc!CicloBiologico)
        rsNew.Update
    End If
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "CICLO_BIOLOGICO"
    
End Sub

Private Sub STATO_LOTTO()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "STATO LOTTO"
Me.txtControllo.Text = Me.txtControllo.Text & "STATO LOTTO" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM StatoLotto "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PO01_StatoLotto"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!StatoLotto) & vbCrLf
    DoEvents
    rsNew.Filter = "StatoLotto=" & fnNormString(rsArc!StatoLotto)
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_StatoLotto = fnGetNewKey("RV_PO01_StatoLotto", "IDRV_PO01_StatoLotto")
            rsNew!StatoLotto = fnNotNull(rsArc!StatoLotto)
            rsNew!CodiceImEx = rsNew!IDRV_PO01_StatoLotto
        rsNew.Update
    End If
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "STATO_LOTTO"
    
End Sub

Private Sub TIPO_PRODUZIONE()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "TIPO PRODUZIONE"
Me.txtControllo.Text = Me.txtControllo.Text & "TIPO PRODUZIONE" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM TipoProduzione "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PO01_TipoProduzione"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!TipoProduzione) & vbCrLf
    DoEvents
    rsNew.Filter = "TipoProduzione=" & fnNormString(rsArc!TipoProduzione)
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_TipoProduzione = fnGetNewKey("RV_PO01_TipoProduzione", "IDRV_PO01_TipoProduzione")
            rsNew!TipoProduzione = fnNotNull(rsArc!TipoProduzione)
            rsNew!CodiceImEx = rsNew!IDRV_PO01_TipoProduzione
        rsNew.Update
    End If
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "TIPO_PRODUZIONE"
    
End Sub

Private Sub TITOLO_TERRENO()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "TITOLO TERRENO"
Me.txtControllo.Text = Me.txtControllo.Text & "TITOLO TERRENO" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM Titolo "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PO01_Titolo"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!Titolo) & vbCrLf
    DoEvents
    rsNew.Filter = "Titolo=" & fnNormString(rsArc!Titolo)
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_Titolo = fnGetNewKey("RV_PO01_Titolo", "IDRV_PO01_Titolo")
            rsNew!Titolo = fnNotNull(rsArc!Titolo)
        rsNew.Update
    End If
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "TITOLO TERRENO"
    
End Sub


Private Sub TIPO_OPERAZIONE()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "TIPO OPERAZIONE"
Me.txtControllo.Text = Me.txtControllo.Text & "TIPO OPERAZIONE" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM TipoOperazione "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PO01_TipoOperazione"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!TipoOperazione) & vbCrLf
    DoEvents
    rsNew.Filter = "TipoOperazione=" & fnNormString(rsArc!TipoOperazione)
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_TipoOperazione = fnGetNewKey("RV_PO01_TipoOperazione", "IDRV_PO01_TipoOperazione")
            rsNew!TipoOperazione = fnNotNull(rsArc!TipoOperazione)
        rsNew.Update
    End If
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "TIPO_OPERAZIONE"
    
End Sub



Private Sub STRUTTURA_SERRA()
On Error GoTo ERR_UNITA_DI_MISURA_COOP

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "STRUTTURA SERRA"
Me.txtControllo.Text = Me.txtControllo.Text & "STRUTTURA SERRA" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO PRODOTTO ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM Struttura "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_PO01_Struttura"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!Struttura) & vbCrLf
    DoEvents
    rsNew.Filter = "Struttura=" & fnNormString(rsArc!Struttura)
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_PO01_Struttura = fnGetNewKey("RV_PO01_Struttura", "IDRV_PO01_Struttura")
            rsNew!Struttura = fnNotNull(rsArc!Struttura)
            rsNew!StrutturaSerra = fnNotNullN(rsArc!StrutturaSerra)
            
        rsNew.Update
    End If
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_UNITA_DI_MISURA_COOP:
    MsgBox Err.Description, vbCritical, "STRUTTURA SERRA"
    
End Sub


Private Sub IDENTIFICATORI_EAN128()
On Error GoTo ERR_IDENTIFICATORI_EAN128

Dim sSQL As String
Dim rsArc As ADODB.Recordset
Dim rsNew As ADODB.Recordset

Me.lblInfo.Caption = "IDENTIFICATORI EAN128"
Me.txtControllo.Text = Me.txtControllo.Text & "IDENTIFICATORI EAN128" & vbCrLf
DoEvents

''''''''RECUPERO DATI TIPO IDENTIFICATORI EAN128 ARCHIVIO'''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIdentificatoreEAN128 "
Set rsArc = New ADODB.Recordset
rsArc.Open sSQL, cnConfig
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POIdentificatoreEAN128 "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsArc.EOF
    Me.txtControllo.Text = Me.txtControllo.Text & fnNotNull(rsArc!Identificatore) & vbCrLf
    DoEvents
    rsNew.Filter = "Identificatore=" & fnNormString(rsArc!Identificatore)
    If rsNew.EOF Then
        rsNew.AddNew
            rsNew!IDRV_POIdentificatoreEAN128 = fnGetNewKey("RV_POIdentificatoreEAN128", "IDRV_POIdentificatoreEAN128")
            rsNew!Identificatore = fnNotNull(rsArc!Identificatore)
            rsNew!Descrizione = fnNotNull(rsArc!Descrizione)
            
            rsNew!Lunghezza = fnNotNull(rsArc!Lunghezza)
            rsNew!Fisso = fnNotNull(rsArc!Fisso)
            rsNew!CalcolaCheck = fnNotNull(rsArc!CalcolaCheck)
            rsNew!Alfanumerico = fnNotNull(rsArc!Alfanumerico)
            rsNew!CampoData = fnNotNull(rsArc!CampoData)
            rsNew!CampoConDecimali = fnNotNull(rsArc!CampoConDecimali)
        rsNew.Update
    End If
    rsNew.Filter = vbNullString

rsArc.MoveNext
Wend
rsNew.Close
Set rsNew = Nothing

rsArc.Close
Set rsArc = Nothing
Exit Sub
ERR_IDENTIFICATORI_EAN128:
    MsgBox Err.Description, vbCritical, "IDENTIFICATORI_EAN128"
    
End Sub
Private Sub ELIMINA_FORMULA_QUANTITA()
On Error GoTo ERR_ELIMINA_FORMULA_QUANTITA
Dim sSQL As String


sSQL = "UPDATE CampoDiamante SET "
sSQL = sSQL & "Formula2=NULL "
sSQL = sSQL & "WHERE CampoDiamante=" & fnNormString("Art_quantita_totale")

CnDMT.Execute sSQL


Exit Sub
ERR_ELIMINA_FORMULA_QUANTITA:
    MsgBox Err.Description, vbCritical, "ELIMINA_FORMULA_QUANTITA"
End Sub
