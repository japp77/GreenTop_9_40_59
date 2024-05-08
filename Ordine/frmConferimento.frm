VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConferimento 
   Caption         =   "COLLEGAMENTO AL CONFERIMENTO"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14055
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConferimento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEla 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   7080
      Width           =   14055
      Begin VB.CommandButton Command1 
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
         Height          =   495
         Left            =   11640
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblInfoConf 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   11415
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   11415
      End
   End
   Begin DmtGridCtl.DmtGrid GrigliaCorpo 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12303
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableMove      =   0   'False
      ColumnsHeaderHeight=   20
   End
End
Attribute VB_Name = "frmConferimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset
Private NumeroRecord As Long


'''''''''''''''''''VARIABILI GLOBALI APPLICAZIONE CONFERIMENTO MERCE''''''''''''''''''''''''''
Public Link_Magazzino_Conferimento As Long
Public Link_Magazzino_Vendita As Long
Public Link_Esercizio As Long
Public Link_PeriodoIVA As Long
Public Link_Sezionale As Long

Public Link_TipoImballo As Long
'Public Link_TipoSocio As Long
Public Link_TipoGrezzo As Long
Public Link_TipoLavorato As Long

'Variabili che identificano un tipo prodotto di scarti e cali pesi
Public Link_TipoScarto As Long
Public Link_TipoCaloPeso As Long
Public Link_TipoAumentoPeso As Long


Public Flag_GestioneArticoli As Boolean
Public Flag_AutomazioneLotti As Boolean


Public Link_Causale_MagCar_Conf As Long
Public Link_Causale_MagScar_Conf As Long
Public Link_Causale_MagCar_Vend As Long
Public Link_Causale_MagScar_Vend As Long
Public Link_UMAcq As Long
Public NumeroDocumentoDisponibile As Long
Public Link_Oggetto As Long
Public DataInizio_Esercizio As String
Public DataFine_Esercizio As String
Public Numero_Decimali_Pesi As Long

Public Link_LottoArticolo As Long
Public Link_Movimento_Carico As Long
Public Link_Movimento_Scarico As Long
Public Link_Movimento_Carico_Imballo As Long

Public Link_UnitaDiMisura_Acquisto As Long
Public Link_UnitaDiMisura_Coop As Long



Public Link_Magazzino_Carico As Long
Public Link_Magazzino_Scarico As Long
Public Link_CausaleCarico As Long
Public Link_CausaleScarico As Long


Const IDDocumento As Long = 4
Public EsistenzaValoreSezionale As Boolean

Public LINK_REGIONE_SOCIO As Long
Public LINK_NAZIONE_SOCIO As Long
Public LINK_COMUNE_SOCIO As Long
Public LINK_PROVINCIA_SOCIO As Long

Private Mov As DmtMovim.cMovimentazione

Private IDTestataConf As Long
Private IDOggettoTestataConf As Long
Private IDTipoOggettoConf As Long
Private IDAnagraficaSocioConf As Long
Private DataDocumentoConf As String
Private NumeroDocumentoConf As Long
Private AnagraficaSocioConf As String
Private NomeSocioConf As String
Private CodiceSocioConf As String
Private DescrizioneFunzioneConf As String

Private DataConf As String
Private NumeroConf As Long
Private IDRigaConf As Long
Private CodiceLottoEntrataConf As String
Private IDArticoloConf As Long
Private IDUMConf As String
Private ArticoloConf As String


Private Sub CREA_RECORDSET()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long

If Not (rsGriglia Is Nothing) Then
    If rsGriglia.State > 0 Then
        rsGriglia.Close
    End If
    
    Set rsGriglia = Nothing
End If

sSQL = "SELECT * FROM RV_POIERigheOrdineConferimento "
sSQL = sSQL & "WHERE IDOggetto=0"

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

For I = 0 To rs.Fields.Count - 1
    Select Case rs.Fields(I).Type
    
        Case adChar, adVarChar, adVarWChar, adWChar
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, rs.Fields(I).DefinedSize, rs.Fields(I).Attributes
        
        Case adInteger
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes

        
        Case adDate, adDBDate, adDBTime, adDBTimeStamp
            rsGriglia.Fields.Append rs.Fields(I).Name, rs.Fields(I).Type, , rs.Fields(I).Attributes

        
        Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
            rsGriglia.Fields.Append rs.Fields(I).Name, adBoolean, , rs.Fields(I).Attributes
       
        Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
            rsGriglia.Fields.Append rs.Fields(I).Name, adDouble, , rs.Fields(I).Attributes
            
    End Select
Next

rsGriglia.Fields.Append "Registra", adBoolean, , adFldIsNullable

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic

rs.Close
Set rs = Nothing


sSQL = "SELECT * FROM RV_POIERigheOrdineConferimento "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND RV_POTipoRiga=1"
sSQL = sSQL & " AND IDAnagraficaSocioOrdine>0"

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF

    If (fnNotNullN(rs!IDCategoriaAnagrafica) = Link_TipoSocio) Then
        rsGriglia.AddNew
            For I = 0 To rs.Fields.Count - 1
                rsGriglia.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
            Next
            rsGriglia!Registra = 0
            If (fnNotNullN(rs!IDValoriOggettoDettaglioOrdineConf) = 0) Then
                rsGriglia!Registra = 1
            End If
        rsGriglia.Update
        NumeroRecord = NumeroRecord + 1
    End If
    
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

GET_GRIGLIA

End Sub

Private Sub Command1_Click()
On Error GoTo ERR_Command1_Click
Dim IDConferimento As Long
Dim IDConferimentoRiga As Long
Dim IDLavorazione As Long

If NumeroRecord = 0 Then Exit Sub


Me.Command1.Enabled = False

rsGriglia.Update

rsGriglia.MoveFirst

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = NumeroRecord

While Not rsGriglia.EOF
    
    If (Me.ProgressBar1.Value + 1) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
    End If
    
    lblInfo.Caption = "Elaborazione " & Me.ProgressBar1.Value & " di " & NumeroRecord
    DoEvents
    
    If rsGriglia!Registra = True Then
        If fnNotNullN(rsGriglia!IDValoriOggettoDettaglioOrdineConf) = 0 Then
            IDConferimento = CREA_CONFERIMENTO(rsGriglia!IDAnagraficaSocioOrdine, Date)
            
            If IDConferimento > 0 Then
                
                IDConferimentoRiga = CREA_RIGA_CONFERIMENTO(IDConferimento, fnNotNullN(rsGriglia!IDValoriOggettoDettaglio))
                
                If (IDConferimentoRiga > 0) Then
                    CREA_RIGA_DI_LAVORAZIONE fnNotNullN(rsGriglia!IDValoriOggettoDettaglio), IDConferimentoRiga, IDConferimento
                End If
                
            End If
        End If
    End If
    
rsGriglia.MoveNext
Wend

CREA_RECORDSET

Me.lblInfo.Caption = ""
Me.lblInfoConf.Caption = "OPERAZIONE AVVENUTA CON SUCCESSO"
Me.ProgressBar1.Value = 0

MsgBox "OPERAZIONE AVVENUTA CON SUCCESSO!", vbInformation, "Collegamento ordine a conferimento"

Me.Command1.Enabled = True

Exit Sub
ERR_Command1_Click:
    MsgBox Err.Description, vbCritical, "Command1_Click"
    Me.Command1.Enabled = False
End Sub

Private Sub Form_Load()
    NumeroRecord = 0
    CREA_RECORDSET
    
End Sub

Private Sub GET_GRIGLIA()

OLDCursor = Cn.CursorLocation
Cn.CursorLocation = 3
    
    With Me.GrigliaCorpo
        .EnableMove = True
        .UpdatePosition = True
        .BooleanType = dgGraphic
        .SelectionMode = dgSelectCell
        .ColumnsHeader.Clear
        .LoadUserSettings
            
            Set cl = .ColumnsHeader.Add("Registra", "Registra", dgBoolean, True, 1500, dgAligncenter)
                cl.Editable = True
            
            .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Link_art_articolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "Art_codice", "Codice articolo", dgchar, True, 1100, dgAlignleft
            .ColumnsHeader.Add "Art_descrizione", "Descrizione articolo", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "Link_Art_unita_di_misura", "Link_Art_unita_di_misura", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "UnitaDiMisuraVendita", "U.M. ordine", dgchar, True, 1100, dgAlignleft
            Set cl = .ColumnsHeader.Add("Art_quantita_totale", "Q.tà ordine", dgDouble, True, 1300, dgAlignRight)
                cl.Editable = False
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            .ColumnsHeader.Add "IDRV_POCaricoMerceRigheCollegato", "IDRV_POCaricoMerceRigheCollegato", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDRV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", dgNumeric, False, 500, dgAlignleft

            .ColumnsHeader.Add "NumeroConferimento", "N° conf.", dgInteger, True, 1000, dgAlignRight
            .ColumnsHeader.Add "DataConferimento", "Data conf.", dgDate, True, 1500, dgAlignleft
            .ColumnsHeader.Add "AnagraficaSocio", "Socio", dgchar, True, 1500, dgAlignleft
            .ColumnsHeader.Add "NomeSocio", "Nome socio", dgchar, True, 1000, dgAlignleft

            .ColumnsHeader.Add "IDRV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDArticolo", "IDArticolo", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "CodiceArticolo", "Codice articolo conf.", dgchar, True, 1100, dgAlignleft
            .ColumnsHeader.Add "Articolo", "Descrizione articolo conf.", dgchar, True, 3000, dgAlignleft
            .ColumnsHeader.Add "IDUnitaDiMisuraConferimento", "IDUnitaDiMisuraConferimento", dgNumeric, False, 500, dgAlignleft
            .ColumnsHeader.Add "UnitaDiMisuraConferimento", "U.M. conf.", dgchar, True, 1100, dgAlignleft
            Set cl = .ColumnsHeader.Add("Qta_UM", "Q.tà conf.", dgDouble, True, 1300, dgAlignRight)
                cl.Editable = False
                cl.BackColor = vbYellow
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericCurSymbol = ""
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
      
                    
        Set .Recordset = rsGriglia
        .Refresh
    End With

Cn.CursorLocation = OLDCursor

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.GrigliaCorpo.Width = Me.Width - 240
    
    Me.fraEla.Width = Me.Width - 240
    Me.fraEla.Top = Me.Height - 640 - Me.fraEla.Height
    
    Me.GrigliaCorpo.Height = Me.fraEla.Top - 105
    
    Me.Command1.Left = Me.fraEla.Width - 240 - Me.Command1.Width
    Me.ProgressBar1.Width = Me.Command1.Left - 640
    Me.lblInfo.Width = Me.ProgressBar1.Width
    Me.lblInfoConf.Width = Me.ProgressBar1.Width
    
End Sub
Private Sub GrigliaCorpo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Nel caso in cui l'utente clicca con il mouse sulla DmtGrid
    'viene intercettata la posizione del cursore per capire se l'utente ha
    'cliccato una riga in corrispondenza della colonna "Selezionato"
    
    'Controlla se l'utente ha cliccato su una riga valida
    If GrigliaCorpo.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaCorpo.ColumnsHeader(1).Width Then
            'Se non siamo in modalità filtri
            If GrigliaCorpo.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsGriglia.Fields("Registra").Value), 2
            End If
        End If
    End If
End Sub

Private Sub GrigliaCorpo_KeyPress(KeyAscii As Integer)
    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If GrigliaCorpo.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGriglia.Fields("Registra").Value), 2
        End If
    End If
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
    
    If Not rsGriglia.EOF And Not rsGriglia.BOF Then
        rsGriglia.Fields("Registra").Value = Abs(CLng(Selected))
        'sbCheckSelected
        Me.GrigliaCorpo.Refresh
    End If
End Sub

Private Function CREA_CONFERIMENTO(IDAnagraficaSocio As Long, DataDocumento As String) As Long
'On Error GoTo ERR_CREA_CONFERIMENTO
Dim sSQL As String
Dim rsNew As ADODB.Recordset

    lblInfoConf.Caption = "Creazione conferimento..."
    DoEvents


    CREA_CONFERIMENTO = 0
    
    ParametriDiDefault DataDocumento
    
    sSQL = "SELECT * FROM RV_POCaricoMerceTesta "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=0"
    
    Set rsNew = New ADODB.Recordset
    
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    rsNew.AddNew
    
        rsNew!IDRV_POcaricoMercetesta = fnGetNewKey("RV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta")
        rsNew!IDSezionale = Link_Sezionale
        rsNew!IDEsercizio = Link_Esercizio
        rsNew!IDTipoDocumentoCoop = 1
        rsNew!NumeroDocumento = fnGetNumeroDocumento
        rsNew!DataDocumento = DataDocumento
        rsNew!IDOggetto = GET_LINK_OGGETTO(rsNew!DataDocumento, rsNew!NumeroDocumento)
        rsNew!IDMagazzinoConferimento = Link_Magazzino_Conferimento
        rsNew!IDMagazzinoVendita = Link_Magazzino_Vendita
        rsNew!IDAzienda = TheApp.IDFirm
        rsNew!IDFiliale = TheApp.Branch
        rsNew!IDAnagrafica = IDAnagraficaSocio
        
        GET_PROPRIETA_ANAGRAFICA_SOCIO rsNew
        
        rsNew!Annotazioni = ""
        rsNew!Annotazioni1 = ""
        rsNew!Annotazioni2 = ""
        rsNew!Annotazioni3 = ""
        rsNew!IDTipoOggetto = fnGetTipoOggetto("RV_POCaricoMerceL")
        rsNew!PrefissoNumeroConferimento = GET_PARAMETRI_SOCIO(Link_Esercizio, IDAnagraficaSocio, "PrefissoNumeroConferimento")
        rsNew!NumeroDocumentoSocio = GET_PARAMETRI_SOCIO(Link_Esercizio, IDAnagraficaSocio, "NumeroConferimento")
    
        rsNew!DataConferimento = DataDocumento
        rsNew!IDAnagraficaFatturazione = IDAnagraficaSocio
        rsNew!IDVettore = 0
        rsNew!IDLuogoPresaMerce = 0
        rsNew!IDLetteraIntento = 0
        rsNew!ImportoTrasporto = 0
        rsNew!IDRV_POTipoOrdineConforme = 0
        rsNew!IDTipoDocumentoAcq = 0
        rsNew!NumeroDocumentoAcq = oDoc.Field("Doc_numero", , sTabellaTestata)
        rsNew!DataDocumentoAcq = oDoc.Field("Doc_data", , sTabellaTestata)
        
    rsNew.Update
    
    CREA_CONFERIMENTO = rsNew!IDRV_POcaricoMercetesta
    AggiornamentoProgressivoSezionale rsNew!NumeroDocumento
    AGGIORNA_PROGRESSIVO_CONFERIMENTO rsNew!IDAnagrafica, Link_Esercizio, rsNew!NumeroDocumentoSocio, rsNew!PrefissoNumeroConferimento
    
    IDTestataConf = rsNew!IDRV_POcaricoMercetesta
    IDOggettoTestataConf = rsNew!IDOggetto
    IDTipoOggettoConf = rsNew!IDTipoOggetto
    IDAnagraficaSocioConf = rsNew!IDAnagrafica
    DataDocumentoConf = rsNew!DataDocumento
    NumeroDocumentoConf = rsNew!NumeroDocumento
    AnagraficaSocioConf = rsNew!Anagrafica
    NomeSocioConf = rsNew!Nome
    CodiceSocioConf = rsNew!CodiceSocio
    DescrizioneFunzioneConf = GET_DESCRIZIONE_FUNZIONE(fnGetTipoOggetto("RV_POCaricoMerceL"))
    DataConf = rsNew!DataDocumento
    NumeroConf = rsNew!NumeroDocumento
    
    
    
    rsNew.Close
    Set rsNew = Nothing
    
Exit Function
ERR_CREA_CONFERIMENTO:
    MsgBox Err.Description, vbCritical, "CREA_CONFERIMENTO"
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
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub ParametriDiDefault(DataDocumento As String)
    
    Link_Esercizio = fnGetEsercizio(DataDocumento)
    Link_PeriodoIVA = fnGetPeriodoIVA(DataDocumento)
    Link_Sezionale = fnGetSezionale(IDDocumento)
    
    Link_Magazzino_Conferimento = fnGetParametriMagazzino("IDMagazzino_Carico")
    Link_Magazzino_Vendita = fnGetParametriMagazzino("IDMagazzino_Vendita")
    Link_Causale_MagCar_Conf = fnGetParametriMagazzino("IDCausale_Carico_Mag_Carico")
    Link_Causale_MagScar_Conf = fnGetParametriMagazzino("IDCausale_Scarico_Mag_Carico")
    Link_Causale_MagCar_Vend = fnGetParametriMagazzino("IDCausale_Carico_Mag_Vendita")
    Link_Causale_MagScar_Vend = fnGetParametriMagazzino("IDCausale_Scarico_Mag_vendita")

    
End Sub


Public Function fnGetSezionale(Link_DocumentoCoop) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT RV_POSezionalePerDocumento.IDSezionale "
    sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
    sSQL = sSQL & "RV_POSezionalePerDocumento ON RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POSezionalePerDocumento.IDRV_POSchemaCoop "
    sSQL = sSQL & "WHERE ((IDFiliale=" & TheApp.Branch & ") "
    sSQL = sSQL & "AND (IDDocumentoCoop=" & Link_DocumentoCoop & ") "
    sSQL = sSQL & "AND (Predefinito=" & fnNormBoolean(1) & "))"
    
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetSezionale = rsEse!IDSezionale
    Else
        fnGetSezionale = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Public Function fnGetParametriMagazzino(NomeCampo As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
    sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetParametriMagazzino = rsEse.adoColumns(NomeCampo).Value
    Else
        sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
        sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
        
        Set rsEse = Cn.OpenResultset(sSQL)
        
        If rsEse.EOF = False Then
            fnGetParametriMagazzino = rsEse.adoColumns(NomeCampo).Value
        Else
            fnGetParametriMagazzino = 0
        End If
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Public Function fnGetNumeroDocumento()
Dim rsEse As DmtOleDbLib.adoResultset
Dim sSQL As String
    
    sSQL = "SELECT ProgressivoDisponibile FROM ProgressivoSezionale "
    sSQL = sSQL & "WHERE ((IDPeriodoIva=" & Link_PeriodoIVA & ") "
    sSQL = sSQL & "AND (IDSezionale=" & Link_Sezionale & ") "
    sSQL = sSQL & "AND (IDTipoModulo=1))"
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        EsistenzaValoreSezionale = True
        NumeroDocumentoDisponibile = rsEse!ProgressivoDisponibile
        fnGetNumeroDocumento = rsEse!ProgressivoDisponibile
    Else
        EsistenzaValoreSezionale = False
        NumeroDocumentoDisponibile = 1
        fnGetNumeroDocumento = 1
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing

End Function
Public Function fnGetPeriodoIVA(dData As String) As Long
'On Error GoTo ERR_fnGetPeriodoIVA
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDPeriodoIVA FROM PeriodoIVA WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (Anno = " & DatePart("yyyy", dData) & "))"
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetPeriodoIVA = rsEse!IDPeriodoIVA
    Else
        fnGetPeriodoIVA = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
Exit Function
ERR_fnGetPeriodoIVA:
    MsgBox Err.Description, vbCritical, "Periodo Iva"
End Function
Public Function fnGetEsercizio(dData As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDEsercizio, DataInizio, DataFine FROM Esercizio WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (DataInizio <= " & fnNormDate(dData) & ")"
    sSQL = sSQL & " AND (DataFine >= " & fnNormDate(dData) & "))"

    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetEsercizio = rsEse!IDEsercizio
        DataInizio_Esercizio = rsEse!DataInizio
        DataFine_Esercizio = rsEse!DataFine
    Else
        fnGetEsercizio = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Private Function GET_LINK_OGGETTO(DataDocumento As String, NumeroDocumento As Long)
On Error GoTo ERR_GET_LINK_OGGETTO
Dim IDOggetto As Long
Dim sSQL As String

    IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
    
    sSQL = "INSERT INTO Oggetto ("
    sSQL = sSQL & "IDOggetto, IDTipoOggetto, IDAzienda, IDAttivitaAzienda, IDSezionale, "
    sSQL = sSQL & "Oggetto, DataEmissione, Numero, DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete, IDFunzione)"
    sSQL = sSQL & " VALUES ("
    sSQL = sSQL & IDOggetto & ", "
    sSQL = sSQL & fnGetTipoOggetto("RV_POCaricoMerceL") & ", "
    sSQL = sSQL & TheApp.IDFirm & ", "
    sSQL = sSQL & GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch) & ", "
    sSQL = sSQL & Link_Sezionale & ", "
    sSQL = sSQL & fnNormString(GET_DESCRIZIONE_FUNZIONE(fnGetTipoOggetto("RV_POCaricoMerceL"))) & ", "
    sSQL = sSQL & fnNormDate(DataDocumento) & ", "
    sSQL = sSQL & fnNormNumber(NumeroDocumento) & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & TheApp.IDUser & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & TheApp.FunctionID & ")"
    
    Cn.Execute sSQL
    GET_LINK_OGGETTO = IDOggetto

Exit Function
ERR_GET_LINK_OGGETTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_OGGETTO"
End Function
Private Function GetAttivitaAzienda(IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivitaAzienda.IDAttivitaAzienda, Azienda.IDAzienda, Filiale.IDFiliale "
sSQL = sSQL & "FROM AttivitaAzienda INNER JOIN "
sSQL = sSQL & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda INNER JOIN "
sSQL = sSQL & "Filiale ON AttivitaAzienda.IDAttivitaAzienda = Filiale.IDAttivitaAzienda "
sSQL = sSQL & "Where (Azienda.IDAzienda =" & IDAzienda & ") And (Filiale.IDFiliale = " & IDFiliale & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GetAttivitaAzienda = 0
Else
    GetAttivitaAzienda = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_DESCRIZIONE_FUNZIONE(IDTipoOggetto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione, Funzione "
sSQL = sSQL & "FROM Funzione "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_FUNZIONE = ""
Else
    GET_DESCRIZIONE_FUNZIONE = fnNotNull(rs!Funzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_PROPRIETA_ANAGRAFICA_SOCIO(rsnewTmp As ADODB.Recordset)
Dim rs As ADODB.Recordset
Dim sSQL As String

    

sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.Cap, "
sSQL = sSQL & "Comune.Comune, Provincia.Provincia, Anagrafica.Indirizzo, "
sSQL = sSQL & "Fornitore.IDAzienda , Fornitore.IDCategoriaAnagrafica, Fornitore.Codice, "
sSQL = sSQL & "Anagrafica.IDComune, Anagrafica.IDNazione, Comune.IDProvincia, Provincia.IDRegione "
sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
sSQL = sSQL & "Anagrafica INNER JOIN "
sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica ON dbo.Comune.IDComune = dbo.Anagrafica.IDComune "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & "AND Anagrafica.IDAnagrafica=" & rsnewTmp!IDAnagrafica

    
Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

If rs.EOF = False Then

    rsnewTmp!Anagrafica = fnNotNull(rs!Anagrafica)
    rsnewTmp!Nome = fnNotNull(rs!Nome)
    rsnewTmp!CodiceSocio = fnNotNull(rs!Codice)
    rsnewTmp!Indirizzo = fnNotNull(rs!Indirizzo)
    rsnewTmp!Comune = fnNotNull(rs!Comune)
    rsnewTmp!Cap = fnNotNull(rs!Cap)
    rsnewTmp!Provincia = fnNotNull(rs!Provincia)
    
    LINK_REGIONE_SOCIO = fnNotNullN(rs!IDRegione)
    LINK_NAZIONE_SOCIO = fnNotNullN(rs!IDNazione)
    LINK_COMUNE_SOCIO = fnNotNullN(rs!IDComune)
    LINK_PROVINCIA_SOCIO = fnNotNullN(rs!IDProvincia)
End If
rs.Close
Set rs = Nothing



End Sub
Private Function GET_PARAMETRI_SOCIO(IDEsercizio As Long, IDAnagrafica As Long, NomeCampo As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_PONumerazionePerSocio "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizio
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    If rs.adoColumns(NomeCampo).DataType = 4 Then
        GET_PARAMETRI_SOCIO = 1
    Else
        GET_PARAMETRI_SOCIO = ""
    End If

Else
    If rs.adoColumns(NomeCampo).DataType = 4 Then
        GET_PARAMETRI_SOCIO = fnNotNullN(rs.adoColumns(NomeCampo).Value)
        If GET_PARAMETRI_SOCIO = 0 Then
            GET_PARAMETRI_SOCIO = 1
        End If
    Else
        GET_PARAMETRI_SOCIO = fnNotNull(rs.adoColumns(NomeCampo).Value)
    End If

    
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_PROGRESSIVO_CONFERIMENTO(IDAnagrafica As Long, IDEsercizio As Long, NumeroConferimento As Long, PrefissoNumeroConferimento As String)
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_PONumerazionePerSocio "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizio
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rs.EOF Then
    If NumeroConferimento >= fnNotNullN(rs!NumeroConferimento) Then
        rs!NumeroConferimento = NumeroConferimento + 1
        rs!PrefissoNumeroConferimento = PrefissoNumeroConferimento
        rs.Update
    End If
Else
    rs.AddNew
        rs!IDRV_PONumerazionePerSocio = fnGetNewKey("RV_PONumerazionePerSocio", "IDRV_PONumerazionePerSocio")
        rs!IDAnagrafica = IDAnagrafica
        rs!IDEsercizio = IDEsercizio
        rs!IDAzienda = TheApp.IDFirm
        rs!Numero = 1
        rs!Prefisso = ""
        rs!NumeroConferimento = NumeroConferimento + 1
        rs!PrefissoNumeroConferimento = PrefissoNumeroConferimento
    rs.Update
End If




rs.Close
Set rs = Nothing

End Sub

Private Sub AggiornamentoProgressivoSezionale(NumeroDocumento)
Dim sSQL As String
Dim rsCtrl As DmtOleDbLib.adoResultset

    If EsistenzaValoreSezionale = True Then
        
        sSQL = "UPDATE ProgressivoSezionale SET "
        sSQL = sSQL & "ProgressivoDisponibile=" & NumeroDocumento + 1 & " "
        sSQL = sSQL & "WHERE ((IDPeriodoIva=" & Link_PeriodoIVA & ") "
        sSQL = sSQL & "AND (IDSezionale=" & Link_Sezionale & ") "
        sSQL = sSQL & "AND (IDTipoModulo=1))"



            
    Else
        sSQL = "INSERT INTO ProgressivoSezionale (IDProgressivoSezionale, IDPeriodoIva, IDTipoModulo, IDSezionale, "
        sSQL = sSQL & "ProgressivoDisponibile, IDUtenteUltimaVariazione, VirtualDelete, DataUltimaVariazione) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("ProgressivoSezionale", "IDProgressivoSezionale") & ", "
        sSQL = sSQL & Link_PeriodoIVA & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & Link_Sezionale & ", "
        sSQL = sSQL & NumeroDocumento + 1 & ", "
        sSQL = sSQL & TheApp.IDUser & ", "
        sSQL = sSQL & 0 & ", "
        sSQL = sSQL & fnNormDate(Date) & ")"
    End If
    
    If Len(sSQL) > 0 Then
        Cn.Execute sSQL
    End If
    
End Sub
Private Function CREA_RIGA_CONFERIMENTO(IDConferimentoTesta As Long, IDRigaOrdine As Long) As Long
On Error GoTo ERR_CREA_RIGA_CONFERIMENTO
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rsOrd As DmtOleDbLib.adoResultset


lblInfoConf.Caption = "Creazione riga conferimento..."
DoEvents

sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & "  AND IDValoriOggettoDettaglio=" & IDRigaOrdine

Set rsOrd = Cn.OpenResultset(sSQL)


sSQL = "SELECT * FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=0"

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


If rsOrd.EOF = False Then
    rsNew.AddNew
        rsNew!IDRV_POCaricoMerceRighe = fnGetNewKey("RV_POCaricoMerceRighe", "IDRV_POCaricoMerceRighe")
        rsNew!IDRV_POcaricoMercetesta = IDConferimentoTesta
        rsNew!IDArticolo = rsOrd!Link_Art_articolo
        rsNew!CodiceArticolo = rsOrd!Art_codice
        rsNew!Articolo = rsOrd!Art_descrizione
        rsNew!IDUnitaDiMisura = fnGetUMCoop(fnNotNullN(rsOrd!Link_Art_unita_di_misura))
        rsNew!IDUnitaDiMisuraDiamante = rsOrd!Link_Art_unita_di_misura
        rsNew!Colli = rsOrd!Art_numero_colli
        rsNew!PesoLordo = rsOrd!Art_peso
        rsNew!PesoNetto = rsOrd!Art_peso - rsOrd!Art_tara
        rsNew!Tara = rsOrd!Art_tara
        rsNew!Pezzi = rsOrd!Art_quantita_pezzi
        rsNew!Qta_UM = rsOrd!Art_quantita_totale
        

        
        rsNew!IDImballo = rsOrd!RV_POIDImballo
        rsNew!CodiceImballo = rsOrd!RV_POCodiceImballo
        rsNew!DescrizioneImballo = rsOrd!RV_PODescrizioneImballo
        rsNew!TaraUnitaria = 0
        
        rsNew!Chiuso = False
        rsNew!AutomazioneLotto = False
        rsNew!IDLottoDiCampagna = 0
        rsNew!Link_Ordinamento = 1
        rsNew!CodiceUtente = ""
        rsNew!IDUtente = TheApp.IDUser
        rsNew!IDProvincia = LINK_PROVINCIA_SOCIO
        rsNew!IDComune = LINK_COMUNE_SOCIO
        rsNew!IDNazione = LINK_NAZIONE_SOCIO
        rsNew!IDRegione = LINK_REGIONE_SOCIO
        rsNew!UtentePC = GET_NOMEUTENTE
        rsNew!NomePC = GET_NOMECOMPUTER
        rsNew!IDArticoloPedana = 0
        rsNew!TaraPedana = 0
        rsNew!CodiceArticoloPedana = ""
        rsNew!ArticoloPedana = ""
        rsNew!QuantitaPedana = 0
        rsNew!OraArrivoMerce = GET_ORARIO(Now)
        rsNew!PrezzoMedio = 0
        rsNew!IDRV_POTipoLavorazione = 0
        rsNew!IDRV_POTipoConfLiquidazione = 0
        rsNew!IDRV_POLottoImballo = 0
        rsNew!TracciaImballo = 0
        
        rsNew!IDCodiceLotto = GET_NUMERO_LOTTO
        rsNew!CodiceLotto = fnGetCodiceLotto(1, 1, "", rsNew!IDCodiceLotto, rsNew, IDConferimentoTesta)
        rsNew!DescrizioneLotto = ""
        
        rsNew!IDValoriOggettoDettaglioOrdine = IDRigaOrdine
        
        
    rsNew.Update
    
    CREA_RIGA_CONFERIMENTO = fnNotNullN(rsNew!IDRV_POCaricoMerceRighe)
    
    IDRigaConf = fnNotNullN(rsNew!IDRV_POCaricoMerceRighe)
    CodiceLottoEntrataConf = fnNotNull(rsNew!CodiceLotto)
    IDArticoloConf = fnNotNullN(rsNew!IDArticolo)
    ArticoloConf = fnNotNull(rsNew!Articolo)
    IDUMConf = fnNotNullN(rsNew!IDUnitaDiMisura)
    
    lblInfoConf.Caption = "Movimentazione riga conferimento..."
    DoEvents
    
    SalvataggioRighe IDTestataConf
    
End If


Exit Function
ERR_CREA_RIGA_CONFERIMENTO:
    MsgBox Err.Description, vbCritical, "CREA_RIGA_CONFERIMENTO"

End Function

Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
    sSQL = sSQL & "IDUnitaDiMisura = " & fnNotNullN(Link_UMAcq)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetUMCoop = fnNotNullN(rs!RV_POIDUnitaDiMisuraCoop)
    Else
        fnGetUMCoop = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function GET_NOMECOMPUTER() As String
Dim dwLen As Long
Dim strString As String
Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    GET_NOMECOMPUTER = strString
    
End Function

Function GET_NOMEUTENTE() As String
    Dim strString As String
    Dim lunghezzaStringa As Long
    lunghezzaStringa = 32
    strString = String(lunghezzaStringa, " ")
    GetUserName strString, lunghezzaStringa
    strString = Left(strString, lunghezzaStringa)
    GET_NOMEUTENTE = strString
    GET_NOMEUTENTE = Mid(GET_NOMEUTENTE, 1, Len(GET_NOMEUTENTE) - 1)
    
End Function
Private Function GET_ORARIO(StringaData As String) As String
Dim Ora As String
Dim Minuti As String
Dim Secondi As String

If Len(DatePart("h", StringaData)) = 1 Then
    Ora = "0" & DatePart("h", StringaData)
Else
    Ora = DatePart("h", StringaData)
End If
If Len(DatePart("n", StringaData)) = 1 Then
    Minuti = "0" & DatePart("n", StringaData)
Else
    Minuti = DatePart("n", StringaData)
End If
If Len(DatePart("s", StringaData)) = 1 Then
    Secondi = "0" & DatePart("s", StringaData)
Else
    Secondi = DatePart("s", StringaData)
End If

GET_ORARIO = Ora & "." & Minuti & "." & Secondi

End Function
Private Function GET_NUMERO_LOTTO() As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT NumerazioneLottoConferimento FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    GET_NUMERO_LOTTO = 0
Else
    If fnNotNullN(rs!NumerazioneLottoConferimento) + 1 < 10000000 Then
        GET_NUMERO_LOTTO = fnNotNullN(rs!NumerazioneLottoConferimento) + 1
        rs!NumerazioneLottoConferimento = GET_NUMERO_LOTTO
    Else
        GET_NUMERO_LOTTO = 1
        rs!NumerazioneLottoConferimento = 1
    
    End If
    rs.Update
End If
End Function
Private Function GET_NUMERO_LOTTO_VENDITA() As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT NumerazioneLotto FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    GET_NUMERO_LOTTO_VENDITA = 0
Else
    If (fnNotNullN(rs!NumerazioneLotto) + 1) < 10000000 Then
        GET_NUMERO_LOTTO_VENDITA = fnNotNullN(rs!NumerazioneLotto) + 1
        rs!NumerazioneLotto = fnNotNullN(rs!NumerazioneLotto) + 1
    Else
        GET_NUMERO_LOTTO_VENDITA = 1
        rs!NumerazioneLotto = 1
    End If
    rs.Update
End If

rs.Close
Set rs = Nothing
End Function
Private Function fnGetCodiceLotto(TipoLotto As Integer, TipoStringaLotto As Integer, StringaLotto As String, Link_LottoArticolo As Long, rsRiga As ADODB.Recordset, IDTestaConferimento As Long) As String
On Error GoTo ERR_fnGetCodiceLotto
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim Codice As String
Dim I As Integer
Dim PosizioneCodice As Integer
Dim Stringa As String
Dim StringaElaborata As String
Dim SenzaIndentificativo As Boolean
Dim rsTestaConferimento As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POLottoCostruzioneRighe.IDRV_POLottoComp, RV_POLottoCostruzioneRighe.Posizione, RV_POLottoCostruzioneRighe.Lunghezza, "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.Testo, RV_POLottoCostruzioneTesta.PosVendita, RV_POLottoCostruzioneTesta.PosConferimento, "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.SXDX, RV_POLottoCostruzioneTesta.SenzaCodiceRiferimento_Conf "
sSQL = sSQL & "FROM RV_POLottoCostruzioneRighe INNER JOIN "
sSQL = sSQL & "RV_POLottoCostruzioneTesta ON "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.IDRV_POLottoCostruzioneTesta = RV_POLottoCostruzioneTesta.IDRV_POLottoCostruzioneTesta "
sSQL = sSQL & "WHERE RV_POLottoCostruzioneTesta.IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoLotto=" & TipoLotto
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoStringaLotto=" & TipoStringaLotto
sSQL = sSQL & " ORDER BY Posizione"


fnGetCodiceLotto = ""
StringaElaborata = ""
Codice = ""
PosizioneCodice = 0
SenzaIndentificativo = False
Set rs = Cn.OpenResultset(sSQL)

If Link_LottoArticolo > 0 Then
    For I = Len(CStr(Link_LottoArticolo)) To 7
        Codice = Codice & "0"
    Next
    Codice = Codice & Link_LottoArticolo
End If
            

If rs.EOF Then
    StringaElaborata = ""
    
Else
    
    PosizioneCodice = fnNotNullN(rs!PosConferimento)
    SenzaIndentificativo = fnNotNullN(rs!SenzaCodiceRiferimento_Conf)
    
    fnGetCodiceLotto = StringaLotto
    
    While Not rs.EOF
        Select Case fnNotNullN(rs!IDRV_POLottoComp)
            Case 1 'Codice socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CodiceSocioConf, 0, fnNotNullN(rs!SXDX))
            Case 2 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), AnagraficaSocioConf, 1, fnNotNullN(rs!SXDX))
            Case 3 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), NomeSocioConf, 1, fnNotNullN(rs!SXDX))
            Case 4 'Giorno conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("d", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 5 'Mese del conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("m", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 6 'Anno del conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("yyyy", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 7 'Giorno lavorazione
                
            Case 8 'Mese lavorazione
                
            Case 9 'Anno lavorazione
                
            Case 10 'calibro
                
            Case 11 'Tipo lavorazione
            
            Case 12 'Tipo categoria
            
            Case 13 'Carattere speciale "_"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("_"), 1, fnNotNullN(rs!SXDX))
            Case 14 'Carattere speciale "-"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("-"), 1, fnNotNullN(rs!SXDX))
            Case 15 'Stringa personalizzata
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(rs!Testo)), 1, fnNotNullN(rs!SXDX))
            Case 16 'Codice imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!CodiceImballo), 1, fnNotNullN(rs!SXDX))
            Case 17 'Descrizione imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!DescrizioneImballo), 1, fnNotNullN(rs!SXDX))
            Case 18 'Codice pedana
                
            Case 19
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!CodiceArticolo), 1, fnNotNullN(rs!SXDX))
            Case 20
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!Articolo), 1, fnNotNullN(rs!SXDX))
            Case 22
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("ww", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 23
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 24
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(""), 1, fnNotNullN(rs!SXDX))
            Case 25
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 26
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(NumeroDocumentoConf), 0, fnNotNullN(rs!SXDX))
            Case 27 'Codice certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 28 'Descrizione certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 29 'Protocollo certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 30 'Codice certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 31 'Descrizione certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 32 'Protocollo certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 33 'Giorno della settimana
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("w", fnNotNull(rsTestaConferimento!DataDocumento)) - 1), 0, fnNotNullN(rs!SXDX))
            Case 34 'Codice utente
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 35 'Descrizione lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 36 'Codice certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 37 'Descrizione certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 38 'Protocollo certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))

        End Select
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing
    

    
    If StringaElaborata = "" Then
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Mid(Codice, 1, Len(Codice) - 1)
            Else
                fnGetCodiceLotto = Mid(Codice, 2, Len(Codice))
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    Else
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Codice & StringaElaborata
            Else
                fnGetCodiceLotto = StringaElaborata & Codice
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    End If
Exit Function

ERR_fnGetCodiceLotto:
    
    If StringaElaborata = "" Then
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Mid(Codice, 1, Len(Codice) - 1)
            Else
                fnGetCodiceLotto = Mid(Codice, 2, Len(Codice))
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    Else
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Codice & StringaElaborata
            Else
                fnGetCodiceLotto = StringaElaborata & Codice
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    End If

End Function

Private Function SalvataggioRighe(IDConferimentoTesta As Long) As Boolean
On Error GoTo ERR_SalvataggioRighe
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsCount As DmtOleDbLib.adoResultset
Dim Unita_progresso As Double

Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection



    Mov.IDTipoOggetto = IDTipoOggettoConf
    Mov.IDOggetto = IDConferimentoTesta
    Mov.Delete

    
    sSQL = "SELECT * FROM RV_POCaricoMerceRighe "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDConferimentoTesta
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection
    
    While Not rs.EOF
        If fnNotNullN(rs!IDArticolo) > 0 Then
            
            If GeneraMovimentoDiCarico(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNull(rs!CodiceLotto), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDUnitaDiMisuraDiamante), fnNotNull(rs!Articolo), fnNotNullN(rs!Qta_UM), _
                fnNotNullN(rs!ImportoUnitario), fnNotNullN(rs!TotaleImponibileRiga), _
                fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), _
                fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi)) = False Then
                
                SalvataggioRighe = False
                rs.Close
                Set rs = Nothing
                Set Mov = Nothing
                Screen.MousePointer = 0
                Exit Function
            End If
                
            If fnNotNullN(rs!IDImballo > 0) Then
                If GeneraMovimentoCaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNull(rs!CodiceLotto), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDImballo), fnNotNullN(8), fnNotNull(rs!DescrizioneImballo), fnNotNullN(rs!Colli), fnNotNullN(rs!IDRV_POLottoImballo)) = False Then
                    SalvataggioRighe = False
                    rs.Close
                    Set rs = Nothing
                    Set Mov = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
                
            End If
            
            If ((fnNotNullN(rs!QuantitaPedana) > 0) And (fnNotNullN(rs!IDArticoloPedana) > 0)) Then
                If GeneraMovimentoCaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNull(rs!CodiceLotto), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDArticoloPedana), fnNotNullN(8), fnNotNull(rs!ArticoloPedana), fnNotNullN(rs!QuantitaPedana), 0) = False Then
                    SalvataggioRighe = False
                    rs.Close
                    Set rs = Nothing
                    Set Mov = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
            End If
            

        End If
       

    DoEvents
    rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing
    
Set Mov = Nothing


SalvataggioRighe = True

Exit Function
ERR_SalvataggioRighe:
    MsgBox Err.Description, vbCritical, "Salvataggio Righe"
    SalvataggioRighe = False
End Function


Private Function GeneraMovimentoDiCarico(IDRigaConferimento As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, _
PrezzoUnitario As Double, PrezzoImponibile As Double, Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double) As Boolean



Mov.DataMovimento = DataDocumentoConf
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = Link_Esercizio
Mov.IDTipoOggetto = IDTipoOggettoConf
Mov.IDOggetto = IDTestataConf
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(IDDocumento, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Link_Magazzino_Carico
Mov.IDMagazzinoUscita = Link_Magazzino_Carico
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", IDAnagraficaSocioConf
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUMDiamante
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", PrezzoImponibile
Mov.Field "PrezzoUnitario", PrezzoUnitario
Mov.Field "DataDocumento", DataDocumentoConf
Mov.Field "Oggetto", Mid(DescrizioneFunzioneConf & " del " & DataDocumentoConf & " Numero " & NumeroDocumentoConf, 1, 100)
Mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDRigaConferimento
Mov.Field "RV_POTipoRiga", 1
Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
Mov.Field "RV_POIDAssegnazioneMerce", 0
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocioConf
Mov.Field "RV_PODataConferimento", DataDocumentoConf
Mov.Field "RV_PONumeroConferimento", NumeroDocumentoConf
Mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
Mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POQuantitaMovimentata", Qta_UM
Mov.Field "RV_PONumeroColli", Colli
Mov.Field "RV_POPesoLordo", PesoLordo
Mov.Field "RV_POPesoNetto", PesoNetto
Mov.Field "RV_POTara", Tara
Mov.Field "RV_POQuantitaPezzi", Pezzi
Mov.Field "RV_POIDAnagraficaFatturazione", IDAnagraficaSocioConf
Mov.Field "RV_POIDLottoImballo", 0
Mov.Field "LottoImballo", ""


Mov.Field "TipoRiga", trcNessuno

GeneraMovimentoDiCarico = Mov.Insert



End Function
Public Function GeneraMovimentoCaricoImballo(IDRigaConferimento As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, IDLottoImballo As Long) As Boolean


Mov.DataMovimento = DataDocumentoConf
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = Link_Esercizio
Mov.IDTipoOggetto = IDTipoOggettoConf
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(IDDocumento, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Link_Magazzino_Carico
Mov.IDMagazzinoUscita = Link_Magazzino_Carico
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", IDAnagraficaSocioConf
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUMDiamante
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", 0
Mov.Field "PrezzoUnitario", 0
Mov.Field "DataDocumento", DataDocumentoConf
Mov.Field "Oggetto", Mid(DescrizioneFunzioneConf & " del " & DataDocumentoConf & " Numero " & NumeroDocumentoConf, 1, 100)
Mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDRigaConferimento
Mov.Field "RV_POTipoRiga", 2
Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
Mov.Field "RV_POIDAssegnazioneMerce", 0
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocioConf
Mov.Field "RV_PODataConferimento", DataDocumentoConf
Mov.Field "RV_PONumeroConferimento", NumeroDocumentoConf
Mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
Mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POIDAnagraficaFatturazione", IDAnagraficaSocioConf

Mov.Field "RV_POIDLottoImballo", 0
Mov.Field "LottoImballo", ""



Mov.Field "TipoRiga", trcNessuno

GeneraMovimentoCaricoImballo = Mov.Insert




End Function

Private Function GET_STRINGALOTTO(IDRV_POLottoComp As Long, Lunghezza As Integer, Stringa As String, TipoStringa As Integer, SXDX As Integer, Optional TestoPersonalizzato As String) As String
Dim Parziale As String
Dim I As Integer
Parziale = ""

        If Len(Stringa) >= Lunghezza Then
            If SXDX = 2 Then 'Da destra verso sinistra
                GET_STRINGALOTTO = Right(Stringa, Lunghezza)
            Else
                'Da sinistra verso destra
                GET_STRINGALOTTO = Mid(Stringa, 1, Lunghezza)
            End If
        Else
            If SXDX <= 1 Then 'Da sinistra verso destra
                For I = Len(Stringa) To Lunghezza - 1
                    If TipoStringa = 0 Then
                        Parziale = "0" & Parziale
                    Else
                        Parziale = "" & Parziale
                    End If
                Next
                GET_STRINGALOTTO = Parziale & Stringa
            Else
                'Da destra verso sinistra
                GET_STRINGALOTTO = Right(Stringa, Lunghezza)
            End If
        End If
End Function
Private Function GET_FUNZIONE_MAGAZZINO(IDTipoDocumentoCoop As Long, IDTipoProcesso As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POProcessiDocumentoCoop.IDFunzione "
sSQL = sSQL & "FROM RV_POProcessiDocumentoCoop INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE RV_POSchemaCoop.IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDDocumentoCoop=" & IDTipoDocumentoCoop
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDTipoProcessoCoop=" & IDTipoProcesso

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Select Case IDTipoProcesso
        Case 1 'Carico
            GET_FUNZIONE_MAGAZZINO = Link_CausaleCarico
        Case 2 'Scarico
            GET_FUNZIONE_MAGAZZINO = Link_CausaleScarico
    End Select
Else
    If fnNotNullN(rs!IDFunzione) = 0 Then
        Select Case IDTipoProcesso
            Case 1 'Carico
                GET_FUNZIONE_MAGAZZINO = Link_CausaleCarico
            Case 2 'Scarico
                GET_FUNZIONE_MAGAZZINO = Link_CausaleScarico
        End Select
    Else
        GET_FUNZIONE_MAGAZZINO = fnNotNullN(rs!IDFunzione)
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_RIGA_ORDINE(IDRigaConferimento As Long, IDRigaOrdine As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoDettaglio0010 SET "
sSQL = sSQL & "IDRV_POCaricoMerceRighe=" & IDRigaConferimento
sSQL = sSQL & " WHERE IDValoriOggettoDettaglio=" & IDRigaOrdine

Cn.Execute sSQL

End Sub
Private Sub CREA_RIGA_DI_LAVORAZIONE(IDRigaOrdine As Long, IDRigaConf As Long, IDTestaConf As Long)
On Error GoTo ERR_CREA_RIGA_DI_LAVORAZIONE
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rsOrd As DmtOleDbLib.adoResultset


lblInfoConf.Caption = "Creazione riga di lavorazione..."
DoEvents

sSQL = "SELECT * FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & "  AND IDValoriOggettoDettaglio=" & IDRigaOrdine

Set rsOrd = Cn.OpenResultset(sSQL)


sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=0"

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


If rsOrd.EOF = False Then
    rsNew.AddNew
        rsNew!IDRV_POAssegnazioneMerce = fnGetNewKey("RV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce")
        rsNew!IDRV_POCaricoMerceRighe = IDRigaConf
        rsNew!IDTipoLavorazione = rsOrd!RV_POIDTipoLavorazione
        rsNew!IDRV_POTipoCategoria = rsOrd!RV_POIDTipoCategoria
        rsNew!IDRV_POCalibro = rsOrd!RV_POIDCalibro
        
        rsNew!DataDocumento = Date
        rsNew!IDArticolo = rsOrd!Link_Art_articolo
        rsNew!CodiceArticolo = rsOrd!Art_codice
        rsNew!Articolo = rsOrd!Art_descrizione
        rsNew!IDUnitaDiMisura = rsOrd!Link_Art_unita_di_misura
        rsNew!IDUnitaDiMisuraCoop = fnGetUMCoop(fnNotNullN(rsOrd!Link_Art_unita_di_misura))
        rsNew!Colli = rsOrd!Art_numero_colli
        rsNew!PesoLordo = rsOrd!Art_peso
        rsNew!PesoNetto = rsOrd!Art_peso - rsOrd!Art_tara
        rsNew!Tara = rsOrd!Art_tara
        rsNew!Pezzi = rsOrd!Art_quantita_pezzi
        rsNew!Qta_UM = rsOrd!Art_quantita_totale
        
        rsNew!ImportoUnitarioArticolo = rsOrd!Art_prezzo_unitario_netto_IVA
        rsNew!Sconto1 = fnNotNullN(rsOrd!Art_sco_in_percentuale_1)
        rsNew!Sconto2 = fnNotNullN(rsOrd!Art_sco_in_percentuale_2)
        rsNew!RV_POImportoUnitarioListino = rsOrd!RV_POImportoUnitarioListino
        rsNew!NotaRigaOrdRaggr = fnNotNull(rsOrd!RV_PONotaRigaOrdRaggr)
        rsNew!IDValoriOggettoDettaglioRigaOrd = rsOrd!IDValoriOggettoDettaglio
        

        GetNumeroPedana Year(rsNew!DataDocumento), rsNew!DataDocumento, rsNew
        
        rsNew!IDImballoVendita = rsOrd!RV_POIDImballo
        rsNew!CodiceImballoVendita = rsOrd!RV_POCodiceImballo
        rsNew!ImballoVendita = rsOrd!RV_PODescrizioneImballo
        
        rsNew!TaraUnitaria = 0
        rsNew!ImportoUnitarioImballo = 0
        rsNew!MerceInclusoImballo = 0

        
        rsNew!NumeroOrdine = frmMain.lngNumero
        rsNew!DataOrdine = frmMain.dtData
        rsNew!idcliente = frmMain.cdAnagrafica.KeyFieldID
        rsNew!IDOggettoOrdine = oDoc.IDOggetto
        
        rsNew!IDAnagraficaSocio = IDAnagraficaSocioConf
        rsNew!CodiceSocio = CodiceSocioConf
        rsNew!AnagraficaSocio = AnagraficaSocioConf
        rsNew!NomeSocio = NomeSocioConf
        rsNew!DataConferimento = DataConf
        rsNew!NumeroConferimento = NumeroConf
        
        rsNew!LottoCliente = 0
        rsNew!AltreAnnotazioniPerCliente = ""
        rsNew!Link_Ordinamento = 1
        rsNew!IDRV_POProcessoIVGamma = 0
        rsNew!IDUtente = TheApp.IDUser
        rsNew!NumeroPedanaCliente = 0
        rsNew!CodiceUtente = ""
        rsNew!UtentePC = GET_NOMEUTENTE
        rsNew!NomePC = GET_NOMECOMPUTER
        rsNew!IDLinguaPredefinita = GET_LINK_LINGUA_PREDEFINITA
        
        rsNew!IDLinguaCliente = GET_LINK_LINGUA_CLIENTE(fnNotNullN(rsNew!IDLinguaPredefinita), frmMain.cdAnagrafica.KeyFieldID)
        
        GET_INFO_CLIENTE frmMain.cdAnagrafica.KeyFieldID, rsNew
        
        GET_ARTICOLO_CODICE_A_BARRE_CLIENTE frmMain.cdAnagrafica.KeyFieldID, rsNew!IDArticolo, rsNew
        
        GET_IMBALLO_CODICE_A_BARRE_CLIENTE frmMain.cdAnagrafica.KeyFieldID, fnNotNullN(rsNew!IDImballoVendita), rsNew
        
        
        rsNew!CodiceABarreArticoloPred = ""
        rsNew!DescrizioneCodiceABarreArticoloPred = ""
        rsNew!CodiceABarreImballoPred = ""
        rsNew!DescrizioneCodiceABarreImballoPred = ""
        
        rsNew!DescrizioneArticoloInLinguaPred = GET_DESCRIZIONE_IN_LINGUA(fnNotNullN(rsNew!IDLinguaPredefinita), fnNotNullN(rsNew!IDArticolo))
        rsNew!DescrizioneCalibroInLinguaPred = GET_DESCRIZIONE_CALIBRO_IN_LINGUA(fnNotNullN(rsNew!IDLinguaPredefinita), fnNotNullN(rsNew!IDRV_POCalibro))
        rsNew!DescrizioneCategoriaInLinguaPred = GET_DESCRIZIONE_CATEGORIA_IN_LINGUA(fnNotNullN(rsNew!IDLinguaPredefinita), fnNotNullN(rsNew!IDRV_POTipoCategoria))
        
        rsNew!DescrizioneArticoloInLinguaCliente = GET_DESCRIZIONE_IN_LINGUA(fnNotNullN(rsNew!IDLinguaCliente), fnNotNullN(rsNew!IDArticolo))
        rsNew!DescrizioneCalibroInLinguaCliente = GET_DESCRIZIONE_CALIBRO_IN_LINGUA(fnNotNullN(rsNew!IDLinguaCliente), fnNotNullN(rsNew!IDRV_POCalibro))
        rsNew!DescrizioneCategoriaInLinguaCliente = GET_DESCRIZIONE_CATEGORIA_IN_LINGUA(fnNotNullN(rsNew!IDLinguaCliente), fnNotNullN(rsNew!IDRV_POTipoCategoria))
        
        rsNew!LinguaPredefinita = ""
        rsNew!LinguaCliente = ""
        
        rsNew!OraLavorazione = GET_ORARIO(Now)
        rsNew!IDRV_POAssegnazioneMercePadre = rsNew!IDRV_POAssegnazioneMerce
        rsNew!IDRV_POProcessoIVGammaRighe = 0
        rsNew!RV_PO01_IDSezionaleRighe = 0
        rsNew!RV_PO01_NumeroPassaportoRighe = 0
        rsNew!TracciaImballo = 0
        rsNew!ConfermaDaUtente = 0
        rsNew!IDRV_POLottoImballo = 0
        

        rsNew!NumeroConfezioniPerImballo = 0
        rsNew!TaraConfezioneImballo = 0
        
        
        rsNew!CodiceLottoVendita = fnGetCodiceLottoVendita(2, 1, "", GET_NUMERO_LOTTO_VENDITA, rsNew, IDTestaConf)
        
    rsNew.Update
    
    
    lblInfoConf.Caption = "Movimentazione lavorazione..."
    DoEvents
    MOVIMENTAZIONE_RIGA_LAVORAZIONE rsNew!IDRV_POAssegnazioneMerce
    
End If

Exit Sub
ERR_CREA_RIGA_DI_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, "CREA_RIGA_DI_LAVORAZIONE"
End Sub
Private Sub GetNumeroPedana(Anno As Long, DataLavorazione As String, rstmp As ADODB.Recordset)
On Error GoTo ERR_GetNumeroPedana
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim CodicePedana As String
Dim NumeroPedana As Long
Dim X As Long
Dim ErroreCoda As Boolean
Dim LINK_PEDANA As Long


sSQL = "SELECT MAX(CodiceID) AS NumeroPedana FROM RV_POPedana "
sSQL = sSQL & "WHERE Anno=" & Anno
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch

Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        CodicePedana = Anno & GET_CODICE_PEDANA(1)
        NumeroPedana = 1
    Else
        NumeroPedana = (fnNotNullN(rs!NumeroPedana) + 1)
        CodicePedana = Anno & GET_CODICE_PEDANA((fnNotNullN(rs!NumeroPedana) + 1))
    End If
    
rs.CloseResultset
Set rs = Nothing

LINK_PEDANA = fnGetNewKey("RV_POPedana", "IDRV_POPedana")

sSQL = "INSERT INTO RV_POPedana ("
sSQL = sSQL & "IDRV_POPedana, CodiceID, IDAzienda, IDFiliale, Anno, Mese, Giorno, IDRV_POTipoPedana, Codice) "
sSQL = sSQL & "VALUES ("
sSQL = sSQL & LINK_PEDANA & ", "
sSQL = sSQL & NumeroPedana & ","
sSQL = sSQL & TheApp.IDFirm & ", "
sSQL = sSQL & TheApp.Branch & ", "
sSQL = sSQL & DatePart("yyyy", DataLavorazione) & ", "
sSQL = sSQL & Month(DataLavorazione) & ", "
sSQL = sSQL & Day(DataLavorazione) & ", "
sSQL = sSQL & GET_TIPO_PEDANA_DEFAULT & ", "
sSQL = sSQL & fnNormString(CodicePedana) & " "
sSQL = sSQL & ")"

Cn.Execute sSQL

        
rstmp!IDRV_POPedana = LINK_PEDANA
rstmp!CodicePedana = CodicePedana
rstmp!DescrizionePedana = ""
        
Exit Sub
ERR_GetNumeroPedana:
    MsgBox Err.Description, vbCritical, "Nuova pedana"
    
    
End Sub
Private Function GET_CODICE_PEDANA(NumeroPedana As String) As String
Dim I As Integer
Const MAX_CAR As Integer = 7
GET_CODICE_PEDANA = ""
For I = Len(NumeroPedana) + 1 To MAX_CAR
GET_CODICE_PEDANA = GET_CODICE_PEDANA & "0"
    
Next
GET_CODICE_PEDANA = GET_CODICE_PEDANA & NumeroPedana
End Function
Private Function GET_TIPO_PEDANA_DEFAULT() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoPedanaDefault "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PEDANA_DEFAULT = 0
Else
    GET_TIPO_PEDANA_DEFAULT = fnNotNullN(rs!IDTipoPedanaDefault)
    
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_LINGUA_PREDEFINITA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDLinguaPredefinita FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_LINK_LINGUA_PREDEFINITA = fnNotNullN(rs!IDLinguaPredefinita)
Else
    GET_LINK_LINGUA_PREDEFINITA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_IN_LINGUA(IDLingua As Long, IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ArticoloPerLinguaDescrizione "
sSQL = sSQL & "FROM ArticoloPerLinguaDescrizione "
sSQL = sSQL & "WHERE IDLinguaDescrizioneArticolo=" & IDLingua
sSQL = sSQL & " AND IDArticolo=" & IDArticolo
sSQL = sSQL & " AND VirtualDelete=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_IN_LINGUA = ""
Else
    GET_DESCRIZIONE_IN_LINGUA = fnNotNull(rs!ArticoloPerLinguaDescrizione)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_CALIBRO_IN_LINGUA(IDLingua As Long, IDcalibro As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CalibroLingua "
sSQL = sSQL & "FROM RV_POCalibroLingua "
sSQL = sSQL & "WHERE IDLingua=" & IDLingua
sSQL = sSQL & " AND IDRV_POCalibro=" & IDcalibro

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_CALIBRO_IN_LINGUA = ""
Else
    GET_DESCRIZIONE_CALIBRO_IN_LINGUA = fnNotNull(rs!CalibroLingua)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_CATEGORIA_IN_LINGUA(IDLingua As Long, IDCategoria As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CategoriaLingua "
sSQL = sSQL & "FROM RV_POCategoriaLingua "
sSQL = sSQL & "WHERE IDLingua=" & IDLingua
sSQL = sSQL & " AND IDRV_POTipoCategoria=" & IDCategoria

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_CATEGORIA_IN_LINGUA = ""
Else
    GET_DESCRIZIONE_CATEGORIA_IN_LINGUA = fnNotNull(rs!CategoriaLingua)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_ARTICOLO_CODICE_A_BARRE_CLIENTE(IDAnagrafica As Long, IDArticolo As Long, rstmp As ADODB.Recordset)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceABarre, DescrizioneCodiceABarre "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteEAN13 "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rstmp!CodiceABarreArticoloCliente = ""
    rstmp!DescrizioneCodiceABarreArticoloCliente = ""

Else
    rstmp!CodiceABarreArticoloCliente = fnNotNull(rs!CodiceABarre)
    rstmp!DescrizioneCodiceABarreArticoloCliente = fnNotNull(rs!DescrizioneCodiceABarre)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_IMBALLO_CODICE_A_BARRE_CLIENTE(IDAnagrafica As Long, IDArticolo As Long, rstmp As ADODB.Recordset)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceABarre, DescrizioneCodiceABarre "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteEAN13 "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rstmp!CodiceABarreImballoCliente = ""
    rstmp!DescrizioneCodiceABarreImballoCliente = ""
Else
    rstmp!CodiceABarreImballoCliente = fnNotNullN(rs!CodiceABarre)
    rstmp!DescrizioneCodiceABarreImballoCliente = fnNotNull(rs!DescrizioneCodiceABarre)

End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_LINK_LINGUA_CLIENTE(IDLinguaPred As Long, idcliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

''''''''''''''''''LINGUA DESCRIZIONE ARTICOLI''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDLinguaDescrizioneArticoli "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & idcliente

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LINGUA_CLIENTE = IDLinguaPred
Else
    If fnNotNullN(rs!IDLinguaDescrizioneArticoli) > 0 Then
        GET_LINK_LINGUA_CLIENTE = fnNotNullN(rs!IDLinguaDescrizioneArticoli)
    Else
       GET_LINK_LINGUA_CLIENTE = IDLinguaPred
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_INFO_CLIENTE(IDAnagrafica As Long, rstmp As ADODB.Recordset)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoInclusoImballo,CodiceGSI, MaxCaratteriPedana, CodiceAssociato "
sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = Cn.OpenResultset(sSQL)



If rs.EOF Then
    
    rstmp!CodiceGSI = ""
    rstmp!CodiceAssociatoPressoCliente = ""
Else
    rstmp!CodiceGSI = fnNotNull(rs!CodiceGSI)
    rstmp!CodiceAssociatoPressoCliente = fnNotNull(rs!CodiceAssociato)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Function fnGetCodiceLottoVendita(TipoLotto As Integer, TipoStringaLotto As Integer, StringaLotto As String, Link_LottoArticolo As Long, rsRiga As ADODB.Recordset, IDTestaConferimento As Long) As String
On Error GoTo ERR_fnGetCodiceLottoVendita
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim Codice As String
Dim I As Integer
Dim PosizioneCodice As Integer
Dim Stringa As String
Dim StringaElaborata As String
Dim SenzaIndentificativo As Boolean
Dim rsTestaConferimento As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POLottoCostruzioneRighe.IDRV_POLottoComp, RV_POLottoCostruzioneRighe.Posizione, RV_POLottoCostruzioneRighe.Lunghezza, "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.Testo, RV_POLottoCostruzioneTesta.PosVendita, RV_POLottoCostruzioneTesta.PosConferimento, "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.SXDX, RV_POLottoCostruzioneTesta.SenzaCodiceRiferimento_Conf "
sSQL = sSQL & "FROM RV_POLottoCostruzioneRighe INNER JOIN "
sSQL = sSQL & "RV_POLottoCostruzioneTesta ON "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.IDRV_POLottoCostruzioneTesta = RV_POLottoCostruzioneTesta.IDRV_POLottoCostruzioneTesta "
sSQL = sSQL & "WHERE RV_POLottoCostruzioneTesta.IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoLotto=" & TipoLotto
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoStringaLotto=" & TipoStringaLotto
sSQL = sSQL & " ORDER BY Posizione"


fnGetCodiceLottoVendita = ""
StringaElaborata = ""
Codice = ""
PosizioneCodice = 0
SenzaIndentificativo = False
Set rs = Cn.OpenResultset(sSQL)

If Link_LottoArticolo > 0 Then
    For I = Len(CStr(Link_LottoArticolo)) To 7
        Codice = Codice & "0"
    Next
    Codice = Codice & Link_LottoArticolo
End If
            

If rs.EOF Then
    StringaElaborata = ""
    
Else
    
    PosizioneCodice = fnNotNullN(rs!PosConferimento)
    SenzaIndentificativo = fnNotNullN(rs!SenzaCodiceRiferimento_Conf)
    
    fnGetCodiceLottoVendita = StringaLotto
    
    While Not rs.EOF
        Select Case fnNotNullN(rs!IDRV_POLottoComp)
            Case 1 'Codice socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CodiceSocioConf, 0, fnNotNullN(rs!SXDX))
            Case 2 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), AnagraficaSocioConf, 1, fnNotNullN(rs!SXDX))
            Case 3 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), NomeSocioConf, 1, fnNotNullN(rs!SXDX))
            Case 4 'Giorno conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("d", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 5 'Mese del conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("m", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 6 'Anno del conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("yyyy", DataDocumentoConf)), 0, fnNotNullN(rs!SXDX))
            Case 7 'Giorno lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("d", Date)), 0, fnNotNullN(rs!SXDX))
            Case 8 'Mese lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("m", Date)), 0, fnNotNullN(rs!SXDX))
            Case 9 'Anno lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("yyyy", Date)), 0, fnNotNullN(rs!SXDX))
                
            Case 10 'calibro
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 11 'Tipo lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 12 'Tipo categoria
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            
            Case 13 'Carattere speciale "_"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("_"), 1, fnNotNullN(rs!SXDX))
            Case 14 'Carattere speciale "-"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("-"), 1, fnNotNullN(rs!SXDX))
            Case 15 'Stringa personalizzata
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(rs!Testo)), 1, fnNotNullN(rs!SXDX))
            Case 16 'Codice imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!CodiceImballo), 1, fnNotNullN(rs!SXDX))
            Case 17 'Descrizione imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!DescrizioneImballo), 1, fnNotNullN(rs!SXDX))
            Case 18 'Codice pedana
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 19
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!CodiceArticolo), 1, fnNotNullN(rs!SXDX))
            Case 20
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), fnNotNull(rsRiga!Articolo), 1, fnNotNullN(rs!SXDX))
            Case 22
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("ww", Date)), 0, fnNotNullN(rs!SXDX))
            Case 23
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", Date)), 0, fnNotNullN(rs!SXDX))
            Case 24
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(""), 1, fnNotNullN(rs!SXDX))
            Case 25
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", Date)), 0, fnNotNullN(rs!SXDX))
            Case 26
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(NumeroDocumentoConf), 0, fnNotNullN(rs!SXDX))
            Case 27 'Codice certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 28 'Descrizione certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 29 'Protocollo certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 30 'Codice certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 31 'Descrizione certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 32 'Protocollo certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 33 'Giorno della settimana
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("w", Date) - 1), 0, fnNotNullN(rs!SXDX))
            Case 34 'Codice utente
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 35 'Descrizione lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 36 'Codice certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 37 'Descrizione certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))
            Case 38 'Protocollo certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), "", 1, fnNotNullN(rs!SXDX))

        End Select
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing
    
    If StringaElaborata = "" Then
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLottoVendita = Mid(Codice, 1, Len(Codice) - 1)
            Else
                fnGetCodiceLottoVendita = Mid(Codice, 2, Len(Codice))
            End If
        Else
            fnGetCodiceLottoVendita = StringaElaborata
        End If
    Else
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLottoVendita = Codice & StringaElaborata
            Else
                fnGetCodiceLottoVendita = StringaElaborata & Codice
            End If
        Else
            fnGetCodiceLottoVendita = StringaElaborata
        End If
    End If
Exit Function

ERR_fnGetCodiceLottoVendita:
    
    If StringaElaborata = "" Then
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLottoVendita = Mid(Codice, 1, Len(Codice) - 1)
            Else
                fnGetCodiceLottoVendita = Mid(Codice, 2, Len(Codice))
            End If
        Else
            fnGetCodiceLottoVendita = StringaElaborata
        End If
    Else
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLottoVendita = Codice & StringaElaborata
            Else
                fnGetCodiceLottoVendita = StringaElaborata & Codice
            End If
        Else
            fnGetCodiceLottoVendita = StringaElaborata
        End If
    End If

End Function

Private Function GeneraMovimentoDiCaricoLav(IDAssegnazione As Long, IDRigaConferimento As Long, CodiceLottoVendita As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, DataLavorazione As String, QuantitaMovimentata As Double, _
Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, IDImballo As Long, IDTipoDocumentoCoop As Long, CodiceImballo As String, DescrizioneImballo As String, _
IDTipoLavorazione As Long, IDTipoCategoria As Long, IDcalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDPedana As Long, IDTipoPedana As Long, CodicePedana As String, PesoPedana As Double) As Long

On Error GoTo ERR_GeneraMovimentoDiCaricoLav

Mov.DataMovimento = DataLavorazione
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = Link_Esercizio
Mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
Mov.IDOggetto = IDRigaConferimento
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Link_Magazzino_Vendita
Mov.IDMagazzinoUscita = Link_Magazzino_Vendita
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", IDAnagraficaSocioConf
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUMDiamante
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", 0
Mov.Field "DataDocumento", DataLavorazione
Mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
Mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
Mov.Field "RV_POTipoRiga", 1
Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
Mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
Mov.Field "RV_PODataConferimento", DataConf
Mov.Field "RV_PONumeroConferimento", NumeroConf
Mov.Field "RV_POCodiceLotto", CodiceLottoEntrataConf
Mov.Field "RV_POCodiceLottoCampagna", ""
Mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POQuantitaMovimentata", QuantitaMovimentata
Mov.Field "RV_PONumeroColli", Colli
Mov.Field "RV_POPesoLordo", PesoLordo
Mov.Field "RV_POPesoNetto", PesoNetto
Mov.Field "RV_POTara", Tara
Mov.Field "RV_POQuantitaPezzi", Pezzi
Mov.Field "RV_POIDTipoDocumentoCoop", 1
Mov.Field "RV_POIDImballo", IDImballo
Mov.Field "RV_POCodiceImballo", CodiceImballo
Mov.Field "RV_PODescrizioneImballo", DescrizioneImballo

Mov.Field "RV_PODataLavorazione", DataLavorazione
Mov.Field "RV_POIDTipoLavorazione", IDTipoLavorazione
Mov.Field "RV_POIDCalibro", IDcalibro
Mov.Field "RV_POIDTipoCategoria", IDTipoCategoria
Mov.Field "RV_POIDTipoLavorazioneConf", IDTipoLavorazioneConf
Mov.Field "RV_POPrezzoMedioConf", PrezzoMedioConf

Mov.Field "RV_POIDPedana", IDPedana
Mov.Field "RV_POIDTipoPedana", IDTipoPedana
Mov.Field "RV_POCodicePedana", CodicePedana
Mov.Field "RV_POPesoPedana", PesoPedana


Mov.Field "TipoRiga", trcNessuno



GeneraMovimentoDiCaricoLav = Mov.Insert

Exit Function
ERR_GeneraMovimentoDiCaricoLav:
    MsgBox Err.Description, vbCritical, "ERR_GeneraMovimentoDiCaricoLav"


End Function

Private Function GeneraMovimentoDiScaricoLav(IDAssegnazione As Long, Qta_UM As Double, DataLavorazione As String) As Boolean

On Error GoTo ERR_GeneraMovimentoDiScaricoLav
Mov.DataMovimento = DataLavorazione
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = Link_Esercizio
Mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
Mov.IDOggetto = IDRigaConf
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 2)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoUscita = Link_Magazzino_Conferimento
Mov.IDMagazzinoEntrata = Link_Magazzino_Conferimento
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", IDAnagrafica
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticoloConf
Mov.Field "IDUnitaDiMisura", IDUMConf
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", ArticoloConf
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", 0
Mov.Field "DataDocumento", DataLavorazione
Mov.Field "Oggetto", "Lavorazione merce del " & DataLavorazione
Mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
Mov.Field "RV_POTipoRiga", 0
Mov.Field "RV_POIDCaricoMerceRighe", 0
Mov.Field "RV_POIDAssegnazioneMerce", 0
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", 0
Mov.Field "RV_PODataConferimento", ""
Mov.Field "RV_PONumeroConferimento", ""
Mov.Field "RV_POCodiceLotto", ""
Mov.Field "RV_POCodiceLottoCampagna", ""
Mov.Field "RV_POCodiceLottoVendita", ""
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POQuantitaMovimentata", 0
Mov.Field "RV_PONumeroColli", 0
Mov.Field "RV_POPesoLordo", 0
Mov.Field "RV_POPesoNetto", 0
Mov.Field "RV_POTara", 0
Mov.Field "RV_POQuantitaPezzi", 0

Mov.Field "RV_POIDTipoDocumentoCoop", 0
Mov.Field "RV_POIDImballo", 0
Mov.Field "RV_POCodiceImballo", ""
Mov.Field "RV_PODescrizioneImballo", ""

Mov.Field "RV_PODataLavorazione", ""
Mov.Field "RV_POIDTipoLavorazione", 0
Mov.Field "RV_POIDCalibro", 0
Mov.Field "RV_POIDTipoCategoria", 0
Mov.Field "RV_POIDTipoLavorazioneConf", 0
Mov.Field "RV_POPrezzoMedioConf", 0

Mov.Field "RV_POIDPedana", 0
Mov.Field "RV_POIDTipoPedana", 0
Mov.Field "RV_POCodicePedana", ""
Mov.Field "RV_POPesoPedana", 0

Mov.Field "TipoRiga", trcNessuno

GeneraMovimentoDiScaricoLav = Mov.Insert

Exit Function
ERR_GeneraMovimentoDiScaricoLav:
    MsgBox Err.Description, vbCritical, "GeneraMovimentoDiScaricoLav"
End Function

Public Function GeneraMovimentoCaricoImballoLav(IDRigaConferimento As Long, IDAssegnazione As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double) As Boolean
On Error GoTo ERR_GeneraMovimentoCaricoImballoLav
Mov.DataMovimento = Date
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = Link_Esercizio
Mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
Mov.IDOggetto = IDRigaConferimento
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Link_Magazzino_Vendita
Mov.IDMagazzinoUscita = Link_Magazzino_Vendita
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", IDAnagraficaSocio
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUMDiamante
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", 0
Mov.Field "DataDocumento", Date
Mov.Field "Oggetto", "Lavorazione merce del " & Date
Mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
Mov.Field "RV_POTipoRiga", 2
Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
Mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocio
Mov.Field "RV_PODataConferimento", DataConf
Mov.Field "RV_PONumeroConferimento", NumeroConf
Mov.Field "RV_POCodiceLotto", CodiceLottoEntrataConf
Mov.Field "RV_POCodiceLottoCampagna", ""
Mov.Field "RV_POCodiceLottoVendita", ""
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POQuantitaMovimentata", 0
Mov.Field "RV_PONumeroColli", 0
Mov.Field "RV_POPesoLordo", 0
Mov.Field "RV_POPesoNetto", 0
Mov.Field "RV_POTara", 0
Mov.Field "RV_POQuantitaPezzi", 0

Mov.Field "RV_POIDTipoDocumentoCoop", 0
Mov.Field "RV_POIDImballo", 0
Mov.Field "RV_POCodiceImballo", ""
Mov.Field "RV_PODescrizioneImballo", ""

Mov.Field "RV_PODataLavorazione", ""
Mov.Field "RV_POIDTipoLavorazione", 0
Mov.Field "RV_POIDCalibro", 0
Mov.Field "RV_POIDTipoCategoria", 0
Mov.Field "RV_POIDTipoLavorazioneConf", 0
Mov.Field "RV_POPrezzoMedioConf", 0

Mov.Field "RV_POIDPedana", 0
Mov.Field "RV_POIDTipoPedana", 0
Mov.Field "RV_POCodicePedana", ""
Mov.Field "RV_POPesoPedana", 0


Mov.Field "TipoRiga", trcNessuno

GeneraMovimentoCaricoImballoLav = Mov.Insert

Exit Function
ERR_GeneraMovimentoCaricoImballoLav:
    MsgBox Err.Description, vbCritical, "GeneraMovimentoCaricoImballoLav"
End Function


Public Function GeneraMovimentoScaricoImballoLav(IDRigaConferimento As Long, IDAssegnazione As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, TracciaImballo As Long) As Boolean
On Error GoTo ERR_GeneraMovimentoScaricoImballoLav
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double
Dim sSQL As String


    Mov.DataMovimento = Date
    Mov.FattoreDiConversione = Null
    
    Mov.GestioneMatricole = False
    Mov.IDEsercizio = Link_Esercizio
    Mov.IDTipoOggetto = fnGetTipoOggetto("RV_POAssegnazioneMerce")
    Mov.IDOggetto = IDRigaConferimento
    Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(10, 2)
    Mov.IDUtente = TheApp.IDUser
    Mov.IDMagazzinoUscita = Link_Magazzino_Conferimento
    Mov.IDMagazzinoEntrata = Link_Magazzino_Conferimento
    Mov.Cessione = 0
    Mov.Field "IDAzienda", TheApp.IDFirm
    Mov.Field "IDAnagrafica", IDAnagraficaSocioConf
    Mov.Field "IDTipoAnagrafica", 3
    Mov.Field "IDArticolo", IDArticolo
    Mov.Field "IDUnitaDiMisura", IDUMDiamante
    Mov.Field "IDcambio", Null
    Mov.Field "DescrizioneArticolo", Articolo
    Mov.Field "QuantitaTotale", Qta_UM
    Mov.Field "Importo", 0
    Mov.Field "DataDocumento", Date
    Mov.Field "Oggetto", "Lavorazione merce del " & Date
    Mov.Field "IDTipoMovimento", 1
    
    'DATI DI CONFERIMENTO
    Mov.Field "IDValoriOggettoDettaglio", IDAssegnazione
    Mov.Field "RV_POTipoRiga", 2
    Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
    Mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
    Mov.Field "RV_POIDProcessoIVGamma", 0
    Mov.Field "RV_POIDAnagraficaSocio", IDAnagraficaSocioConf
    Mov.Field "RV_PODataConferimento", DataConf
    Mov.Field "RV_PONumeroConferimento", NumeroConf
    Mov.Field "RV_POCodiceLotto", CodiceLottoEntrataConf
    Mov.Field "RV_POCodiceLottoCampagna", ""
    Mov.Field "RV_POCodiceLottoVendita", ""
    Mov.Field "RV_POQuantitaLiquidazione", 0
    Mov.Field "RV_POImportoInclusoImballo", 0
    Mov.Field "RV_POImportoLiquidazione", 0
    Mov.Field "RV_POQuantitaMovimentata", 0
    Mov.Field "RV_PONumeroColli", 0
    Mov.Field "RV_POPesoLordo", 0
    Mov.Field "RV_POPesoNetto", 0
    Mov.Field "RV_POTara", 0
    Mov.Field "RV_POQuantitaPezzi", 0
    
    Mov.Field "RV_POIDTipoDocumentoCoop", 0
    Mov.Field "RV_POIDImballo", 0
    Mov.Field "RV_POCodiceImballo", ""
    Mov.Field "RV_PODescrizioneImballo", ""
    
    Mov.Field "RV_PODataLavorazione", ""
    Mov.Field "RV_POIDTipoLavorazione", 0
    Mov.Field "RV_POIDCalibro", 0
    Mov.Field "RV_POIDTipoCategoria", 0
    Mov.Field "RV_POIDTipoLavorazioneConf", 0
    Mov.Field "RV_POPrezzoMedioConf", 0
    
    Mov.Field "RV_POIDPedana", 0
    Mov.Field "RV_POIDTipoPedana", 0
    Mov.Field "RV_POCodicePedana", ""
    Mov.Field "RV_POPesoPedana", 0
    Mov.Field "RV_POIDLottoImballo", 0
    Mov.Field "LottoImballo", ""

    
    Mov.Field "TipoRiga", trcNessuno
    

    
    GeneraMovimentoScaricoImballoLav = Mov.Insert
    

        
    


Exit Function
ERR_GeneraMovimentoScaricoImballoLav:
    MsgBox Err.Description, vbCritical, "GeneraMovimentoScaricoImballoLav"
End Function
Private Function MOVIMENTAZIONE_RIGA_LAVORAZIONE(IDLavorazione As Long) As String
On Error GoTo ERR_MOVIMENTAZIONE_RIGA_LAVORAZIONE
Dim OLD_CURSOR As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
    
MOVIMENTAZIONE_RIGA_LAVORAZIONE = ""
    
Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection

'''''''''''''''''''''''ELIMINAZIONE MOVIMENTI DELLA RIGA DI LAVORAZIONE'''''''''''
sSQL = "SELECT IDMovimento FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & fnGetTipoOggetto
sSQL = sSQL & " AND IDOggetto=" & IDRigaConf
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDLavorazione

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If Mov.Delete(fnNotNullN(rs!IDMovimento)) = False Then
        MOVIMENTAZIONE_RIGA_LAVORAZIONE = "Problema riscontrato con l'eliminazione della movimentazione della riga di lavorazione"
    End If
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then

    If GeneraMovimentoDiCaricoLav(IDLavorazione, IDRigaConf, fnNotNull(rs!CodiceLottoVendita), "", fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDUnitaDiMisura), _
        fnNotNull(rs!Articolo), fnNotNullN(rs!Qta_UM), Date, fnNotNullN(rs!Qta_UM), _
        fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi), _
        fnNotNullN(rs!IDImballoVendita), 1, fnNotNull(rs!CodiceImballoVendita), fnNotNull(rs!ImballoVendita), _
        fnNotNullN(rs!IDTipoLavorazione), fnNotNullN(rs!IDRV_POTipoCategoria), fnNotNullN(rs!IDRV_POCalibro), 0, 0, _
        fnNotNullN(rs!IDRV_POPedana), GET_TIPO_PEDANA(fnNotNullN(rs!IDRV_POPedana)), fnNotNull(rs!CodicePedana), 0) = False Then
    
        MOVIMENTAZIONE_RIGA_LAVORAZIONE = MOVIMENTAZIONE_RIGA_LAVORAZIONE & "Problema riscontrato con la movimentazione della riga di lavorazione" & vbCrLf
    
    End If
    
    If GeneraMovimentoDiScaricoLav(IDLavorazione, fnNotNullN(rs!Colli), Date) = False Then
        MOVIMENTAZIONE_RIGA_LAVORAZIONE = MOVIMENTAZIONE_RIGA_LAVORAZIONE & "Problema riscontrato con la movimentazione della riga di conferimento" & vbCrLf
    End If
    
    
    If ((fnNotNullN(rs!IDImballoVendita) > 0) And (fnNotNullN(rs!Colli) > 0)) Then
        If GeneraMovimentoCaricoImballoLav(IDRigaConf, IDLavorazione, CodiceLottoEntrataConf, "", fnNotNullN(rs!IDImballoVendita), 8, fnNotNull(rs!ImballoVendita), fnNotNullN(rs!Colli)) = False Then
            MOVIMENTAZIONE_RIGA_LAVORAZIONE = MOVIMENTAZIONE_RIGA_LAVORAZIONE & "Problema riscontrato con la movimentazione della riga di carico imballo" & vbCrLf
    
        End If
        If GeneraMovimentoScaricoImballoLav(IDRigaConf, IDLavorazione, CodiceLottoEntrataConf, "", fnNotNullN(rs!IDImballoVendita), 8, fnNotNull(rs!ImballoVendita), fnNotNullN(rs!Colli), 0) = False Then
            MOVIMENTAZIONE_RIGA_LAVORAZIONE = MOVIMENTAZIONE_RIGA_LAVORAZIONE & "Problema riscontrato con la movimentazione della riga di scarico imballo" & vbCrLf
        End If
    End If

End If

rs.CloseResultset
Set rs = Nothing

Set Mov = Nothing
Exit Function
ERR_MOVIMENTAZIONE_RIGA_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, "MOVIMENTAZIONE_RIGA_LAVORAZIONE"
End Function
Private Function GET_TIPO_PEDANA(IDPedana As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPedana "
sSQL = sSQL & "FROM RV_POPedana "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PEDANA = 0
Else
    GET_TIPO_PEDANA = fnNotNullN(rs!IDRV_POTipoPedana)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_ANA_SOCIO(IDAnagrafica As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica "

rs.CloseResultset
Set rs = Nothing
End Function
