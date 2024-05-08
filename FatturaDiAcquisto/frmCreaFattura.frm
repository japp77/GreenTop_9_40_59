VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Begin VB.Form frmCreaFattura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CREAZIONE DOCUMENTO DI ACQUISTO"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
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
   ScaleHeight     =   6315
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraTesta 
      Height          =   6255
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5295
      Begin VB.CheckBox chkRipImbMerce 
         Caption         =   "Riporta imballo della merce"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   4800
         Width           =   4935
      End
      Begin VB.CommandButton cmdElabora 
         Caption         =   "CREA DOCUMENTO"
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   5640
         Width           =   1935
      End
      Begin DMTDataCmb.DMTCombo cboFunzione 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboTipoOggetto 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboSezionale 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboPagamento 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTEDITNUMLib.dmtCurrency txtSpeseTrasporto 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   480
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencyDecimalPlaces=   5
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtSpeseImballo 
         Height          =   315
         Left            =   3480
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencyDecimalPlaces=   5
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtSpeseAltre 
         Height          =   315
         Left            =   3480
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencyDecimalPlaces=   5
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtSpeseVarie 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   " 0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencyDecimalPlaces=   5
         CurrencySymbol  =   ""
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboBancaAzienda 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboBancaFornitore 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   4080
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkRipLottoImb 
         Caption         =   "Riporta carico imballo da gestione imballi"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   5160
         Width           =   4935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   120
         X2              =   5040
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   5040
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label1 
         Caption         =   "Banca fornitore"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Banca azienda"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Varie"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Altre"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Imballo"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Trasporto"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   3360
         X2              =   3360
         Y1              =   240
         Y2              =   4560
      End
      Begin VB.Label Label1 
         Caption         =   "Pagamento"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Sezionale"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo documento"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Funzione"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmCreaFattura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Oggetto utilizzato per gestire l'inserimento / variazione del documento (DmtDocs.Dll)
Private objDoc As DmtDocs.cDocument
'Variabile utilizzata per ottenere il nome della tabella di testata del documento
Private sTabellaTestata As String
'Variabile utilizzata per ottenere il nome della tabella di dettaglio del documento
Private sTabellaDettaglio As String
'Variabile utilizzata per ottenere il nome della tabella delle scadenze del documento
Private sTabellaScadenze As String
'Variabile utilizzata per ottenere il nome della tabella del castelletto IVA del documento
Private sTabellaIVA As String


Private Sub cboFunzione_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoOggetto FROM Funzione "
sSQL = sSQL & "WHERE IDFunzione=" & Me.cboFunzione.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboTipoOggetto.WriteOn 0
Else
    Me.cboTipoOggetto.WriteOn fnNotNullN(rs!IDTipoOggetto)
End If

rs.CloseResultset
Set rs = Nothing

End Sub

Private Sub cboTipoOggetto_Click()
    With Me.cboSezionale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT  Sezionale.IDSezionale, Sezionale.Sezionale, RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "FROM Sezionale INNER JOIN "
        .SQL = .SQL & "RegistroIvaPerTipoOggetto ON Sezionale.IDRegistroIva = RegistroIvaPerTipoOggetto.IDRegistroIva AND "
        .SQL = .SQL & "Sezionale.IDFiliale = RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & Me.cboTipoOggetto.CurrentID
        .SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & TheApp.Branch
    End With
End Sub

Private Sub cmdElabora_Click()

Const TestoMsgBox As String = "Creazione documento"
If Me.cboFunzione.CurrentID = 0 Then
    MsgBox "Manca la funzione", vbCritical, TestoMsgBox
    Exit Sub
End If

If Me.cboTipoOggetto.CurrentID = 0 Then
    MsgBox "Manca il tipo di documento", vbCritical, TestoMsgBox
    Exit Sub
End If

If Me.cboSezionale.CurrentID = 0 Then
    MsgBox "Manca il sezionale", vbCritical, TestoMsgBox
    Exit Sub
End If

If Me.cboPagamento.CurrentID = 0 Then
    MsgBox "Manca il tipo di pagamento", vbCritical, TestoMsgBox
    Exit Sub
End If


ELABORAZIONE_DOCUMENTO

Unload Me
End Sub

Private Sub INIT_CONTROLLI()
    
    'Funzione
    With Me.cboFunzione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione"
        .Fill
    End With

    'Tipo oggetto
    With Me.cboTipoOggetto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDTipoOggetto"
        .DisplayField = "Oggetto"
        .SQL = "SELECT * FROM TipoOggetto"
        .Fill
    End With

    'Pagamento
    With Me.cboPagamento
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDPagamento"
        .DisplayField = "Pagamento"
        .SQL = "SELECT * FROM Pagamento ORDER BY Pagamento"
        .Fill
    End With

    With Me.cboBancaAzienda
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDBancaPerAnagrafica"
        .DisplayField = "BancaPerAnagrafica"
        .SQL = "SELECT BancaPerAnagrafica.IDBancaPerAnagrafica, BancaPerAnagrafica.BancaPerAnagrafica "
        .SQL = .SQL & "FROM Anagrafica INNER JOIN "
        .SQL = .SQL & "Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica INNER JOIN "
        .SQL = .SQL & "BancaPerAnagrafica ON Azienda.IDAnagrafica = BancaPerAnagrafica.IDAnagrafica "
        .SQL = .SQL & " WHERE ((BancaPerAnagrafica.IDAzienda = " & TheApp.IDFirm & "))"
        .SQL = .SQL & " ORDER BY BancaPerAnagrafica.BancaPerAnagrafica"
    End With

    With Me.cboBancaFornitore
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDBancaPerAnagrafica"
        .DisplayField = "BancaPerAnagrafica"
        .SQL = "SELECT IDBancaPerAnagrafica, BancaPerAnagrafica FROM BancaPerAnagrafica "
        .SQL = .SQL & " WHERE IDAnagrafica = " & frmMain.CDSocioFatt.KeyFieldID
        .SQL = .SQL & " ORDER BY BancaPerAnagrafica"
    End With
    

    
    If frmMain.cboTipoDocumentoAcq.CurrentID = 1 Then
        Me.cboFunzione.WriteOn LINK_FUNZIONE_FA
    End If
    
    If frmMain.cboTipoDocumentoAcq.CurrentID = 2 Then
        Me.cboFunzione.WriteOn LINK_FUNZIONE_DDT
    End If
    
        
    fnSetSezionale
    Me.cboPagamento.WriteOn GET_LINK_PAGAMENTO(frmMain.CDSocioFatt.KeyFieldID, TheApp.IDFirm)
End Sub
Function fnSetSezionale() As Long
Dim Rw As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT Sezionale.IDSezionale, Sezionale.Sezionale, Sezionale.Prefisso "
sSQL = sSQL & "FROM Sezionale INNER JOIN "
sSQL = sSQL & "DefaultFilialePerTipoOggetto ON Sezionale.IDFiliale = DefaultFilialePerTipoOggetto.IDFiliale AND "
sSQL = sSQL & "Sezionale.IDSezionale = DefaultFilialePerTipoOggetto.IDSezionale "
sSQL = sSQL & "WHERE DefaultFilialePerTipoOggetto.IDFiliale = " & TheApp.Branch
sSQL = sSQL & "AND DefaultFilialePerTipoOggetto.IDTipoOggetto = " & Me.cboTipoOggetto.CurrentID

Set Rw = Cn.OpenResultset(sSQL)

If Not Rw.EOF Then
    Me.cboSezionale.WriteOn fnNotNullN(Rw!IDSezionale)
Else
    Me.cboSezionale.WriteOn 0
End If

Rw.CloseResultset
Set Rw = Nothing

End Function
Private Function fnSetCausaleDocumento() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CausaleTrasporto FROM CausaleTrasportoPerFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & Me.cboFunzione.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    fnSetCausaleDocumento = ""
Else
    fnSetCausaleDocumento = fnNotNull(rs!CausaleTrasporto)
End If

rs.CloseResultset
Set rs = Nothing
    
End Function
Private Sub ELABORAZIONE_DOCUMENTO()
    
    Settaggio
    
    fncTestata
    
    fncRighe LINK_TESTA_DOCUMENTO
    
    InserimentoDMT
    
    
End Sub
Private Sub Settaggio()
On Error GoTo ERR_Settaggio
Set objDoc = New cDocument
    With objDoc
        Set .Connection = Cn
        .IDAzienda = TheApp.IDFirm
        .IDAttivitaAzienda = GET_LINK_ATTIVITA_AZIENDA(TheApp.Branch)
        .IDFiliale = TheApp.Branch
        .SetTipoOggetto Me.cboTipoOggetto.CurrentID
        .IDFunzione = Me.cboFunzione.CurrentID
        .TablesNames objDoc.IDTipoOggetto, sTabellaTestata, sTabellaDettaglio, sTabellaIVA, sTabellaScadenze
        .UseAutomation = True
        .IDEsercizio = frmMain.cboEsercizio.CurrentID
        .IDSezionale = Me.cboSezionale.CurrentID
        .IDTipoAnagrafica = 3
        .IDUtente = TheApp.IDUser
        .Descrizione = GET_DESCRIZIONE_TIPOOGGETTO(Me.cboTipoOggetto.CurrentID)
        .DataEmissione = frmMain.txtDataDocumentoAcq.Text
        .Numero = 0
        
        If .Tables.Count = 0 Then
            .Clear
            .SetTipoOggetto Me.cboTipoOggetto.CurrentID
        Else
            .ClearValues
        End If
    
    End With
Exit Sub
ERR_Settaggio:
    MsgBox Err.Description, vbCritical, "Settaggio"
End Sub
Private Function fncTestata() As Boolean
On Error GoTo ERR_fncTestata
Dim IDListinoDefault As Long
Dim Link_Pagamento As Long
Dim Link_Valuta_Cliente As Long

Dim LINK_VALUTA_NAZIONALE As Long
Dim SPESE_TRASPORTO_CLIENTE As Double


VARErroreFunzione = "fncTestata"
         
         With objDoc.Tables
        
            'Imposta la riga attiva per la tabella di testata
            
            objDoc.Tables(sTabellaTestata).SetActiveRetail 1
            
            objDoc.ReadDataFromCliFo frmMain.CDSocioFatt.KeyFieldID, sTabellaTestata
            
            If frmMain.cboVettore.CurrentID > 0 Then
                objDoc.Field "Link_Doc_spedizione", 3, sTabellaTestata
                objDoc.ReadDataFromCarrier frmMain.cboVettore.CurrentID, MainCarrier, sTabellaTestata
            End If
                        
            .Field "Doc_causale_trasporto", fnSetCausaleDocumento, sTabellaTestata
            .Field "Link_Doc_Magazzino", frmMain.cboMagazzinoConf.CurrentID, sTabellaTestata
            .Field "Link_Doc_sezionale", Me.cboSezionale.CurrentID, sTabellaTestata
            .Field "Doc_prefisso", GET_PREFISSO_SEZ(Me.cboSezionale.CurrentID), sTabellaTestata
            .Field "Doc_data", objDoc.DataEmissione, sTabellaTestata
            .Field "Doc_data_inizio_trasporto", objDoc.DataEmissione, sTabellaTestata
            .Field "Doc_ora_inizio_trasporto", time, sTabellaTestata
            .Field "Spe_trasporto_neutro", Me.txtSpeseTrasporto.Value, sTabellaTestata
            .Field "Spe_imballo_neutro", Me.txtSpeseImballo.Value, sTabellaTestata
            .Field "Spe_altre_neutro", Me.txtSpeseAltre.Value, sTabellaTestata
            .Field "Spe_varie_neutro", Me.txtSpeseVarie.Value, sTabellaTestata
            
            .Field "Doc_crea_scadenze", objDoc.DBDefaults.CreaScadenzeDaDocAcquisto, sTabellaTestata
            
            
            .Field "Doc_data_presso_nom", frmMain.txtDataDocumentoAcq.Text, sTabellaTestata
            .Field "Doc_numero_presso_nom", frmMain.txtNumeroDocumentoAcq.Text, sTabellaTestata
            
             objDoc.ReadIvaFromLetter frmMain.txtIDLetteraIntento.Value, , sTabellaTestata
            
            'PAGAMENTO DOPO I DATI DELL'ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Me.cboPagamento.CurrentID > 0 Then
                objDoc.ReadDataFromPayment Me.cboPagamento.CurrentID
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            'VALUTA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            .Field "Link_Val_valuta", objDoc.DBDefaults.Link_Val_valuta_nazionale, sTabellaTestata
            
            .Field "RV_POIDCaricoMerceTesta", LINK_TESTA_DOCUMENTO, sTabellaTestata
            
            .Field "Link_Doc_contratto_bancario_Az", Me.cboBancaAzienda.CurrentID, sTabellaTestata
            .Field "Link_Nom_contratto_bancario", Me.cboBancaFornitore.CurrentID, sTabellaTestata
        
            End With
        
        fncTestata = True
     
Exit Function
ERR_fncTestata:
    fncTestata = False
    
End Function

Private Function fncRighe(IDConferimentoTesta As Long) As Boolean
On Error GoTo ERR_fncRighe
VARErroreFunzione = "fncRighe"
Dim I As Integer
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


    I = 1
    
    sSQL = "SELECT * FROM RV_POCaricoMerceRighe "
    sSQL = sSQL & " WHERE IDRV_POCaricoMerceTesta=" & IDConferimentoTesta
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        'MERCE
        objDoc.Tables(sTabellaDettaglio).SetActiveRetail I
        objDoc.ReadDataFromArticle fnNotNullN(rs!IDArticolo), sTabellaDettaglio

        objDoc.Field "Art_descrizione", fnNotNull(rs!Articolo), sTabellaDettaglio
        objDoc.Field "Art_quantita_totale", fnNotNullN(rs!Qta_UM), sTabellaDettaglio
        objDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitario), sTabellaDettaglio
        objDoc.Field "Link_art_IVA", fnNotNullN(rs!IDIvaAcquisto), sTabellaDettaglio
        objDoc.Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), sTabellaDettaglio
    
        objDoc.Field "Link_Art_Magazzino", frmMain.cboMagazzinoConf.CurrentID, sTabellaDettaglio
        
        objDoc.Field "Art_numero_colli", fnNotNullN(rs!Colli), sTabellaDettaglio
        objDoc.Field "Art_Peso", fnNotNullN(rs!PesoLordo), sTabellaDettaglio
        objDoc.Field "Art_tara", fnNotNullN(rs!Tara), sTabellaDettaglio
        objDoc.Field "Art_quantita_pezzi", fnNotNullN(rs!Pezzi), sTabellaDettaglio
        objDoc.Field "Link_Art_unita_di_misura", fnNotNullN(rs!IDUnitaDiMisuraDiamante), sTabellaDettaglio
   
        I = I + 1
        'IMBALLO
        If Me.chkRipImbMerce.Value = vbChecked Then
            objDoc.Tables(sTabellaDettaglio).SetActiveRetail I
            objDoc.ReadDataFromArticle fnNotNullN(rs!IDImballo), sTabellaDettaglio
    
            objDoc.Field "Art_descrizione", fnNotNull(rs!DescrizioneImballo), sTabellaDettaglio
            objDoc.Field "Art_quantita_totale", fnNotNullN(rs!Colli), sTabellaDettaglio
            objDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitarioImballo), sTabellaDettaglio
            objDoc.Field "Link_art_IVA", fnNotNullN(rs!IDIvaAcquisto), sTabellaDettaglio
            objDoc.Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), sTabellaDettaglio
        
            objDoc.Field "Link_Art_Magazzino", frmMain.cboMagazzinoConf.CurrentID, sTabellaDettaglio
            
            objDoc.Field "Link_Art_unita_di_misura", GET_UM_ARTICOLO(fnNotNullN(rs!IDImballo)), sTabellaDettaglio
            
            I = I + 1
        End If
    rs.MoveNext
    Wend
rs.CloseResultset
Set rs = Nothing


If Me.chkRipLottoImb.Value = vbChecked Then
    sSQL = "SELECT * FROM RV_POCaricoMerceImballi "
    sSQL = sSQL & " WHERE IDRV_POCaricoMerceTesta=" & IDConferimentoTesta
    sSQL = sSQL & " AND IDRV_POTipoProcessoCoop=1"
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        objDoc.Tables(sTabellaDettaglio).SetActiveRetail I
        objDoc.ReadDataFromArticle fnNotNullN(rs!IDArticolo), sTabellaDettaglio

        'objDoc.Field "Art_descrizione", fnNotNull(rs!Articolo), sTabellaDettaglio
        objDoc.Field "Art_quantita_totale", fnNotNullN(rs!Quantita), sTabellaDettaglio
        'objDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitario), sTabellaDettaglio
        'objDoc.Field "Link_art_IVA", fnNotNullN(rs!IDIvaAcquisto), sTabellaDettaglio
        'objDoc.Field "Art_aliquota_IVA", fnNotNullN(rs!AliquotaIva), sTabellaDettaglio
        objDoc.ReadDataFromPriceList fnNotNullN(objDoc.Field("Link_Doc_listino", , sTabellaTestata))
        objDoc.ReadDataFromDiscountsList
        
        
        objDoc.Field "Link_Art_Magazzino", frmMain.cboMagazzinoConf.CurrentID, sTabellaDettaglio
        
        'objDoc.Field "Art_numero_colli", fnNotNullN(rs!Colli), sTabellaDettaglio
        'objDoc.Field "Art_Peso", fnNotNullN(rs!PesoLordo), sTabellaDettaglio
        'objDoc.Field "Art_tara", fnNotNullN(rs!Tara), sTabellaDettaglio
        'objDoc.Field "Art_quantita_pezzi", fnNotNullN(rs!Pezzi), sTabellaDettaglio
        'objDoc.Field "Link_Art_unita_di_misura", fnNotNullN(rs!IDUnitaDiMisuraDiamante), sTabellaDettaglio
   
        I = I + 1
        
    rs.MoveNext
    Wend

rs.CloseResultset
Set rs = Nothing
End If


fncRighe = True
Exit Function
ERR_fncRighe:
    fncRighe = False

End Function
Private Function InserimentoDMT() As Boolean
On Error GoTo ERR_InserimentoDMT
Dim VarNumeroDoc As String
Dim Link_Oggetto As Long
Dim sSQL As String

Set objDoc.Scadenze = Nothing
objDoc.PerformDocument Nothing
    
'CONTROLLO PLAFOND''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim sMsgPlafond As String
If objDoc.PlafondExceed Then
    sMsgPlafond = objDoc.PlafondLastMessage
    If Len(sMsgPlafond) > 0 Then
        If objDoc.PlafondLastMessageStyle = vbCritical Then
            sbMsgError sMsgPlafond, TheApp.FunctionName
            Screen.MousePointer = 0
            Exit Function
        Else
            If MsgBox(sMsgPlafond, vbInformation + vbYesNo, TheApp.FunctionName) = vbNo Then
                Screen.MousePointer = 0
                Exit Function
            End If
            TestoMessaggio = "ATTENZIONE!!!" & vbCrLf
            TestoMessaggio = TestoMessaggio & "Il documento è stato emesso regolarmente, "
            TestoMessaggio = TestoMessaggio & "tuttavia a causa del plafond superato, "
            TestoMessaggio = TestoMessaggio & "è necessario visionarlo ed applicare le opportune modifiche"
            
            MsgBox TestoMessaggio, vbInformation, TheApp.FunctionName

        End If
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
Screen.MousePointer = vbHourglass
    VarNumeroDoc = objDoc.Insert
Screen.MousePointer = vbDefault
    
Set objDoc = Nothing

Exit Function

ERR_InserimentoDMT:
    InserimentoDMT = False
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    
End Function
Private Function GET_PREFISSO_SEZ(IDSezionale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Prefisso FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDSezionale=" & IDSezionale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREFISSO_SEZ = ""
Else
    GET_PREFISSO_SEZ = Trim(fnNotNull(rs!Prefisso))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_TIPOOGGETTO(IDTipoOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select Oggetto "
    sSQL = sSQL & "FROM TipoOggetto "
    sSQL = sSQL & "WHERE IDTipoOggetto = " & IDTipoOggetto
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        GET_DESCRIZIONE_TIPOOGGETTO = fnNotNull(rs!Oggetto)
    Else
        GET_DESCRIZIONE_TIPOOGGETTO = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
End Function
Private Function GET_LINK_ATTIVITA_AZIENDA(IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAttivitaAzienda FROM Filiale "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ATTIVITA_AZIENDA = 0
Else
    GET_LINK_ATTIVITA_AZIENDA = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Form_Load()
    INIT_CONTROLLI
    
    Me.txtSpeseTrasporto.Value = frmMain.txtImportoTrasporto.Value
    
    
End Sub
Private Function GET_LINK_PAGAMENTO(IDAnagrafica As Long, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDPagamentoDefault FROM Fornitore "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & IDAzienda


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_PAGAMENTO = 0
Else
    GET_LINK_PAGAMENTO = fnNotNullN(rs!IDPagamentoDefault)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_UM_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraAcquisto "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_UM_ARTICOLO = 0
Else
    GET_UM_ARTICOLO = fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
End If

rs.CloseResultset
Set rs = Nothing
End Function

