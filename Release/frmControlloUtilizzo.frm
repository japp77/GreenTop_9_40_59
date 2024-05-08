VERSION 5.00
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Begin VB.Form frmControlloUtilizzo 
   Caption         =   "CONTROLLO UTILIZZO FUNZIONALITA'"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18960
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControlloUtilizzo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   18960
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   11415
      Left            =   0
      ScaleHeight     =   11385
      ScaleWidth      =   18945
      TabIndex        =   2
      Top             =   0
      Width           =   18975
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   11175
         Left            =   120
         ScaleHeight     =   11145
         ScaleWidth      =   18705
         TabIndex        =   3
         Top             =   120
         Width           =   18735
         Begin VB.CommandButton Command1 
            Caption         =   "AVVIA CONTROLLO"
            Height          =   495
            Left            =   14400
            TabIndex        =   4
            Top             =   10560
            Width           =   4215
         End
         Begin DmtGridCtl.DmtGrid Griglia 
            Height          =   10335
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   18495
            _ExtentX        =   32623
            _ExtentY        =   18230
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
   End
End
Attribute VB_Name = "frmControlloUtilizzo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGriglia As ADODB.Recordset

Private Posizione As Integer



Private Sub CREA_RECORDSET()
On Error GoTo ERR_CREA_RECORDSET
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long

Set rsGriglia = New ADODB.Recordset
rsGriglia.CursorLocation = adUseClient

rsGriglia.Fields.Append "ID", adInteger
rsGriglia.Fields.Append "Descrizione", adVarChar, 250
rsGriglia.Fields.Append "Utilizzato", adBoolean
rsGriglia.Fields.Append "NumeroRecord", adDecimal
rsGriglia.Fields.Append "DataUltimoUtilizzo", adDBDate, , adFldIsNullable
rsGriglia.Fields.Append "Posizione", adInteger
rsGriglia.Fields.Append "TabellaDaControllare", adVarChar, 250
rsGriglia.Fields.Append "CampoPerNumeroRecord", adVarChar, 250
rsGriglia.Fields.Append "CampoPerDataUltimoUtilizzo", adVarChar, 250
rsGriglia.Fields.Append "ScriptWhere", adVarChar, 1000
rsGriglia.Fields.Append "MacroFunzionalita", adVarChar, 250

rsGriglia.Open , , adOpenKeyset, adLockBatchOptimistic


ADDRECORD 5, "AZIENDA AGRICOLA", "Terreni", "RV_PO01_TerrenoTesta", "IDRV_PO01_TerrenoTesta", "", ""
ADDRECORD 6, "AZIENDA AGRICOLA", "Particelle catastali", "RV_PO01_TerrenoRighe", "IDRV_PO01_TerrenoRighe", "", ""
ADDRECORD 7, "AZIENDA AGRICOLA", "Unità produttive", "RV_PO01_Serra", "IDRV_PO01_Serra", "", ""
ADDRECORD 8, "AZIENDA AGRICOLA", "Periodo di campagna", "RV_PO01_PeriodoCampagna", "IDRV_PO01_PeriodoCampagna", "", ""
ADDRECORD 9, "AZIENDA AGRICOLA", "Certificazioni", "RV_PO01_Certificazione", "IDRV_PO01_Certificazione", "", ""
ADDRECORD 10, "AZIENDA AGRICOLA", "Lotto di produzione", "RV_PO01_LottoCampagna", "IDRV_PO01_LottoCampagna", "DataUltimaModifica", ""
ADDRECORD 12, "AZIENDA AGRICOLA", "Lotto di produzione - Lista articoli", "RV_PO01_DettaglioLotto", "IDRV_PO01_DettaglioLotto", "DataUltimaModifica", ""
ADDRECORD 13, "AZIENDA AGRICOLA", "Lotto di produzione - Unità produttive", "RV_PO01_SerraPerLotto", "IDRV_PO01_SerraPerLotto", "", ""
ADDRECORD 14, "AZIENDA AGRICOLA", "Lotto di produzione - Certificazioni", "RV_PO01_LottoCampagna", "IDRV_PO01_CertificazioneSocio", "", "NOT(IDRV_PO01_CertificazioneSocio IS NULL) AND (IDRV_PO01_CertificazioneSocio > 0)"
ADDRECORD 14, "AZIENDA AGRICOLA", "Lotto di produzione - Verifiche", "RV_PO01_LottoCampagnaVerifica", "IDRV_PO01_LottoCampagnaVerifica", "", ""
ADDRECORD 15, "CONFIGURAZIONE FORNITORE/SOCIO", "Configurazione socio", "RV_PO01_ConfigurazioneSocio", "IDRV_PO01_ConfigurazioneSocio", "", ""
ADDRECORD 16, "CONFIGURAZIONE FORNITORE/SOCIO", "Configurazione socio - Fatturazione elettronica", "RV_PO01_ConfigurazioneSocio", "IDRV_PO01_ConfigurazioneSocio", "", "NOT(IDAziendaFE IS NULL) AND (IDAziendaFE>0)"
ADDRECORD 17, "CONFIGURAZIONE FORNITORE/SOCIO", "Configurazione socio - Certificazione generale", "RV_PO01_CertificazioneSocio", "IDRV_PO01_CertificazioneSocio", "", ""
ADDRECORD 18, "CONFIGURAZIONE FORNITORE/SOCIO", "Configurazione socio - Certificazione per famiglia", "RV_PO01_CertificazioneSocioFamiglia", "IDRV_PO01_CertificazioneSocioFamiglia", "", ""

ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente", "RV_POConfigurazioneCliente", "IDRV_POConfigurazioneCliente", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Trattenute", "RV_POConfigurazioneClienteArtTratt", "IDRV_POConfigurazioneClienteArtTratt", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Articoli vendita", "RV_POConfigurazioneClienteArtVend", "IDRV_POConfigurazioneClienteArtVend", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Codice EAN13", "RV_POConfigurazioneClienteEAN13", "IDRV_POConfigurazioneClienteEAN13", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Imballi", "RV_POConfigurazioneClienteImb", "IDRV_POConfigurazioneClienteImb", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Cauzioni", "RV_POConfigurazioneClienteImbCauz", "IDRV_POConfigurazioneClienteImbCauz", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Listino", "RV_POConfigurazioneClienteListino", "IDRV_POConfigurazioneClienteListino", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Modalità di pagamento", "RV_POConfigurazioneClientePagamenti", "IDRV_POConfigurazioneClientePagamenti", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Numerazione pedana", "RV_POConfigurazioneClientePedana", "IDRV_POConfigurazioneClientePedana", "", ""
ADDRECORD 15, "CONFIGURAZIONE CLIENTE", "Configurazione cliente - Trasporto", "RV_POConfigurazioneClienteTrasporto", "IDRV_POConfigurazioneClienteTrasporto", "", ""


ADDRECORD 1, "ENTRATA MERCE", "Conferimento", "RV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", "DataDocumento", "IDTipoDocumentoCoop=1 AND PreConferimento = 0"
ADDRECORD 2, "ENTRATA MERCE", "Conferimento Acquisto", "RV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", "DataDocumento", "IDTipoDocumentoCoop=2 AND PreConferimento = 0"
ADDRECORD 1, "ENTRATA MERCE", "Preconferimento", "RV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", "DataDocumento", "IDTipoDocumentoCoop=1 AND PreConferimento = 1"
ADDRECORD 2, "ENTRATA MERCE", "Perconferimento di acquisto", "RV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", "DataDocumento", "IDTipoDocumentoCoop=2 AND PreConferimento = 1"
ADDRECORD 2, "ENTRATA MERCE", "Pesature", "RV_POCaricoMerceRighePes", "IDRV_POCaricoMerceRighePes", "", ""
ADDRECORD 2, "ENTRATA MERCE", "Campionatura", "RV_POCampionatura", "IDRV_POCampionatura", "", ""
ADDRECORD 2, "ENTRATA MERCE", "Gestione imballi", "RV_POCaricoMerceImballi", "IDRV_POCaricoMerceImballi", "", ""
ADDRECORD 2, "ENTRATA MERCE", "Altri addebiti", "RV_POCaricoMerceAddebiti", "IDRV_POCaricoMerceAddebiti", "", ""
ADDRECORD 2, "ENTRATA MERCE", "Ordine cliente per lavorazione automatica", "RV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", "", "NOT(IDOggettoOrdine IS NULL) AND (IDOggettoOrdine>0)"
ADDRECORD 3, "LAVORAZIONE", "Lavorazione", "RV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce", "DataDocumento", ""
ADDRECORD 3, "LAVORAZIONE", "Quadrature", "RV_POLavorazione", "IDRV_POLavorazione", "DataDocumento", ""
ADDRECORD 3, "LAVORAZIONE", "Processo di IV Gamma", "RV_POProcessoIVGamma", "IDRV_POProcessoIVGamma", "", ""
ADDRECORD 3, "LAVORAZIONE", "Processo di IV Gamma in entrata", "RV_POProcessoIVGammaRighe", "IDRV_POProcessoIVGammaRighe", "", ""
ADDRECORD 3, "LAVORAZIONE", "Processo di IV Gamma in uscita", "RV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce", "", "NOT(IDRV_POProcessoIVGamma IS NULL) AND (IDRV_POProcessoIVGamma>0)"
'ADDRECORD 3, "LAVORAZIONE", "Spaccature di merce lavorata", "RV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce", "", "IDRV_POAssegnazioneMercePadre <> IDRV_POAssegnazioneMerce"
ADDRECORD 3, "LAVORAZIONE", "Collegamento alle righe ordine", "RV_POAssegnazioneMerce", "IDRV_POAssegnazioneMerce", "", "NOT(IDValoriOggettoDettaglioRigaOrd IS NULL) AND (IDValoriOggettoDettaglioRigaOrd>0)"
ADDRECORD 3, "LAVORAZIONE", "Gestione KIT", "RV_POAssegnazioneMerceImbPrim", "ID", "", ""
ADDRECORD 4, "CONTRATTO", "Contratti", "ValoriOggettoPerTipo000E", "IDOggetto", "doc_data", ""
ADDRECORD 4, "CONTRATTO", "Contratti con corpo", "ValoriOggettoDettaglio0038", "IDValoriOggettoDettaglio", "", ""
ADDRECORD 4, "ORDINI", "Ordini a cliente", "ValoriOggettoPerTipo000F", "IDOggetto", "doc_data", ""
ADDRECORD 4, "ORDINI", "Ordini con corpo", "ValoriOggettoDettaglio0010", "IDValoriOggettoDettaglio", "", ""
ADDRECORD 4, "ORDINI", "Ordini collegati a contratto", "ValoriOggettoPerTipo000F", "IDOggetto", "doc_data", "NOT(RV_POIDOggettoContratto IS NULL) AND (RV_POIDOggettoContratto>0)"
ADDRECORD 4, "ORDINI", "Ordini con rettifiche", "RV_POCorpoOrdineRettifica", "ID", "", ""
ADDRECORD 4, "ORDINI", "Ordini con commissioni", "RV_POCommissioniPerDoc INNER JOIN Oggetto ON RV_POCommissioniPerDoc.IDOggetto = Oggetto.IDOggetto", "RV_POCommissioniPerDoc.IDOggetto", "", "Oggetto.IDTipoOggetto=15"

ADDRECORD 4, "VENDITA", "Documento di trasporto", "ValoriOggettoPerTipo0002", "IDOggetto", "doc_data", ""
ADDRECORD 4, "VENDITA", "Documento di trasporto - Triangolazione", "ValoriOggettoPerTipo0002", "IDOggetto", "doc_data", "NOT(RV_POIDAnagraficaDestinazione IS NULL) AND (RV_POIDAnagraficaDestinazione>0)"
ADDRECORD 4, "VENDITA", "Documento di trasporto - Quadrangolazione", "ValoriOggettoPerTipo0002", "IDOggetto", "doc_data", "NOT(RV_POIDAnagraficaSocio IS NULL) AND (RV_POIDAnagraficaSocio>0)"
ADDRECORD 4, "VENDITA", "Documento di trasporto - Lettere d'intento", "ValoriOggettoPerTipo0002", "IDOggetto", "", "NOT(Link_Nom_lettera_intento IS NULL) AND (Link_Nom_lettera_intento>0)"
ADDRECORD 4, "VENDITA", "Documento di trasporto - Protocollo ICE", "ValoriOggettoPerTipo0002", "IDOggetto", "", "NOT(RV_POIDProtocolloICE IS NULL) AND (RV_POIDProtocolloICE>0)"
ADDRECORD 4, "VENDITA", "Documento di trasporto - Non in valuta", "ValoriOggettoPerTipo0002", "IDOggetto", "", "Link_Val_valuta<>9"
ADDRECORD 4, "VENDITA", "Documento di trasporto - Commissioni", "RV_POCommissioniPerDoc INNER JOIN Oggetto ON RV_POCommissioniPerDoc.IDOggetto = Oggetto.IDOggetto", "RV_POCommissioniPerDoc.IDOggetto", "", "Oggetto.IDTipoOggetto=2"
ADDRECORD 4, "VENDITA", "Documento di trasporto (Corpo) - Collegamento entrata merce", "ValoriOggettoDettaglio0004", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDConferimentoRighe IS NULL) AND (RV_POIDConferimentoRighe>0)"
ADDRECORD 4, "VENDITA", "Documento di trasporto (Corpo) - Collegamento lavorazione prima gamma", "ValoriOggettoDettaglio0004", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDAssegnazioneMerce IS NULL) AND (RV_POIDAssegnazioneMerce>0)"
ADDRECORD 4, "VENDITA", "Documento di trasporto (Corpo) - Collegamento lavorazione quarta gamma", "ValoriOggettoDettaglio0004", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDProcessoIVGamma IS NULL) AND (RV_POIDProcessoIVGamma>0)"
ADDRECORD 4, "VENDITA", "Documento di trasporto (Corpo) - Riscontro peso", "ValoriOggettoDettaglio0004", "IDValoriOggettoDettaglio", "", "(RV_PORigaRiscontroPeso=1)"

ADDRECORD 4, "VENDITA", "Fattura accompagnatoria", "ValoriOggettoPerTipo0072", "IDOggetto", "doc_data", ""
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria - Triangolazione", "ValoriOggettoPerTipo0072", "IDOggetto", "doc_data", "NOT(RV_POIDAnagraficaDestinazione IS NULL) AND (RV_POIDAnagraficaDestinazione>0)"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria - Quadrangolazione", "ValoriOggettoPerTipo0072", "IDOggetto", "doc_data", "NOT(RV_POIDAnagraficaSocio IS NULL) AND (RV_POIDAnagraficaSocio>0)"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria - Lettere d'intento", "ValoriOggettoPerTipo0072", "IDOggetto", "", "NOT(Link_Nom_lettera_intento IS NULL) AND (Link_Nom_lettera_intento>0)"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria - Protocollo ICE", "ValoriOggettoPerTipo0072", "IDOggetto", "", "NOT(RV_POIDProtocolloICE IS NULL) AND (RV_POIDProtocolloICE>0)"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria - Non in valuta", "ValoriOggettoPerTipo0072", "IDOggetto", "", "Link_Val_valuta<>9"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria - Commissioni", "RV_POCommissioniPerDoc INNER JOIN Oggetto ON RV_POCommissioniPerDoc.IDOggetto = Oggetto.IDOggetto", "RV_POCommissioniPerDoc.IDOggetto", "", "Oggetto.IDTipoOggetto=114"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria (Corpo) - Collegamento entrata merce", "ValoriOggettoDettaglio0001", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDConferimentoRighe IS NULL) AND (RV_POIDConferimentoRighe>0)"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria (Corpo) - Collegamento lavorazione prima gamma", "ValoriOggettoDettaglio0001", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDAssegnazioneMerce IS NULL) AND (RV_POIDAssegnazioneMerce>0)"
ADDRECORD 4, "VENDITA", "Fattura accompagnatoria (Corpo) - Collegamento lavorazione quarta gamma", "ValoriOggettoDettaglio0001", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDProcessoIVGamma IS NULL) AND (RV_POIDProcessoIVGamma>0)"

ADDRECORD 4, "VENDITA", "Buono di consegna", "ValoriOggettoPerTipo0008", "IDOggetto", "doc_data", ""
ADDRECORD 4, "VENDITA", "Buono di consegna - Triangolazione", "ValoriOggettoPerTipo0008", "IDOggetto", "doc_data", "NOT(RV_POIDAnagraficaDestinazione IS NULL) AND (RV_POIDAnagraficaDestinazione>0)"
ADDRECORD 4, "VENDITA", "Buono di consegna - Quadrangolazione", "ValoriOggettoPerTipo0008", "IDOggetto", "doc_data", "NOT(RV_POIDAnagraficaSocio IS NULL) AND (RV_POIDAnagraficaSocio>0)"
ADDRECORD 4, "VENDITA", "Buono di consegna - Lettere d'intento", "ValoriOggettoPerTipo0008", "IDOggetto", "", "NOT(Link_Nom_lettera_intento IS NULL) AND (Link_Nom_lettera_intento>0)"
ADDRECORD 4, "VENDITA", "Buono di consegna - Protocollo ICE", "ValoriOggettoPerTipo0008", "IDOggetto", "", "NOT(RV_POIDProtocolloICE IS NULL) AND (RV_POIDProtocolloICE>0)"
ADDRECORD 4, "VENDITA", "Buono di consegna - Non in valuta", "ValoriOggettoPerTipo0008", "IDOggetto", "", "Link_Val_valuta<>9"
ADDRECORD 4, "VENDITA", "Buono di consegna - Commissioni", "RV_POCommissioniPerDoc INNER JOIN Oggetto ON RV_POCommissioniPerDoc.IDOggetto = Oggetto.IDOggetto", "RV_POCommissioniPerDoc.IDOggetto", "", "Oggetto.IDTipoOggetto=8"
ADDRECORD 4, "VENDITA", "Buono di consegna (Corpo) - Collegamento entrata merce", "ValoriOggettoDettaglio0034", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDConferimentoRighe IS NULL) AND (RV_POIDConferimentoRighe>0)"
ADDRECORD 4, "VENDITA", "Buono di consegna (Corpo) - Collegamento lavorazione prima gamma", "ValoriOggettoDettaglio0034", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDAssegnazioneMerce IS NULL) AND (RV_POIDAssegnazioneMerce>0)"
ADDRECORD 4, "VENDITA", "Buono di consegna (Corpo) - Collegamento lavorazione quarta gamma", "ValoriOggettoDettaglio0034", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDProcessoIVGamma IS NULL) AND (RV_POIDProcessoIVGamma>0)"
ADDRECORD 4, "VENDITA", "Buono di consegna (Corpo) - Invenduto", "ValoriOggettoDettaglio0034", "IDValoriOggettoDettaglio", "", "NOT(RV_POInvenduto IS NULL) AND (RV_POInvenduto>0)"

ADDRECORD 4, "VENDITA", "Nota di credito", "ValoriOggettoPerTipo000B", "IDOggetto", "doc_data", ""
ADDRECORD 4, "VENDITA", "Nota di credito (Corpo) - Collegamento entrata merce", "ValoriOggettoDettaglio0016", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDConferimentoRighe IS NULL) AND (RV_POIDConferimentoRighe>0)"

ADDRECORD 4, "VENDITA", "Nota di debito", "ValoriOggettoPerTipo006B", "IDOggetto", "doc_data", ""
ADDRECORD 4, "VENDITA", "Nota di debito (Corpo) - Collegamento entrata merce", "ValoriOggettoDettaglio0007", "IDValoriOggettoDettaglio", "", "NOT(RV_POIDConferimentoRighe IS NULL) AND (RV_POIDConferimentoRighe>0)"
ADDRECORD 4, "TOUR", "Gestione", "RV_POTour", "IDRV_POTour", "", ""
ADDRECORD 4, "TOUR", "Collegamenti a documenti", "RV_POTourRighe", "IDRV_POTourRighe", "", ""

ADDRECORD 4, "ETICHETTE PROFESSIONAL", "Etichette di pedana", "RV_POConfigurazioneEtichetta", "IDRV_POConfigurazioneEtichetta", "", "IDRV_POTipoEtichetta=2"
ADDRECORD 4, "ETICHETTE PROFESSIONAL", "Etichette di lavorazione", "RV_POConfigurazioneEtichetta", "IDRV_POConfigurazioneEtichetta", "", "IDRV_POTipoEtichetta=1"
ADDRECORD 4, "ETICHETTE PROFESSIONAL", "Configurazione codici EAN128", "RV_POCostruzioneCodiceEAN128", "IDRV_POCostruzioneCodiceEAN128", "", ""

ADDRECORD 4, "PRODUZIONE", "Linea di lavorazione", "RV_POProcessoLavorazione", "ID", "", ""
ADDRECORD 4, "PRODUZIONE", "Ordini in lavorazione", "RV_POOrdiniInLavorazione", "ID", "", ""
ADDRECORD 4, "PRODUZIONE", "Prelievi da conferimento", "RV_POCaricoMerceRighePrelievi", "ID", "", ""
ADDRECORD 4, "PRODUZIONE", "Controllo accessi", "RV_POCaricoMerceRigheControlloAccessi", "ID", "", ""
ADDRECORD 4, "PRODUZIONE", "Controllo uscite pedane", "RV_POControlloUscitaPedane", "ID", "", ""

ADDRECORD 4, "LIQUIDAZIONI", "Liquidazioni", "RV_POLiquidazione", "IDRV_POLiquidazione", "", ""
ADDRECORD 4, "LIQUIDAZIONI", "Liquidazioni Ufficiali", "RV_POLiquidazione", "IDRV_POLiquidazione", "", "Ufficiale=1"
ADDRECORD 4, "LIQUIDAZIONI", "Configurazione trattenute", "RV_POTrattenutaPerLiquidazione", "IDRV_POTrattenutaPerLiquidazione", "", ""
ADDRECORD 4, "LIQUIDAZIONI", "Gestione anticipazione", "RV_POAnticipazioniSocio", "IDRV_POAnticipazioniSocio", "", ""

ADDRECORD 4, "GENERALE", "Tracciabilità imballi", "RV_POLottoImballo", "IDRV_POLottoImballo", "", ""
ADDRECORD 4, "GENERALE", "Tracciabilità imballi (No sistema)", "RV_POLottoImballo", "IDRV_POLottoImballo", "", "Sistema=0"
ADDRECORD 4, "GENERALE", "Statistiche imballi", "RV_POSaldoImballo", "IDRV_POSaldoImballo", "", ""
ADDRECORD 4, "GENERALE", "Distinta base articoli", "RV_PODistintaBase", "IDRV_PODistintaBase", "", ""
ADDRECORD 4, "GENERALE", "Distinta base articoli - confezione", "RV_PODistintaBaseRighe", "IDRV_PODistintaBaseRighe", "", ""
ADDRECORD 4, "GENERALE", "Distinta base articoli - Kit", "RV_PODistintaBaseRigheConf", "IDRV_PODistintaBaseRigheConf", "", ""

ADDRECORD 4, "FATTURA ELETTRONICA", "Configurazione - Dati per cliente/Articolo", "DatoFatturaPAClientePerArticolo", "IDDatoFatturaPAClientePerArticolo", "", ""
ADDRECORD 4, "FATTURA ELETTRONICA", "Configurazione - Dati per Articolo", "DatoFatturaPAPerArticolo", "IDDatoFatturaPAPerArticolo", "", ""
ADDRECORD 4, "FATTURA ELETTRONICA", "Configurazione - Dati per cliente", "DatoFatturaPAPerCliente", "IDDatoFatturaPAPerCliente", "", ""
ADDRECORD 4, "FATTURA ELETTRONICA", "Configurazione - Dati XML personalizzati di testa", "DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc", "", "IDBloccoXML=8"

GET_GRIGLIA


Exit Sub
ERR_CREA_RECORDSET:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET"
End Sub
Private Sub ADDRECORD(id As Integer, macrofunz As String, descrizione As String, tabella As String, idtabella As String, campodatatabella As String, scriptwhere As String)

rsGriglia.AddNew

    rsGriglia!id = id
    rsGriglia!descrizione = descrizione
    rsGriglia!Utilizzato = False
    rsGriglia!NumeroRecord = 0
    rsGriglia!DataUltimoUtilizzo = Null
    rsGriglia!Posizione = Posizione
    rsGriglia!TabellaDaControllare = tabella
    rsGriglia!CampoPerNumeroRecord = idtabella
    rsGriglia!CampoPerDataUltimoUtilizzo = campodatatabella
    rsGriglia!scriptwhere = scriptwhere
    rsGriglia!MacroFunzionalita = macrofunz
rsGriglia.Update

Posizione = Posizione + 1

End Sub
Private Sub GET_GRIGLIA()
On Error GoTo ERR_GET_GRIGLIA
Dim OLD_Cursor As Long

OLDCursor = CnDMT.CursorLocation
CnDMT.CursorLocation = 3



With Me.Griglia
    .EnableMove = True
    .UpdatePosition = True
    .BooleanType = dgGraphic
    .SelectionMode = dgSelectCell
    .ColumnsHeader.Clear
        
    .ColumnsHeader.Add "ID", "ID", dgInteger, False, 500, dgAlignRight
    .ColumnsHeader.Add "MacroFunzionalita", "Modulo", dgchar, True, 4500, dgAlignleft
    .ColumnsHeader.Add "Descrizione", "Funzionalità", dgchar, True, 5500, dgAlignleft
    .ColumnsHeader.Add "Utilizzato", "Utilizzato", dgBoolean, True, 1000, dgAligncenter
    Set cl = .ColumnsHeader.Add("NumeroRecord", "Numero record", dgDouble, True, 2000, dgAlignRight)
        cl.BackColor = vbYellow
        cl.FormatOptions.FormatNumericRegionalSettings = False
        cl.FormatOptions.UseFormatControlSettings = False
        cl.FormatOptions.FormatNumericCurSymbol = ""
        cl.FormatOptions.FormatNumericDecSep = ","
        cl.FormatOptions.FormatNumericDecimals = 0
        cl.FormatOptions.FormatNumericThousandSep = "."
    .ColumnsHeader.Add "DataUltimoUtilizzo", "Ultimo utilizzo", dgDate, True, 2500, dgAlignleft
    
    Set .Recordset = rsGriglia
    .Refresh
    .LoadUserSettings
End With

CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_GET_GRIGLIA:
    MsgBox Err.Description, vbCritical, "GET_GRIGLIA"
End Sub

Private Sub Command1_Click()
rsGriglia.MoveFirst

While Not rsGriglia.EOF
    DoEvents
    SET_IMPOSTAZIONI rsGriglia
    DoEvents
rsGriglia.MoveNext
Wend



End Sub

Private Sub Form_Load()
    With HScroll1
      .Max = (Pic1.ScaleWidth)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
      
    End With

    With VScroll1
      .Max = (Pic1.ScaleHeight)
      .LargeChange = .Max \ 10
      .SmallChange = .Max \ 10
    End With


    CREA_RECORDSET
End Sub

Private Sub SET_IMPOSTAZIONI(rstmp As ADODB.Recordset)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(" + rstmp!CampoPerNumeroRecord + ") AS NumeroRecord"
If Len(rstmp!CampoPerDataUltimoUtilizzo) > 0 Then
    sSQL = sSQL & ", MAX(" + rstmp!CampoPerDataUltimoUtilizzo + ") AS DataUltimoRecord "
End If
sSQL = sSQL & " FROM " + rstmp!TabellaDaControllare
If Len(rstmp!scriptwhere) > 0 Then
    sSQL = sSQL & " WHERE " + rstmp!scriptwhere
End If

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    rstmp!NumeroRecord = rs!NumeroRecord
    If Len(rstmp!CampoPerDataUltimoUtilizzo) > 0 Then
        rstmp!DataUltimoUtilizzo = rs!DataUltimoRecord
    End If
    If fnNotNullN(rs!NumeroRecord) > 0 Then
         rstmp!Utilizzato = True
    End If
End If



rs.CloseResultset
Set rs = Nothing


End Sub

Private Sub Form_Resize()
On Error Resume Next
  If Me.WindowState <> 1 Then
            
        If Me.Width > 19200 Then
            Me.Pic1.Width = Me.Width - 240
            Me.Picture2.Width = Me.Pic1.Width - 120
            Me.Command1.Left = Me.Picture2.Width - 120 - Me.Command1.Width
            
            Me.Griglia.Width = Me.Picture2.Width - 120
            
            
        End If
        If Me.Height > 12000 Then
            Me.Pic1.Height = Me.Height - 720
            Me.Picture2.Height = Me.Pic1.Height - 240
            Me.Command1.Top = Me.Picture2.Height - 120 - Me.Command1.Height
            Me.Griglia.Height = Me.Command1.Top - Me.Command1.Height - 120  ' - Me.Frame1.Height
        End If

        If Me.ScaleWidth < Me.Pic1.ScaleWidth Then
            Me.HScroll1.Visible = True
            Me.HScroll1.Top = Me.ScaleHeight - Me.HScroll1.Height
            Me.HScroll1.Left = 0
            
        Else
            Me.HScroll1.Visible = False
        End If
        
        If Me.ScaleHeight < Me.Pic1.ScaleHeight Then
            Me.VScroll1.Visible = True
            Me.VScroll1.Top = 0
            Me.VScroll1.Left = Me.ScaleWidth - Me.VScroll1.Width
            
        Else
            Me.VScroll1.Visible = False
        End If
        
        If (VScroll1.Visible = True) And (HScroll1.Visible = False) Then
            Me.VScroll1.Height = Me.ScaleHeight '- Me.HScroll1.Height
        Else
            Me.VScroll1.Height = Me.ScaleHeight - Me.HScroll1.Height
        End If
        
        If (HScroll1.Visible = True) And (HScroll1.Visible = True) Then
            Me.HScroll1.Width = Me.ScaleWidth '- Me.VScroll1.Width
        Else
            Me.HScroll1.Width = Me.ScaleWidth - Me.VScroll1.Width
        End If
            
        With HScroll1
            .Max = (Pic1.ScaleWidth - Me.ScaleWidth + Me.VScroll1.Width)
            If .Max > 0 Then
                .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With

        With VScroll1
            .Max = (Pic1.ScaleHeight - Me.ScaleHeight + Me.HScroll1.Height)
            If .Max > 0 Then
                 .LargeChange = .Max \ 10
                .SmallChange = .Max \ 10
            End If
        End With
    End If
End Sub
