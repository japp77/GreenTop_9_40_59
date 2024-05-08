Attribute VB_Name = "Globali"
Option Explicit

'Declares
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Sub sbOpenURL Lib "Diamante.dll" (ByVal hwnd As Long, ByVal sURL As String)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOP = 0
Public Const WM_SETREDRAW = &HB

'Costanti globali
Public Const TOTAL_CONTROLS_NUMBER = 10
Public Const SPLITLIMIT = 1000
Public Const SRCNEXT = 1
Public Const SRCPREVIOUS = 2
Public Const HELP_FINDER = &HB
Public Const HELP_CONTEXT = &H1
Public Const URL_DIAMANTE = "http://www.diamante.it"

'*** Costanti per la gestione della Attivazione-Disattivazione Menu e ToolBar
Public Const BTN_NEW = 1
Public Const BTN_SAVE = 2
Public Const BTN_PRINT = 4
Public Const BTN_PREVIEW = 8
Public Const BTN_CUT = 16
Public Const BTN_COPY = 32
Public Const BTN_PASTE = 64
Public Const BTN_DELETE = 128
Public Const BTN_CLEAR = 256
Public Const BTN_FIND = 512
Public Const BTN_SEARCH = 1024
Public Const BTN_VIEWMODE = 2048
Public Const BTN_PREVIOUS = 4096
Public Const BTN_NEXT = 8192
Public Const BTN_WORD = 16384
Public Const BTN_EXCEL = 32768
Public Const BTN_HTML = 65536
Public Const BTN_SEARCHFORM = 131072
Public Const BTN_SEARCHTABLE = 262144
Public Const BTN_FILTER = 262144 * 2
Public Const BTN_TOOLS = BTN_FILTER * 2
Public Const BTN_PDF = BTN_TOOLS * 2
Public Const BTN_EXPORT = BTN_PDF * 2
Public Const BTN_ALL = BTN_EXPORT * 2 - 1

'Il nome della ToolBar dell'Anteprima di stampa
Public Const BAND_CLOSE_PREVIEW = "Band_ClosePreview"

'Elenco errori
Public Const ERR_TABLE_STRUCT = vbObjectError + 10000
Public Const ERR_NO_DEFAULT_TABLEVIEW = vbObjectError + 10001
Public Const ERR_NO_PROCESSES = vbObjectError + 10002
Public Const ERR_NDELFILTER = vbObjectError + 2500



'La variabile globale TheApp mantiene un riferimento all'oggetto
'applicazione che viene utilizzato per eseguire le funzionalità
'ed i relativi processi del gestore.
Public TheApp As Application

'La variabile globale gResource mantiene un riferimento all'oggetto
'utilizzato per l'accesso alle risorse stringa, icon e bitmap di Diamante
Public gResource As Resource

Public Cn As DmtOleDbLib.adoConnection
Public Db As DMTDataLayer.Database


Public REGISTRY_KEY As String


'Variabili per il piano dei conti del contratto
Public VarIDEsercizio As Long
Public oPDC As DmtPDC.PDCServices
Public Link_PianoDeiConti As Long
Public Link_ContoPDC As Long
Public Link_ContoPDC_RigaNeg As Long
Public Link_ContoPDC_RigaPos As Long

Public ControlIDOrdine As Control

Public LINK_PARAMETRI_FILIALE As Long

Public Link_TipoScarto As Long
Public Link_TipoCaloPeso As Long
Public Link_TipoAumentoPeso As Long

'ALTRI PARAMETRI ORDINE
Public PAR_NonCalcImpDaAssVeloce As Long
Public PAR_CalcImpAConfOrd As Long
Public PAR_NonVisMsgImpZeroConfOrd As Long
Public PAR_AssNewPedDaAssSingola As Long
Public CONFERMA_ALTRI_PARAMETRI  As Boolean
Public PAR_NonCalcPrezzoDueRefArtOrd As Long

Public PAR_VIS_NOTE_RIGA_ORD_ELENCO As Long
Public PAR_VIS_NOTE_RIGA_ORD As Long


'ALTRI PARAMETRI IV GAMMA
Public ATT_MULT_SEL_IV_GAMMA  As Long
Public CHIUDI_CONF_QTASEL_ZERO As Long
Public FOCUS_IVGAMMA_LAV As Long
Public FOCUS_IVGAMMA_CONF As Long
Public LINK_UM_COOP_SEL_AUT_IVGAMMA As Long

'PARAMETRI CONTRATTO
Public GG_DATA_SCADENZA_CONTR As Long

Public DATA_SCADENZA_OBBL As Long

'PARMETRI DI VENDITA
Public PAR_IDARTICOLO_BOLLO As Long
Public PAR_IDSOCIO_RISC_PESO As Long

Public CONFERMA_TRACC_ONLINE As Long

'PARAMETRI MIGROS
Public GtinMigros As String
Public UrlMigros As String
Public NomeUtenteMigros As String
Public PasswordMigros As String

'PARAMETRI FEEDENTITY
Public ChiaveFeed As String
Public UrlFeed As String
Public AttivaFeed As Long
Public ATTIVA_SUB_CONTATTI_FEED As Long
Public ATTIVA_SEL_MULT_LOTTI_FEED As Long


'PARAMETRI PRODUZIONE
Public ATTIVA_PROD_CONF_LAV As Long
Public ATTIVA_PROD_ORD_LAV As Long
Public DISATTIVA_PROD_ORD_TESTA As Long
Public ATTIVA_PROD_CALC_COLLI_PED As Long
Public ATTIVA_PROD_FIFO_ACCESSI_CONF As Long

Public ATTIVA_SEL_SUB_LOTTO_PROD As Long

Public CodiceAnnullaOperazione As String
Public CodiceConfermaOrdine As String
Public CodiceGestioneErrori As String
Public CodiceConfermaViaggio As String

Public ATTIVA_GESTIONE_CARICO_MERCE As Long
Public NonVisMsgCreaListaPrelAutPedUscita As Long
Public ATTIVA_VisAndGrigliaRigaOdine As Long
Public ATTIVA_VisAndGrigliaConferimento As Long
Public ATTIVA_VisAndGrigliaOrdine As Long
Public IDTipoCollegRigaOrdRigaConf As Long
Public ATTIVA_RICALCOLO_COLLI_PREL_CHIUSO As Long

Public NonVisRigheOrdiniComplInSelLav As Long
Public ModificaLavorazioneSenzaRicalcolo As Long
Public RISCONTRO_PESO_PER_CONF As Long
Public ATTIVA_ELIMINAZIONE_ACCESSI As Long

Public IDSOCIO_PRE_CONF As Long

Public ATTIVA_COMMISSIONI_ORDINI As Long
Public RIC_COMM_TIPO_PED_EVAS_ORD As Long
Public VIS_ELENCO_RIGHE_ORD As Long
Public ATTIVA_CALCOLO_N_PED_BI_VENDITE As Long
Public RISCONTRO_PESO_VAL_SOCIO As Long
Public RISCONTRO_PESO_VAL_FORNITORE As Long
Public OBBL_N_DOC_PES_CONF As Long

Public ATTIVA_OBBL_SFALCIO_CONF As Long
Public ATTIVA_SEQ_SFALCIO As Long

Public RIPORTA_RIF_DETTAGLIO_XML_NC As Long
Public RIPORTA_RIF_DETTAGLIO_XML_ND As Long
Public CONF_AUT_CONTR_PRESA_VISIONE As Long

Public NO_RIP_AGENTE_IN_DOC_EVASIONE As Long

Public Rip_InXMLRifLetteraIntento As Long
Public Rip_InXMLRifNoteIva As Long
Public Rip_InXMLRifNota01Doc As Long
Public Rip_InXMLRifNota02Doc As Long
Public Rip_InXMLRifNota03Doc As Long
Public Rip_InXMLRifNotaDoc As Long
Public Rip_InXMLRifIstrMitt As Long
Public Rip_InXMLRifVettSucc As Long
Public Rip_InXMLRifAgenziaTrasp As Long
Public Rip_InXMLRifTargaAutoMezzo As Long
Public NonRiportaInXMLRifVsNumOrd As Long

Public CONFERMA_PARAMETRI_XML As Boolean

''''ALYANTE''''''''''''''''''''''''''''''''''
Public CONFERMA_PARAMETRI_ALY As Boolean
Public NOME_SERVER_ALY As String
Public NOME_DB_ALY As String
Public USER_PROP_SERVER As String
Public PWD_USER_PROP As String
Public GRUPPO_ALY As String
Public USER_ALY As String
Public PWD_USER_ALY As String

Public COMPANY_CODE As Long
Public ATTIVA_FIDO_ALY As Long
Public DISATTIVA_DDT_FIDO_ALY As Long
Public DISATTIVA_FD_FIDO_ALY As Long
Public DISATTIVA_FA_FIDO_ALY As Long
Public DISATTIVA_NC_FIDO_ALY As Long
Public DISATTIVA_ND_FIDO_ALY As Long

Public DISATTIVA_SCALATA_COMM_TRASP As Long
Public ATTIVA_MULTILINE_ARTICOLO As Long
Public ATTIVA_FEED_MULTI_LIVELLO As Long
Public IDFEED_AZIENDA As String

'PARAMETRI QUALITATIVI
Public QUAL01 As Double
Public QUAL02 As Double
Public QUAL03 As Double
Public QUAL04 As Double
Public QUAL05 As Double
Public QUAL06 As Double
Public QUAL07 As Double
Public QUAL08 As Double
Public QUAL09 As Double
Public QUAL10 As Double
Public QUAL11 As Double
Public QUAL12 As Double
Public QUAL13 As Double
Public QUAL14 As Double
Public QUAL15 As Double
Public QUAL16 As Double
Public QUALPRZ16 As Double
Public CONFERMA_PARAMETRI_QUAL As Boolean
Public NUMERO_COLLI_PRED_CERT As Long

Public SincronizzaAutFeed As Long
Public IvaArticoloDaDocColl As Long
Public LetteraIntentoDaDocColl As Long
Public RifComuneDaConfigSocio As Long
Public DocAnaDestUgualeAnaCoop As Long

Public IDClassLottoProdPerFuoriQuota As Long
Public MsgInDocSeRigaMerceSenzaImballo As Long

Public CarettareClassLottoProd As String
Public CarettareSeparazioneClassLottoProd As String
Public IDAnagraficaDestinazionePerCertificato As Long

Public IDCategoriaAnagraficaSocioDiretto As Long
Public IDCategoriaAnagraficaProdAcq As Long
Public IDCategoriaAnagraficaNoProd As Long
Public IDArticoloScartoPerCertificato As Long

Public CodiceCampoIDFeedPerClass01 As String
Public CodiceCampoIDFeedPerClass02 As String
Public CodiceCampoIDFeedPerAcquisto As String

Public RiportaDestinazioneDaContrattoCertificato As Long
Public RiportaVettoreDaContrattoCertificato As Long
Public ForzaDestinazioneDaContrattoCertificato As Long
Public ForzaVettoreDaContrattoCertificato As Long

Public NonInviareCodiceForInFeedentity As Long
Public PrendiVarietaDaTipologiaFeedentity As Long

Public AttivaSelezioneSocioCertPerVarieta As Long
Public AttivaSelezioneAnaVeloceInCert As Long

Public AttivaPaginazioneContattiFeed As Long
Public NumeroElementiContattiPerPagina As Long

Public AttivaRicercaFatturaAccontoBIVendite As Long
Public AttivaRicercaFatturaAccontoBIFatturato As Long
Public NumeroMesiPerDataRevocaCertificato As Long

Public DataUltimaSincronizzazioneContattiFeed As String
Public DataUltimaSincronizzazioneLottoFeed As String
Public AttivaGestioneUltimaSincronizzazioneFeed As Long
Public NonRiportareRifCerticatoInDDT As Long

Public AttivaControlloEsistenzaLottiInFeedAutInSync As Long
Public NonEliminareLottiDefinitivamenteDaFeed As Long
Public NonEliminareLottiProvvDefinitivamenteDaFeed As Long
Public NonAggiornareRifLottoInFeed As Long

Public DBNameRegFatture As String
Public AttivaMappaturaDaRegFatture As Long


