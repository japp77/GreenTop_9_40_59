Attribute VB_Name = "Globali"
Option Explicit

'Declares
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Sub sbOpenURL Lib "Diamante.dll" (ByVal hwnd As Long, ByVal sURL As String)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
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

Public Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hwnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long



Public Const SW_SHOWNORMAL = 1
Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&
Public Const CSIDL_COMMON_APPDATA = &H1C '&H23
Public Const MAX_PATH = 260

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
Public Link_TipoImballo As Long
Public Link_Articolo As Long

Public NumeroRiga As Long
Public NumeroProgSingolaRiga As Long
Public Link_UMCoop As Long

Public NumeroColliRiga As Long
Public NuovoDocumento As Integer

Public Val_ColliCaricati As Long
Public Val_ColliVenduti  As Long
Public Val_ColliVendutiOriginali As Long

Public Val_ColliVendutiTotali As Long
Public Val_ColliColcolati As Long
Public QtaDBLotto As Double



Public Var_LostFocus_Colli As Integer
Public Val_PesoLordo As Double
Public Val_Pezzi As Double
'PROTOCOLLO ICE

Public Var_Modalita_Ricerca As Integer
Public Var_FermaRicerca As Integer
Public Var_RiportoRiga_Da_Nuova_Lavorazione As Integer

Public Link_Protocollo_ICE As Long
Public NumeroProtocolloICE_DB As Long
Public Link_SchemaProtocollo As Long
Public OLD_STATO_PROTOCOLLO As Integer
Public Link_IDOggetto_OLD As Long

Public sSQL_Ricerca As String
Public AggiornaRiga As Integer

'Variabili Intrastat articolo
Public MassaNetta_Art As Double
Public Link_Nom_Comb_Art As Long
Public Link_Nat_Trans_Art As Long
'Variabili Intrastat imballo
Public MassaNetta_Imb As Double
Public Link_Nom_Comb_Imb As Long
Public Link_Nat_Trans_Imb As Long

Public Link_ArticoloPadre As Long

Public LINK_CLIENTE_IVA As Long

Public Link_TipoSocio As Long
Public Link_TipoLavorato As Long
Public Link_TipoGrezzo As Long
Public GestioneArticoli As Boolean

Public Link_Socio As Long
Public Link_RigaConferimento As Long
Public Link_RigaAssegnazioneMerce As Long
Public Link_RigaProcessoIVGamma As Long
Public Link_TipoUtilizzoProcesso As Long

Public Par_OBBLIGATORIO As Integer
Public Link_Oggetto As Long

Public Elaborazione_da_wizard As Boolean

Public Quantita_da_accreditare As Double
Public Colli_da_accreditare As Double
Public PesoLordo_da_accreditare As Double
Public Tara_da_accreditare As Double
Public PesoNetto_da_accreditare As Double
Public Pezzi_da_accreditare As Double

Public Importo_da_accreditare As Double

Public Quantita_Totale_Selezionata As Double
Public Colli_Totali_Selezionati As Double
Public PesoLordo_Totali_Selezionati As Double
Public Tara_Totali_Selezionati As Double
Public PesoNetto_Totali_Selezionati As Double
Public Pezzi_Totali_Selezionati As Double

Public Rif_UltimoIDOggetto As Long
Public Rif_UltimoTipoOggetto As Long
Public Rif_UltimoNumeroDoc As String
Public Rif_UltimoDataDoc As String
Public Rif_UltimoPrefissoDoc As String
Public Rif_LetteraIntento As Long

Public LINK_UM_LIQ As Long
Public QUANTITA_PER_COLLI As Double
Public Moltiplicatore As Double

Public TOTALE_MERCE As Double

Public ATTIVAZIONE_NUOVO_CALCOLO As Double

'Variabili che identificano un tipo prodotto di scarti e cali pesi
Public Link_TipoScarto As Long
Public Link_TipoCaloPeso As Long
Public Link_TipoAumentoPeso As Long

'Oggetto utilizzato per gestire l'inserimento / variazione del documento (DmtDocs.Dll)
Public oDoc As DmtDocs.cDocument
'Variabile utilizzata per ottenere il nome della tabella di testata del documento
Public sTabellaTestata As String
'Variabile utilizzata per ottenere il nome della tabella di dettaglio del documento
Public sTabellaDettaglio As String
'Variabile utilizzata per ottenere il nome della tabella delle scadenze del documento
Public sTabellaScadenze As String
'Variabile utilizzata per ottenere il nome della tabella del castelletto IVA del documento
Public sTabellaIVA As String


Public LINK_TIPO_FIDO_CLIENTE As Long
Public PASSWORD_FIDO_CLIENTE As String
Public DATA_SBLOCCO_FIDO_CLIENTE As String
Public LINK_TIPO_FIDO_AZIENDA As Long
Public PASSWORD_FIDO_AZIENDA As String
Public DATA_SBLOCCO_FIDO_AZIENDA As String

Public LINK_BLOCCO_CLIENTE As Long
Public AVVIA_FIDO_DOPO_CONTROLLO As Boolean
Public Numero_Decimali_Pesi As Long


Public oPDC As DmtPDC.PDCServices
Public Link_PianoDeiConti As Long


Public LINK_TIPO_LIQ_CONF As Long
Public PREZZO_MEDIO_AUT As Long
Public AGGIORNA_PREZZO_MEDIO As Long
Public AGGIORNA_TIPO_LAVORAZIONE As Long

Public MODULO_ATTIVATO As Long
Public MODULO_DESCRIZIONE As String
Public Const MODULO_CODICE As String = "GT001"
Public TIPO_VARIAZIONE_DA_WIZARD As Long

Public rsGrigliaSelDoc As ADODB.Recordset
Public rsGrigliaSelDocTMP As ADODB.Recordset
Public rsGrigliaReturn As ADODB.Recordset
Public StampaDocSel As Boolean


Public NUMERO_COPIE_SEL_DOC As Integer
Public ORIENTAMENTO_SEL_DOC As OrientationConsts

Public COLLEGAMENTO_NOTA_PER_LOTTO As Long

Public RiportaTuttoDocumento As Boolean

Public NON_CALC_COMM As Long
Public NON_CALC_INCIDENZA_IMB As Long

Public Rif_PA_Riga_Doc_Merce As String
Public Rif_PA_Riga_Doc_Imballo As String
Public Tipo_Riga_Sel_EF As Long
Public ID_ART_PROG_MERCE As Long
Public ID_ART_PROG_IMBALLO As Long
Public LINK_LETTERA_INTENTO_RIF_DOC As Long

Public LINK_OGGETTO_COLLEGATO As Long
Public LINK_TIPO_OGGETTO_COLLEGATO As Long
Public LINK_LOTTO_PROD_DA_LAV_MOV As Long

Public DATA_COMPETENZA_LIQ As String
Public ANNOTAZIONE_01 As String
Public ANNOTAZIONE_02 As String
Public ANNOTAZIONE_03 As String
Public CONFERMA_ALTRI_DATI As Long

Public LINK_TIPO_NOTA_SEL As Long
Public CONFERMA_RIGA_NOTA As Long
Public RETURN_RIGA_NOTA As String

Public RIPORTA_RIF_DETTAGLIATI As Long

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
