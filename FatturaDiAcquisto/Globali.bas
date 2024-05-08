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
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOP = 0
Public Const WM_SETREDRAW = &HB

Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

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


'''''''''''''''''''VARIABILI GLOBALI APPLICAZIONE CONFERIMENTO MERCE''''''''''''''''''''''''''
Public Link_Magazzino_Conferimento As Long
Public Link_Magazzino_Vendita As Long
Public Link_Esercizio As Long
Public Link_PeriodoIVA As Long
Public Link_Sezionale As Long

Public Link_TipoImballo As Long
Public Link_TipoSocio As Long
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


Public Const IDDocumento As Long = 1
Public EsistenzaValoreSezionale As Boolean

Public CodiceSocio As String
Public CambioSocio As Boolean
Public CambioSocioPerLiquidazioni As Boolean
Public Link_Socio_OLD As Long



Public DATA_INIZIO As String
Public DATA_FINE As String
Public NUMERO_LIQUIDAZIONE As Long
Public LINK_PERIODO As Long
Public TIPO_IMPORTO_ARTICOLO As Integer
Public LINK_SOCIO As Long
Public TIPO_IMPORTO_DOCUMENTO As Long
Public TIPO_QUANTITA As Long
Public ARTICOLI_DI_QUAD As Boolean
Public TIPO_QUADRATURA As Long
Public TIPO_LIQUIDAZIONE As Long
Public TIPO_CALCOLO_PREZZO_MEDIO As Long
Public ANNOTAZIONI_PERIODO As String

Public Controllo_Quad_Conf As Boolean

Public LINK_DOCUMENTO_TMP_LIQ As Long

Public S_Annotazioni1 As String
Public S_Annotazioni2 As String
Public S_Annotazioni3 As String

Public LINK_ORDINAMENTO As Long
Public ATTIVAZIONE_NUOVO_CALCOLO As Boolean

Public LINK_REGIONE_SOCIO As Long
Public LINK_NAZIONE_SOCIO As Long
Public LINK_COMUNE_SOCIO As Long
Public LINK_PROVINCIA_SOCIO As Long


'VARIABILI LOTTO DI CAMPAGNA'''''''''''''''''''''''''''''
Public LINK_VARIETA_LOTTO_CAMPAGNA As Long
Public LINK_FAMIGLIA_LOTTO_CAMPAGNA As Long
Public LINK_TIPO_PRODUZIONE_LOTTO_CAMPAGNA As Long
Public DATA_SBLOCCO_LOTTO_CAMPAGNA As String

Public VARIETA_LOTTO_CAMPAGNA As String
Public FAMIGLIA_LOTTO_CAMPAGNA As String
Public TIPO_PRODUZIONE_LOTTO_CAMPAGNA As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''

'VARIABILI DA PARAMETRI''''''''''''''''''''''''''''''
Public LOTTO_CAMPAGNA_OBBLIGATORIO As Long
Public LINK_TIPO_ARTICOLO_CONFERITO As Long

Public LINK_TESTA_DOCUMENTO As Long

Public APERTURA_FORM_CODA As Boolean

Public LINK_TIPO_ARROTONDAMENTO As Long

Public Link_RigaConferimento As Long
Public LINK_ARTICOLO_CONFERITO As Long

''''''VARIABILI DA ALTRI GESTORI CHE RICHIAMANO LA PROCEDURA''''''''''''''''''''''''''''''
Public LINK_RIGA_CONFERIMENTO_DA_MOV As Long
Public ELABORAZIONE_DA_MOV As Long

'IDENTIFICATIVO DEL TOUR''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public LINK_TOUR As Long
Public LINK_TOUR_RIGHE As Long
Public POSIZIONE_TOUR As Long

Public LINK_OGGETTO_ACQ As Long
Public LINK_FUNZIONE_DDT As Long
Public LINK_FUNZIONE_FA As Long


Public LINK_TIPO_LIQ_CONF As Long
Public LINK_STATO_LIQ_NUOVO As Long
Public LINK_STATO_LIQ_CHIUSO As Long

Public PREZZO_MEDIO_AUT As Long
Public AGGIORNA_PREZZO_MEDIO As Long
Public AGGIORNA_TIPO_LAVORAZIONE As Long


Public MODULO_ATTIVATO As Long
Public MODULO_DESCRIZIONE As String
Public Const MODULO_CODICE As String = "GT001"

Public rsLottoImballo As ADODB.Recordset
Public rsLottoImballoSel As ADODB.Recordset
Public rslottoImballoDelete As ADODB.Recordset

Public LINK_LOTTO_IMBALLO_PRED As Long
Public CONFERMA_LOTTO_IMBALLO_DA_UTENTE As Long
Public LINK_LOTTO_IMBALLO As Long
Public LOADING_ELIMINAZIONE_DOCUMENTO As Boolean

Public CALCOLA_RIEP_IMBALLI_STAMPA As Long

Public Link_Ordinamento_riga_conf As Long
Public ATTIVA_CALCOLO_PESO_LORDO As Long

Public FORM_PESI_MEDI_SHOW As Boolean

Public VISUALIZZA_IMPORTO_F4 As Long

Public QUANTITA_PER_COLLO As Double
Public MOLTIPLICATORE As Double
Public PESO_LORDO_ARTICOLO As Double
Public TIPO_PESO_ARTICOLO As Long
Public ATTIVA_LOTTO_PROD_ANA_FATT As Long

Public CONFERMA_SALVA_PESATURA As Long
Public SALVA_RIGA_OK As Long
Public SALVA_DOC_OK As Long

Public IDSOCIO_PRE_CONF As Long
Public ATTIVA_OBBLIGO_N_DOC_SOCIO As Long
