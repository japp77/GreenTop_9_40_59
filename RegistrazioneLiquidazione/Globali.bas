Attribute VB_Name = "Globali"
Option Explicit

'API di uso comune.
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


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
Public Const NATIVE_LANGUAGE = 1
Public Const URL_DIAMANTE = "http://www.diamante.it"


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

'Oggetto Semaforo usato per gestire i conflitti di multiutenza.
Public gSemaphore As Semaforo.dmtSemaphore



'La variabile globale Application_Name è valorizzata nella Sub Main.
Public Application_Name As String


'La variabile globale Current_Process_ID è valorizzata nella Sub Main
'e rappresenta l'ID del processo in esecuzione.
Public Current_Process_ID As Long

'////////////////////////////////////////////////////////
'Impostare questa costante con il nome del
'processo previsto per questa manutenzione
'////////////////////////////////////////////////////////
Public Const PROCESS_NAME = "Manutenzione"

Public TipoScelta As Integer
'Variabile di connessione
Public CnDMT As DmtOleDbLib.adoConnection


'Variabili filtri per liquidazione
Public DATA_INIZIO As String
Public DATA_FINE As String
Public NUMERO_LIQUIDAZIONE As Long
Public LINK_PERIODO As Long
Public TIPO_IMPORTO_ARTICOLO As Integer
Public LINK_SOCIO As Long
Public LINK_ARTICOLO As Long
Public TIPO_IMPORTO_DOCUMENTO As Long
Public TIPO_IMPORTO_RIGA_DOCUMENTO As Long
Public TIPO_QUANTITA As Long
Public ARTICOLI_DI_QUAD As Boolean
Public TIPO_QUADRATURA As Long
Public TIPO_LIQUIDAZIONE As Long
Public TIPO_CALCOLO_PREZZO_MEDIO As Long
Public ANNOTAZIONI_PERIODO As String
Public NUMERO_PROTOCOLLO As Long
Public TRATTENUTA_PER_IMPORTO As Double
Public TRATTENUTA_PER_PERCENTUALE As Double
Public LINK_LISTINO As Long
Public LINK_CAT_MERCE As Long
Public AGG_COSTO_KIT_PRZ_LIQ As Long
Public AGG_COSTO_CONFEZ_PRZ_LIQ As Long
Public COMMISSIONI_DA_PRZ_VENDITA As Long
Public NON_CALC_COMM As Long
Public NON_CALC_INCIDENZA_IMB As Long
Public NO_LIQ_VEND_UFF As Long
Public ATTIVA_CALCOLO_QTA_DA_ABB As Long
Public RICALCOLA_TUTTO As Long
Public TIPO_RAGGR_ABB As Long
Public DEC_QTA_LIQ As Long
Public DEC_IMP_UNI_LIQ As Long
Public NO_CALC_PRELIQ_ND As Long
Public NO_CALC_PRELIQ_NC As Long


Public LINK_TIPO_LIQ_CONF As Long   'Utilizzare per liquidare il tipo di stato della riga di conferimento
Public CALCOLA_PM_CAMP As Long      'Indica se bisogna calcolare il prezzo medio nella campionatura e aggiornare la riga
Public AGGIORNA_PM_CAMP As Long     'Aggiorna i prezzi medi delle campionature non bloccati ad ogni elaborazione
Public CALCOLA_TRATT_CAMP As Long   'Indica che deve calcolare le trattenute nella campionature
Public LIQ_CONGUAGLIO As Long       'Indica che è una liquidazione di conguaglio
Public LINK_SOCIO_SEL As Long




Public LINK_TIPO_PREZZO_MEDIO_ARTICOLO As Long

Public Nuova_Liquidazione As Integer

Public Controllo_Quad_Conf As Boolean

Public LINK_DOCUMENTO_TMP_LIQ As Long

Public Link_TipoSocio As Long

'**********************VARIABILI GLOBALI AZIENDA**************************
    Public VarIDAttivitaAzienda As Long
    Public VarIDEsercizio As Long
    
'*************************************************************************

Public VarPassword As String
Public VarUtente As String

Public LIQUIDA_FORNITORE As Long
Public RICALCOLA_VALORI_LIQ As Long

Public CONFERMA_SEL_DOCUMENTI As Long

Public MODULO_ATTIVATO As Long
Public MODULO_DESCRIZIONE As String
Public Const MODULO_CODICE As String = "GT003"
Public COLLEGAMENTO_NOTA_PER_LOTTO As Long

Public RISCONTRO_PESO_SOCIO_VAL As Long
Public RISCONTRO_PESO_FORN_VAL As Long
Public LINK_TIPO_CATEGORIA_SOCIO As Long
Public NO_RIP_SCARTI_IN_LIQ As Long

Public LINK_TIPO_AUMENTO_PESO As Long

