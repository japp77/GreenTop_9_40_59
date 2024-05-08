Attribute VB_Name = "Globali"
Option Explicit

Public Const HandCursor = 32649&
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long



'API di uso comune.
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Const HWND_TOP As Long = 0
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1



Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER


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

Public LINK_DOCUMENTO_TMP_LIQ As Long

'**********************VARIABILI GLOBALI AZIENDA**************************
    Public VarIDAzienda As Long
    Public VarIDAttivitaAzienda As Long
    Public VarIDFiliale As Long
    Public VarIDEsercizio As Long
    Public VarIDUtente As Long
    
'*************************************************************************

Public VarPassword As String
Public VarUtente As String

Public Link_TipoSocio As Long
'1 = Su
'2 = Giu
Public DeviceSpostamento As Integer

Public LINK_CLIENTE_ORD_PRED As Long
Public DATA_ORDINE_PRED As String
Public NUMERO_ORDINE_PRED As Long

Public LINK_CLIENTE_ORD_SMIST As Long
Public DATA_ORDINE_SMIST As String
Public NUMERO_ORDINE_SMIST As Long


Public WHERE_TROVA_PEDANA As Integer


Public LINK_ORDINE As Long
Public LINK_ORDINE_PRED As Long

Public LINK_UM_LIQ As Long
Public LINK_UM_COOP As Long

Public QUANTITA_PER_COLLI As Double
Public MOLTIPLICATORE As Double

Public LINK_ASSEGNAZIONE_MERCE_PER_SMISTAMENTO As Long
Public LINK_ORDINE_MERCE_PER_SMISTAMENTO As Long
Public LINK_CLIENTE_ORDINE_MERCE_PER_SMISTAMENTO As Long
Public NUMERO_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO As Long
Public DATA_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO As String

Public LINK_TIPO_OGGETTO_LAVORAZIONE  As Long
Public LINK_TIPO_OGGETTO_QUADRATURA As Long
Public LINK_RIGA_CONFERIMENTO_DA_MOV As Long
