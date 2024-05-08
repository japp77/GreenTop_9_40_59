Attribute VB_Name = "Globali"
Option Explicit

'API di uso comune.
Public Declare Function fnAnsi2Jet Lib "Diamante.dll" Alias "fnAnsi2jet" (ByVal sSQL As String) As String
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long)
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function PlaySound Lib "winmm" Alias "PlaySoundA" (ByVal szName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


'''''''''''''''''''''''''''''Variabili per il PlaySound''''''''''''''''''''''''''''''
Public Const SND_SYNC = &H0 '(predefinito) Il valore di ritorno di PlaySound verrà restituito al termine dell'esecuzione del file, che non sarà possibile interrompere.
Public Const SND_ASYNC = &H1 'Il file viene eseguito in sottofondo, e il programma proseguirà la sua esecuzione. Passando Null a szName alla prossima chiamata dell'API, si interrompe l'esecuzione.
Public Const SND_ALIAS = &H10000 'szName contiene un alias di un suono usato per un evento di sistema (vedi sotto).
Public Const SND_FILENAME = &H20000

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
Public CnDMT As DmtOleDbLib.AdoConnection


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
Public Moltiplicatore As Double

Public LINK_ASSEGNAZIONE_MERCE_PER_SMISTAMENTO As Long
Public LINK_ORDINE_MERCE_PER_SMISTAMENTO As Long
Public LINK_ORDINE_PADRE_MERCE_PER_SMISTAMENTO As Long
Public LINK_CLIENTE_ORDINE_MERCE_PER_SMISTAMENTO As Long
Public NUMERO_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO As Long
Public DATA_DOCUMENTO_ORDINE_MERCE_PER_SMISTAMENTO As String
Public NUMERO_LISTA_ORDINE_MERCE_PER_SMISTAMENTO As Long

Public COMANDO_SPEZZATUTA As Long '0 verso l'ordine in alto; 1 verso merce in giacenza
Public COMANDO_RIPESATURA As Long

Public LINK_TIPO_FIDO_CLIENTE As Long
Public PASSWORD_FIDO_CLIENTE As String
Public DATA_SBLOCCO_FIDO_CLIENTE As String
Public LINK_TIPO_FIDO_AZIENDA As Long
Public PASSWORD_FIDO_AZIENDA As String
Public DATA_SBLOCCO_FIDO_AZIENDA As String

Public LINK_BLOCCO_CLIENTE As Long
Public AVVIA_FIDO_DOPO_CONTROLLO As Boolean

Public AGGIORNA_GRIGLIA_ORDINI_PER_EVASIONE As Boolean
Public LINK_LUOGO_MERCE_PER_EVASIONE As Long
Public LINK_VETTORE_SUCCESSIVO_PER_EVASIONE  As Long
Public DESCRIZIONE_CORPO_PER_EVASIONE As String

Public BLoading As Boolean

'''''PARAMETRIZZIONE EVASIONE ORDINI'''''''''''''''''''''''
Public ERRORE_EVASIONE As String
Public BLOCCA_ORDINE_PAR As Boolean
Public FLAG_ASSEGNAZIONE_VELOCE As Boolean
Public CODICE_CONFERMA_ORDINE As String
Public RIPRISTINA_PARAMETRI As String
Public CODICE_GESTIONE_ERRORI As String

Public LINK_LISTINO_CLIENTE As Long
Public LINK_LISTINO_CLIENTE_SOTTO As Long
Public LINK_LISTINO_AZIENDA As Long
Public LINK_LISTINO_IMBALLI_AZIENDA As Long


Public FLAG_VIS_TOT_ORD_PREP As Long
Public FLAG_VIS_TOT_ORD_SMIST As Long

Public PREZZI_ARTICOLI_DA_ORDINE As Long
Public PREZZI_IMBALLI_DA_ORDINE As Long
Public PREZZO_INCLUSO_IMBALLO_DA_ORDINE As Long

Public Control_Return As Control
Public Control_Return_Cliente As Control
Public Control_Return_Data_Ordine As Control

Public NUMERO_ZERI_DOC_RIF As Long
Public LINK_CONF_RIGA_PER_SCARTI As Long


Public LINK_TIPO_LIQ_CONF As Long
Public PREZZO_MEDIO_AUT As Long
Public AGGIORNA_PREZZO_MEDIO As Long
Public AGGIORNA_TIPO_LAVORAZIONE As Long

Public rsPedana As ADODB.Recordset
Public RIPORTA_IN_ORDINE_DA_PESATURA As Long
Public NUMERO_PEDANA_PESATURA As Long
Public AVVIO_VELOCE_RIPESATURA As Long
Public ATTIVA_MASCHERE As Long

Public rsLavorazioni As ADODB.Recordset
Public N_EliminaLav As Long


Public PESO_PEDANA_IN_VENDITA As Long

Public MODULO_ATTIVATO As Long
Public MODULO_DESCRIZIONE As String
Public Const MODULO_CODICE As String = "GT002"

Public GESTIONE_ORDINE_VIVAIO As Long
Public LINK_ARTICOLO_COMM_CONF As Long
Public LINK_TIPO_TRATT_AGG_COMM_CONF As Long

Public LINK_PORTO_NO_TRASP As Long
Public TIPO_CALCOLO_TRASPORTO As Long

Public NUOVO_ORDINE_LISTA_PRELIEVO As Boolean
Public LINK_ORDINE_PADRE_SEL As Long
Public STAMPA_DOCUMENTO_ATTIVO As Long
Public STAMPA_DOCUMENTO_NON_ATTIVO As Long


Public TROVA_PREZZI_ORD_CAL As Long
Public TROVA_PREZZI_ORD_CAT As Long
Public TROVA_PREZZI_NO_IMB As Long


Public LINK_ARTICOLO_ORDINE As Long 'Link articolo della lavorazione a cui biosogna collegare una riga di ordine
Public CONFERMA_SEL_PREZZO_DA_ORD As Long 'Indica se dal form della lista degli articoli dell'ordine è stata confermata la selezione
Public RETURN_SEL_PREZZO_IMB_DA_ORD As Long 'Indica se nella funzione del recupero prezzo da ordine è stato recuperato anche il prezzo dell'imballo
Public RECUPERO_PREZZI_DA_ORD As Boolean 'Indica che è stato recuperato solamente il prezzo dall'ordine e quindi non deve scatenarsi il change dell'IDRigaOrdine
Public LINK_ORDINE_PER_PREZZO As Long 'Riferimento dell'ordine
Public MODALITA_RECUPERO_RIGA_ORD As Long 'Determina il comportamento della form CorpoOrdine (1=Recupera tutte le righe dell'ordine - 0=Recupera solamente le righe per articolo merce)
Public RECORDSET_RETURN_PER_PREZZO As ADODB.Recordset 'Recordset che contiene la lavorazione a cui cambiare i prezzi
Public LINK_LAVORAZIONE_PER_PREZZO_ORD As Long 'Riferimento della riga lavorazione da aggiornare nella prezzatura veloce

Public SPACCATURA_MERCE_VERSO 'Indica il verso della spaccatura 1=Da ordine su a ordine giu - 0=Da ordine giu a ordine su

Public AGG_COSTO_KIT_PRZ_LIQ As Long
Public AGG_COSTO_CONFEZ_PRZ_LIQ As Long
Public COMMISSIONI_DA_PRZ_VENDITA As Long

Public USA_PROT_ICE_PERIODO As Long

Public AVVIA_FATTURAZIONE As Long

Public PAR_NonCalcImpDaAssVeloce As Long
Public PAR_CalcImpAConfOrd As Long
Public PAR_NonVisMsgImpZeroConfOrd As Long
Public PAR_AssNewPedDaAssSingola As Long
Public PAR_NonCalcPrezzoDueRefArtOrd As Long

Public ATTIVA_COMMISSIONI_DA_ORDINE As Long
Public RIC_COMM_TIPO_PED_DA_ORD As Long
Public VIS_ELECO_RIGHE_ORD As Long

Public LINK_TIPO_NOTA_SEL As Long
Public CONFERMA_RIGA_NOTA As Long
Public RETURN_RIGA_NOTA As String
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

Public DISATTIVA_SCALATA_COMM_TRASP As Long
Public RIP_IVA_DA_DOC_COLL As Long
Public RIP_LET_INT_DA_DOC_COLL As Long

