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
Public Declare Function GetProfileString Lib "kernel32" _
    Alias "GetProfileStringA" _
    (ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long) As Long
 
Public Declare Function WriteProfileString Lib "kernel32" _
    Alias "WriteProfileStringA" _
    (ByVal lpszSection As String, _
    ByVal lpszKeyName As String, _
    ByVal lpszString As String) As Long

Public Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

''''CURSORE MANINA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const HandCursor = 32649&
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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

Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A


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

Public Flag_GestioneArticoli As Boolean
Public Flag_AutomazioneLotti As Boolean
Public Link_Arrontondamento As Long
Public Numero_Decimali_Pesi As Long

'Variabili che identificano un tipo prodotto di scarti e cali pesi
Public Link_TipoScarto As Long
Public Link_TipoCaloPeso As Long
Public Link_TipoAumentoPeso As Long
Public Link_TipoQuadratura As Long
Public Link_TipoProdotto As Long
Public Link_TipoProdotto_Q As Long

Public Link_Causale_MagCar_Conf As Long
Public Link_Causale_MagScar_Conf As Long
Public Link_Causale_MagCar_Vend As Long
Public Link_Causale_MagScar_Vend As Long
Public Link_UMAcq As Long
Public NumeroDocumentoDisponibile As Long
Public Link_Oggetto As Long
Public DataInizio_Esercizio As String
Public DataFine_Esercizio As String

Public Link_LottoArticolo As Long
Public Link_Movimento_Carico As Long
Public Link_Movimento_Scarico As Long
Public Link_UnitaDiMisura_Acquisto As Long
Public Link_UnitaDiMisura_Coop As Long
Public Link_UnitaDiMisura_Coop_Conferimento As Long
Public Link_UnitaDiMisura_Coop_Q As Long

Public SegnoMovimentoScarto As String


Public Link_Magazzino_Carico As Long
Public Link_Magazzino_Scarico As Long
Public Link_CausaleCarico As Long
Public Link_CausaleScarico As Long



Public VAR_QtaMinimaPerConferimento As Double
Public VAR_QtaMinimaPerVendita As Double

Public Const IDDocumento As Long = 10
Public Link_TipoLavorazione As Long
Public Link_ArticoloPadre As Long
Public Bool_EsistenzaArticoloDerivato As Boolean

Public Link_Numero_Riga_Interno As Long
Public LINK_CONFERIMENTO_RIGA As Long

Public CodiceSocio As String


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
Public CODICE_WEB_AZIENDA As String

Public Controllo_Quad_Conf As Boolean

Public LINK_DOCUMENTO_TMP_LIQ As Long

Public LINK_ORDINE_RIF As Long
Public LINK_PEDANA As Long

Public LINK_CLIENTE_ORD_PRED As Long
Public DATA_ORDINE_PRED As String
Public NUMERO_ORDINE_PRED As Long
Public WHERE_TROVA_PEDANA As Integer
Public LINK_ASSEGNAZIONE As Long

Public LINK_UM_LIQUIDAZIONE As Long
Public QUANTITA_PER_COLLO As Double
Public MOLTIPLICATORE As Double
Public NOME_COMPUTER As String


Public LINK_VARIETA_LOTTO_CAMPAGNA As Long
Public LINK_FAMIGLIA_LOTTO_CAMPAGNA As Long
Public LINK_TIPO_PRODUZIONE_LOTTO_CAMPAGNA As Long

Public VARIETA_LOTTO_CAMPAGNA As String
Public FAMIGLIA_LOTTO_CAMPAGNA As String
Public TIPO_PRODUZIONE_LOTTO_CAMPAGNA As String
Public DATA_SBLOCCO_LOTTO_CAMPAGNA As String


Public ENTE_CERTIFICAZIONE_LOTTO_ETI As String
Public CODICE_CERTIFICAZIONE_LOTTO_ETI As String
Public DESCRIZIONE_CERTIFICAZIONE_LOTTO_ETI As String
Public PROTOCOLLO_CERTIFICAZIONE_LOTTO_ETI As String

Public ENTE_CERTIFICAZIONE_SOCIO_ETI As String
Public CODICE_CERTIFICAZIONE_SOCIO_ETI As String
Public DESCRIZIONE_CERTIFICAZIONE_SOCIO_ETI As String
Public PROTOCOLLO_CERTIFICAZIONE_SOCIO_ETI As String

Public ATTIVAZIONE_NUOVO_CALCOLO As Boolean
Public LINK_ORDINAMENTO As Long



Public AZIENDA_DENOMINAZIONE As String
Public AZIENDA_INDIRIZZO As String
Public AZIENDA_COMUNE As String
Public AZIENDA_REGIONE As String
Public AZIENDA_NAZIONE As String
Public AZIENDA_CAP As String
Public AZIENDA_PROVINCIA As String
Public AZIENDA_TELEFONO As String
Public AZIENDA_FAX As String
Public AZIENDA_EMAIL As String
Public AZIENDA_INDIRIZZOWEB As String
Public AZIENDA_PARTITA_IVA As String

Public FILIALE_DENOMINAZIONE As String
Public FILIALE_INDIRIZZO As String
Public FILIALE_COMUNE As String
Public FILIALE_REGIONE As String
Public FILIALE_NAZIONE As String
Public FILIALE_CAP As String
Public FILIALE_PROVINCIA As String


Public CLIENTE_NOME As String
Public CLIENTE_INDIRIZZO As String
Public CLIENTE_COMUNE As String
Public CLIENTE_REGIONE As String
Public CLIENTE_NAZIONE As String
Public CLIENTE_CAP As String
Public CLIENTE_PROVINCIA As String
Public CLIENTE_TELEFONO As String
Public CLIENTE_FAX As String
Public CLIENTE_EMAIL As String
Public CLIENTE_INDIRIZZOWEB As String
Public CLIENTE_PARTITA_IVA As String

Public SOCIO_INDIRIZZO As String
Public SOCIO_COMUNE As String
Public SOCIO_REGIONE As String
Public SOCIO_NAZIONE As String
Public SOCIO_CAP As String
Public SOCIO_PROVINCIA As String
Public SOCIO_TELEFONO As String
Public SOCIO_FAX As String
Public SOCIO_EMAIL As String
Public SOCIO_INDIRIZZOWEB As String
Public SOCIO_PARTITA_IVA As String

Public DATA_PARTENZA_ORDINE As String

Public DESTINAZIONE_MERCE As String
Public DESTINAZIONE_INDIRIZZO As String
Public DESTINAZIONE_COMUNE As String
Public DESTINAZIONE_PROVINCIA As String
Public DESTINAZIONE_CAP As String
Public DESTINAZIONE_TELEFONO As String
Public DESTINAZIONE_FAX As String
Public DESTINAZIONE_NAZIONE As String

Public VETTORE As String
Public VETTORE_INDIRIZZO As String
Public VETTORE_COMUNE As String
Public VETTORE_PROVINCIA As String
Public VETTORE_CAP As String
Public VETTORE_TELEFONO As String
Public VETTORE_FAX As String
Public VETTORE_PARTITA_IVA As String
Public VETTORE_NUMERO_ALBO As String
Public VETTORE_NAZIONE As String

Public CODICE_CERTIFICAZIONE_AZIENDA As String
Public DESCRIZIONE_CERTIFICAZIONE_AZIENDA As String
Public PROTOCOLLO_CERTIFICAZIONE_AZIENDA As String
Public ENTE_CERTIFICAZIONE_AZIENDA As String


Public CODICE_CERTIFICAZIONE_FAM_LAV As String
Public DESCRIZIONE_CERTIFICAZIONE_FAM_LAV As String
Public PROTOCOLLO_CERTIFICAZIONE_FAM_LAV As String
Public ENTE_CERTIFICAZIONE_FAM_LAV As String

Public CODICE_CERTIFICAZIONE_FAM_CONF As String
Public DESCRIZIONE_CERTIFICAZIONE_FAM_CONF As String
Public PROTOCOLLO_CERTIFICAZIONE_FAM_CONF As String
Public ENTE_CERTIFICAZIONE_FAM_CONF As String


Public VARIAZIONE_DA_PESATURA As Boolean

Public VARIETA_ARTICOLO_CONFERITO As String
Public FAMIGLIA_ARTICOLO_CONFERITO As String
Public VARIETA_ARTICOLO_LAVORATO As String
Public FAMIGLIA_ARTICOLO_LAVORATO As String
Public TIPO_PRODOTTO_ARTICOLO_LAVORATO As String
Public CATEGORIA_MERCEOLOGICA_ARTICOLO_LAVORATO As String

Public PESO_LORDO_ARTICOLO As Double
Public TIPO_COMPORTAMENTO_LAVORAZIONE As Long
Public TIPO_PESO_ARTICOLO As Long
Public STATO_ORDINE As Long
'0 Aperto o non assegnato
'1 Confermato
'2 Chiuso
Public APERTURA_FORM_CODA As Boolean

Public Link_RigaConferimento As Long

Public LINK_TIPO_FIDO_CLIENTE As Long
Public PASSWORD_FIDO_CLIENTE As String
Public DATA_SBLOCCO_FIDO_CLIENTE As String
Public LINK_TIPO_FIDO_AZIENDA As Long
Public PASSWORD_FIDO_AZIENDA As String
Public DATA_SBLOCCO_FIDO_AZIENDA As String

Public LINK_BLOCCO_CLIENTE As Long
Public AVVIA_FIDO_DOPO_CONTROLLO As Boolean

Public LEGGI_DATI_ORDINE As Boolean

Public B_LOADING_RIGA As Boolean

'''''''VARIABILI PER LAVORAZIONE AUTOMATICA''''''''''''''''''''''''
Public LAVORAZIONE_AUTOMATICA As Long
Public PEDANA_AUTOMATICA As Long
Public ORDINE_AUTOMATICO As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''VARIABILI CHE SONO RICHIAMATE DA ALTRI GESTORI''''''''''''''
Public LINK_LAVORAZIONE_DA_MOV As Long
Public TIPO_LAVORAZIONE_DA_MOV As Long

Public ELABORAZIONE_DA_RICHIAMO As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public VISUALIZZA_ANDAMENTO_ORDINE As Long
Public LINK_ORDINE_SELEZIONATO As Long
Public LINK_ART_ORD_PADRE_SEL As Long

Public LINK_PROCESSO_RIGA As Long

Public PREZZI_ARTICOLI_DA_ORDINE As Long
Public PREZZI_IMBALLI_DA_ORDINE As Long
Public PREZZO_INCLUSO_IMBALLO_DA_ORDINE As Long

Public LINK_LISTINO_AZIENDA As Long
Public LINK_LISTINO_IMBALLI_AZIENDA As Long
Public ETICHETTE_PRO_PER_UTENTE_DMT_LAV As Long
Public ETICHETTE_PRO_PER_UTENTE_DMT_PED As Long


'''''''VARIABILI PER PROTOCOLLATURA''''''''''''''''''''''''''
Public Link_Periodo_IVA_Protocollo As Long
Public Link_Sezionale_Dettagliato As Long
Public Numero_Protocollo_Dettagliato As Long
'Public Link_Periodo_IVA_Protocollo As Long
Public Link_Tipo_Passaporto_Azienda As Long
Public Link_TipoPassaportoArticolo As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Control_Return As Control
Public Control_Return_Cliente As Control
Public Control_Return_Data_Ordine As Control

Public AVVIO_RICERCA_ORDINE As Boolean
Public FORZA_LINK_ORDINE As Boolean

Public LINK_TIPO_LIQ_CONF As Long
Public LINK_STATO_LIQ_NUOVO As Long
Public LINK_STATO_LIQ_CHIUSO As Long

Public PREZZO_MEDIO_AUT As Long
Public AGGIORNA_PREZZO_MEDIO As Long
Public AGGIORNA_TIPO_LAVORAZIONE As Long

Public msg_CONFERIMENTO_NEGATIVO As Long

Public LINK_LAVORAZIONE_RIEP As Long

Public MODULO_ATTIVATO As Long
Public MODULO_DESCRIZIONE As String
Public Const MODULO_CODICE As String = "GT002"

Public rsLottoImballo As ADODB.Recordset
Public rsLottoImballoSel As ADODB.Recordset
Public rslottoImballoDelete As ADODB.Recordset

Public LINK_LOTTO_IMBALLO_PRED As Long
Public CONFERMA_LOTTO_IMBALLO_DA_UTENTE As Long
Public LINK_LOTTO_IMBALLO As Long


Public LOADING_ELIMINAZIONE_DOCUMENTO As Boolean

Public GESTIONE_ORDINE_VIVAIO As Long

Public RIGA_ORDINE_SELEZIONATA As Long

Public FORM_PESI_MEDI_SHOW As Boolean

Public VISUALIZZA_IMPORTO_F4 As Long

Public NUOVO_ORDINE_LISTA_PRELIEVO As Boolean

Public LINK_ORDINE_PADRE_SEL As Long

Public TROVA_PREZZI_ORD_CAL As Long
Public TROVA_PREZZI_ORD_CAT As Long
Public NUOVA_RIGA_INIZIO As Long
Public TROVA_PREZZI_NO_IMB As Long

Public B_LOADING_LOTTO_IMBALLO As Boolean

Public LINK_ARTICOLO_ORDINE As Long 'Link articolo della lavorazione a cui biosogna collegare una riga di ordine
Public CONFERMA_SEL_PREZZO_DA_ORD As Long 'Indica se dal form della lista degli articoli dell'ordine è stata confermata la selezione
Public RETURN_SEL_PREZZO_IMB_DA_ORD As Long 'Indica se nella funzione del recupero prezzo da ordine è stato recuperato anche il prezzo dell'imballo
Public RECUPERO_PREZZI_DA_ORD As Boolean 'Indica che è stato recuperato solamente il prezzo dall'ordine e quindi non deve scatenarsi il change dell'IDRigaOrdine
Public LINK_ORDINE_PER_PREZZO As Long 'Riferimento dell'ordine
Public MODALITA_RECUPERO_RIGA_ORD As Long 'Determina il comportamento della form CorpoOrdine (1=Recupera tutte le righe dell'ordine - 0=Recupera solamente le righe per articolo merce)

Public VISUALIZZA_MERCE_ORDINATA As Long

Public rsDistintaGriglia As ADODB.Recordset
Public rsDistintaReturn As ADODB.Recordset

Public REGIONE_PROVENIENZA As String
Public NAZIONE_PROVENIENZA As String
Public COMUNE_PROVENIENZA As String
Public PROVINCIA_PROVENIENZA As String

Public CODICE_ASSOCIATO_PRED As String
Public BNDOO As String

Public rsLottoImballoPrim As ADODB.Recordset
Public rsLottoImballoPrimSel As ADODB.Recordset
Public rslottoImballoPrimDelete As ADODB.Recordset
Public CONFERMA_LOTTO_IMBALLO_DA_UTENTE_PRIM As Long
Public LINK_LOTTO_IMBALLO_PRIM As Long
Public B_LOADING_LOTTO_IMBALLO_PRIM As Boolean

Public rsKIT As ADODB.Recordset
Public rsKITLotti As ADODB.Recordset


Public IDIMBALLOPRIM_SEL As Long
Public IDIMBALLO_SEL As Long

Public COPIA_RIGA_COME_NUOVO As Long
Public RIP_DIFF_COLLI_DA_ORD As Long

''ORDINE CLIENTE
Public ORDINE_TIPO_ORDINE As String
Public ORDINE_TARGA_AUTOMEZZO As String
Public ORDINE_ISTRUZIONI_MITT As String
Public ORDINE_NOTE_FATT As String
Public ORDINE_NOTE_CORPO As String
Public ORDINE_NOTE_INTERNE As String

'IMBALLO SECONDARIO
Public DESCR_IMB_LINGUA_PRED As String
'IMBALLO PRIMARIO
Public DESCR_IMB_PRIM_LINGUA_CLIENTE As String
Public DESCR_IMB_PRIM_LINGUA_PRED As String
Public COD_A_BARRE_IMB_PRIM_CLI As String
Public DESCR_A_BARRE_IMB_PRIM_CLI As String
Public COD_A_BARRE_IMB_PRIM_PRED As String
Public DESCR_A_BARRE_IMB_PRIM_PRED As String


Public RIPORTA_RIGA_DA_ORDINE As Long
Public VIS_ELENCO_ORD_NEW_PED As Long

Public NON_PROP_ULT_ART_PED As Long


Public PERCORSO_ETICHETTE_PDF As String
Public STAMPA_ETICHETTE_PDF As Long

Public VIS_NOTE_LAV_DA_ORD As Long
Public IDProcessoProd As Long
Public IDRigaProcessoProd As Long
Public IDLineaProduzione As Long
Public LINK_PROCESSO_RIGHE_PROD As Long
Public LINK_PROCESSO_PROD As Long
Public LINK_LINEA_PROD As Long
Public PESO_TOTALE_PEDANA_DA_PROD As Double

Public ATTIVA_SEL_LOTTO_PROD_IN_LAV As Long
Public LINK_LOTTO_PROD_LAV As Long
Public LINK_SOCIO_LOTTO_PROD_LAV As Long

Public NON_VIS_RIGHE_ORD_COMPLETE As Long
Public NON_RICALCOLARE_ALLE_MODIFICHE As Long

Public IDSOCIO_PRE_CONF As Long

Public QUANTITA_KIT_SEL As Double
Public LINK_KIT_SEL As Long
Public ARTICOLO_KIT_SEL As String
Public BLOCCO_QTA_LAV As Long

