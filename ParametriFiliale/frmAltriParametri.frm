VERSION 5.00
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{A83BB158-4E50-11D2-B95E-002018813989}#8.3#0"; "DmtSearchAccount.OCX"
Begin VB.Form frmAltriParametri 
   Caption         =   "Altri parametri"
   ClientHeight    =   13035
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20730
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAltriParametri.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13035
   ScaleWidth      =   20730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check26 
      Caption         =   "Prendi in considerazione le commissioni sia nel tipo pedana che nella configurazione del cliente"
      Height          =   255
      Left            =   240
      TabIndex        =   88
      Top             =   2280
      Width           =   9975
   End
   Begin VB.Frame Frame7 
      Caption         =   "Gestione sfalci"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   15600
      TabIndex        =   81
      Top             =   11640
      Width           =   5055
      Begin VB.CheckBox Check21 
         Caption         =   "Attiva sequenza sfalci nei conferimenti"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   600
         Width           =   4815
      End
      Begin VB.CheckBox Check20 
         Caption         =   "Attiva riferimento obbligatorio nei conferimenti"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Pesatura arrivi merce"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   79
      Top             =   11520
      Width           =   10215
      Begin VB.CheckBox Check19 
         Caption         =   "Rendi obbligatorio il riferimento documento"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Statistiche BI Vendite"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   10440
      TabIndex        =   74
      Top             =   11640
      Width           =   5175
      Begin VB.CheckBox Check38 
         Caption         =   "Attiva ricerca fattura acconto (BI Venduto)"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   600
         Width           =   3975
      End
      Begin VB.CheckBox Check37 
         Caption         =   "Attiva ricerca fattura acconto (BI fatturato)"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   840
         Width           =   4695
      End
      Begin VB.CheckBox Check16 
         Caption         =   "Attiva calcolo numero pedane effettive "
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pre-Conferimento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   69
      Top             =   10680
      Width           =   10215
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   15
         Left            =   0
         TabIndex        =   78
         Top             =   960
         Width           =   10215
      End
      Begin DmtSearchAccount.DmtSearchACS ACSSocioPerConf 
         Height          =   585
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   1032
         WidthDescription=   3800
         WidthSecondDescription=   1500
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SearchType      =   2
         HideLeaf        =   0   'False
         BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionDescription=   "Socio/Fornitore per movimentazione imballi"
         CaptionCode     =   "Codice"
         OnlyAccounts    =   -1  'True
      End
   End
   Begin VB.Frame fraCodiceABarreVeloce 
      Caption         =   "Attivazione codici per evasione ordini veloce - controllo accessi - controllo uscita pedane"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   10440
      TabIndex        =   49
      Top             =   4800
      Width           =   10215
      Begin VB.TextBox txtCodiceConfermaViaggio 
         Height          =   285
         Left            =   7560
         TabIndex        =   56
         ToolTipText     =   "Deve inziare "
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtCodiceConfermaOrdine 
         Height          =   285
         Left            =   120
         TabIndex        =   52
         ToolTipText     =   "Deve iniziare con ""C"""
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtCodiceAnnullaOperazione 
         Height          =   285
         Left            =   2760
         TabIndex        =   51
         ToolTipText     =   "Deve iniziare con ""R"""
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtCodiceGestioneErrori 
         Height          =   285
         Left            =   5160
         TabIndex        =   50
         ToolTipText     =   "Deve iniziare con ""E"""
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Conferma carico della merce"
         Height          =   255
         Index           =   1
         Left            =   7560
         TabIndex        =   57
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "Conferma ordine/operazioni"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label18 
         Caption         =   "Ripristino operazioni"
         Height          =   255
         Left            =   2760
         TabIndex        =   54
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label19 
         Caption         =   "Gestione errori"
         Height          =   255
         Left            =   5160
         TabIndex        =   53
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame fraIVGamma 
      Caption         =   "IV Gamma"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   10215
      Begin DMTDataCmb.DMTCombo cboFocusIVGammaLav 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkChiudiConf 
         Caption         =   "Chiudi conferimento quando la quantità rimasta è uguale o minore di zero (solo con selezione multipla)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   9975
      End
      Begin VB.CheckBox chkAttSelMult 
         Caption         =   "Attiva selezione multipla dei conferimenti o delle lavorazioni nel processo di entrata"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9975
      End
      Begin DMTDataCmb.DMTCombo cboFocusIVGammaConf 
         Height          =   315
         Left            =   3840
         TabIndex        =   14
         Top             =   1200
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboUMCoopSelAut 
         Height          =   315
         Left            =   7560
         TabIndex        =   16
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "U.M. per sel. aut"
         Height          =   255
         Index           =   3
         Left            =   7560
         TabIndex        =   17
         ToolTipText     =   "Unità di misura coop per selezione automatica"
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Focus controllo in selezione conferimenti"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   15
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Focus controllo in selezione lavorazioni"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Linee di produzione "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4815
      Left            =   10440
      TabIndex        =   40
      Top             =   0
      Width           =   10215
      Begin VB.CheckBox Check12 
         Caption         =   "Consenti eliminazione accessi dei conferimenti in lavorazione"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3960
         Width           =   9975
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Alla chiusura del prelievo riporta i colli realmente utilizzati dalla gestione del controllo pedane in entrata"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3600
         Width           =   9975
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Attiva la visualizzazione dell'andamento in lavorazione nella lista delle righe di ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3240
         Width           =   9975
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Attiva la visualizzazione dell'andamento in lavorazione nella lista dei conferimenti"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2880
         Width           =   9975
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Non visualizzare il messaggio di conferma della gestione del carico merce"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2520
         Width           =   9975
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Attiva nel controllo pedane in uscita la gestione del carico merce "
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   2160
         Width           =   9975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Attiva modalità FIFO nella gestione entrata merce del conferimento"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1800
         Width           =   6615
      End
      Begin VB.CheckBox chkAttPesoColloPesoPed 
         Caption         =   "Attiva il calcolo del peso per collo in base al totale peso della pedana"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   6615
      End
      Begin VB.CheckBox chkAttSoloRigheOrd 
         Caption         =   "Attiva solamente la lista delle righe ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   6615
      End
      Begin VB.CheckBox chkAttOrdInLav 
         Caption         =   "Attiva gestione ordini in lavorazione"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   6375
      End
      Begin VB.CheckBox chkAttConfInLav 
         Caption         =   "Attiva gestione conferimenti in lavorazione"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   6375
      End
      Begin DMTDataCmb.DMTCombo cboTipoCollFiltroConf 
         Height          =   315
         Left            =   5880
         TabIndex        =   63
         Top             =   4320
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label17 
         Caption         =   "Quando seleziono una riga di ordine filtra la lista conferimenti per "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   4365
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametri per FEEDENTITY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   10440
      TabIndex        =   37
      Top             =   7320
      Width           =   10215
      Begin VB.CheckBox Check43 
         Caption         =   "Non aggiornare il riferimento del lotto"
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   2640
         Width           =   4575
      End
      Begin VB.CheckBox Check42 
         Caption         =   "Non eliminare mai definitivamente i lotti provvisori"
         Height          =   255
         Left            =   4800
         TabIndex        =   114
         Top             =   2880
         Width           =   5175
      End
      Begin VB.CheckBox Check41 
         Caption         =   "Non eliminare mai definitivamente i lotti di produzione"
         Height          =   255
         Left            =   4800
         TabIndex        =   113
         Top             =   2640
         Width           =   5175
      End
      Begin VB.CheckBox Check40 
         Caption         =   "Attiva controllo eliminazione lotti in sincronizzazione"
         Height          =   255
         Left            =   4800
         TabIndex        =   112
         Top             =   2400
         Width           =   5175
      End
      Begin VB.CheckBox Check39 
         Caption         =   "Attiva data ultima sincronizzazione"
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   2400
         Width           =   3495
      End
      Begin DMTEDITNUMLib.dmtNumber txtNEleContattiPerPag 
         Height          =   255
         Left            =   2760
         TabIndex        =   108
         Top             =   2160
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         DecimalPlaces   =   0
         AllowEmpty      =   0   'False
      End
      Begin VB.CheckBox Check36 
         Caption         =   "Attiva paginazione contatti"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CheckBox Check35 
         Caption         =   "Varietà da tipologia coltura"
         Height          =   255
         Left            =   4800
         TabIndex        =   106
         Top             =   2160
         Width           =   4335
      End
      Begin VB.CheckBox Check34 
         Caption         =   "Non sincronizzare codice anagrafica"
         Height          =   255
         Left            =   4800
         TabIndex        =   105
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   5160
         TabIndex        =   92
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   120
         TabIndex        =   90
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   2640
         TabIndex        =   91
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CheckBox Check29 
         Caption         =   "Comune dalla configurazione del produttore"
         Height          =   255
         Left            =   4800
         TabIndex        =   96
         Top             =   1680
         Width           =   4335
      End
      Begin VB.CheckBox Check28 
         Caption         =   "Sincronizza automaticamente il produttore se non esiste"
         Height          =   255
         Left            =   4800
         TabIndex        =   95
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   7680
         TabIndex        =   93
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox Check27 
         Caption         =   "Attiva configurazione multilivello"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Attiva seleziona multipla dei lotti sincronizzati"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1680
         Width           =   4335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Attiva metodo sub contatti"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   4335
      End
      Begin VB.CheckBox chkAttivaFeed 
         Caption         =   "Attiva importazione dati da Feedentity"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtChiaveFeed 
         Height          =   315
         Left            =   5280
         TabIndex        =   34
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtURLFeed 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "SINCRONIZZAZIONE LOTTO DI PRODUZIONE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3000
         TabIndex        =   104
         Top             =   3480
         Width           =   4695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   120
         X2              =   10080
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label1 
         Caption         =   "Codice acquisto merce"
         Height          =   255
         Index           =   13
         Left            =   5160
         TabIndex        =   103
         ToolTipText     =   "Carettere di separazione tra il carettere che indica un fuori quota e le note presenti nel lotto"
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Codice Classificazione 1"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   102
         ToolTipText     =   "Carettere che identifica un lotto di produzione fuori quota presente in Feedentity"
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Codice classificazione 2"
         Height          =   255
         Index           =   12
         Left            =   2640
         TabIndex        =   101
         ToolTipText     =   "Carettere di separazione tra il carettere che indica un fuori quota e le note presenti nel lotto"
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Riferimento Feedentity"
         Height          =   255
         Index           =   8
         Left            =   7680
         TabIndex        =   94
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Chiave"
         Height          =   255
         Index           =   11
         Left            =   5280
         TabIndex        =   39
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "URL"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   720
         Width           =   6975
      End
   End
   Begin VB.Frame fraMigros 
      Caption         =   "Parametri per MIGROS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   10440
      TabIndex        =   24
      Top             =   5880
      Width           =   10215
      Begin VB.TextBox txtPwdMigros 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6600
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtUserMigros 
         Height          =   315
         Left            =   3120
         TabIndex        =   29
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtURLMigros 
         Height          =   315
         Left            =   3120
         TabIndex        =   27
         Top             =   480
         Width           =   6975
      End
      Begin VB.TextBox txtGTIN 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
         Height          =   255
         Index           =   7
         Left            =   6600
         TabIndex        =   35
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nome utente"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   30
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "URL"
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   28
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label Label1 
         Caption         =   "GTIN"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraVendita 
      Caption         =   "Vendita"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   120
      TabIndex        =   22
      Top             =   7680
      Width           =   10215
      Begin VB.CheckBox Check33 
         Caption         =   "Visualizza un messaggio se nella riga non è presente un imballo"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   2640
         Width           =   9975
      End
      Begin VB.CheckBox Check32 
         Caption         =   "Riporta nell'anagrafica di destinazione l'anagrafica di fatturazione presente nella configurazione del cedente"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   2400
         Width           =   9975
      End
      Begin VB.CheckBox Check25 
         Caption         =   "Non riportare l'agente se non è presente nell'ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   2160
         Width           =   9975
      End
      Begin VB.CheckBox Check23 
         Caption         =   "Riporta riferimenti dei documenti collegati per ogni riga della nota di debito"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1920
         Width           =   9975
      End
      Begin VB.CheckBox Check22 
         Caption         =   "Riporta riferimenti dei documenti collegati per ogni riga della nota di credito"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   1680
         Width           =   9975
      End
      Begin VB.CheckBox Check18 
         Caption         =   "Attiva il calcolo a valore per un riscontro peso riferito ad un acquisto merce"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1440
         Width           =   9975
      End
      Begin VB.CheckBox Check17 
         Caption         =   "Attiva il calcolo a valore per un riscontro peso riferito ad un conferimento"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1200
         Width           =   9975
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Attiva riscontro peso per ogni conferimento presente nel documento"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   960
         Width           =   6375
      End
      Begin DmtSearchAccount.DmtSearchACS ACSSocio 
         Height          =   585
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   1032
         WidthDescription=   3800
         WidthSecondDescription=   1500
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SearchType      =   2
         HideLeaf        =   0   'False
         BeginProperty FontLabel {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionDescription=   "Socio/Fornitore per riscontro peso"
         CaptionCode     =   "Codice"
         OnlyAccounts    =   -1  'True
      End
   End
   Begin VB.Frame fraLavorazione 
      Caption         =   "Lavorazione (I° Gamma e IV° Gamma)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   10215
      Begin VB.CheckBox Check10 
         Caption         =   "Modifica la lavorazione senza ricalcolo"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   9975
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Non visualizzare le righe ordine completate "
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   9975
      End
      Begin VB.CheckBox chkAttivaSelSubLottoProd 
         Caption         =   "Attiva la selezione del sub lotto di produzione"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   9975
      End
      Begin VB.CheckBox chkVisNoteLavMsg 
         Caption         =   "Quando viene selezionata una riga di ordine avvia un messaggio con le note di lavorazione presenti"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   9975
      End
      Begin VB.CheckBox chkVisNoteLavElenco 
         Caption         =   "Visualizza le note di lavorazione nell'elenco degli articoli presenti nell'ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   8655
      End
   End
   Begin VB.Frame fraContratto 
      Caption         =   "Contratto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   10215
      Begin VB.CheckBox Check24 
         Caption         =   "Conferma automaticamente la presa in visione"
         Height          =   255
         Left            =   1800
         TabIndex        =   86
         Top             =   600
         Width           =   7815
      End
      Begin VB.CheckBox chkDataScadObbl 
         Caption         =   "La data scadenza deve essere obbligatoria"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   8055
      End
      Begin DMTEDITNUMLib.dmtNumber txtGGDataScadContr 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "GG Data scad."
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Numero di giorni che passano dalla data di inizio contratto"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkNonCalPrzDueRifDaOrd 
      Caption         =   "Non calcolare il prezzo quando nell'ordine sono presenti due righe uguali"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   9975
   End
   Begin VB.CommandButton cmdConferma 
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
      Left            =   120
      TabIndex        =   36
      Top             =   12480
      Width           =   2175
   End
   Begin VB.Frame frmEvazioneOrdini 
      Caption         =   "Evasione ordini"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.CheckBox Check31 
         Caption         =   "Preleva la lettera d'intento dal documento collegato"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   2760
         Width           =   9975
      End
      Begin VB.CheckBox Check30 
         Caption         =   "Preleva l'aliquota I.V.A. dell'articolo dal documento collegato"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   2520
         Width           =   9975
      End
      Begin VB.CheckBox Check15 
         Caption         =   "Alla prezzattura, quando non si trovano associazioni automaticamente, visualizza elenco righe ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   2040
         Width           =   9975
      End
      Begin VB.CheckBox Check14 
         Caption         =   "Ricalcola commissioni per tipo pedana quando si evade un ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1800
         Width           =   9975
      End
      Begin VB.CheckBox Check13 
         Caption         =   "Attiva commissioni da ordine cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   1560
         Width           =   9975
      End
      Begin VB.CheckBox chkAssPedInSpostSing 
         Caption         =   "Assegna di nuovo la pedana quando si porta la merce nell'ordine selezionato"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   9975
      End
      Begin VB.CheckBox chkNonVisMsgImpZeroConfOrd 
         Caption         =   "Non visualizzare il messaggio di prezzi a zero alla conferma dell'ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   9975
      End
      Begin VB.CheckBox chkCalcImpConfOrd 
         Caption         =   "Calcola i prezzi alla ""Conferma"" dell'ordine"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   9975
      End
      Begin VB.CheckBox chkNonCalcImpDaOrdVel 
         Caption         =   "Non calcolare i prezzi da assegnazione veloce della pedana"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9975
      End
   End
End
Attribute VB_Name = "frmAltriParametri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdConferma_Click()
    PAR_NonCalcImpDaAssVeloce = Abs(Me.chkNonCalcImpDaOrdVel.Value)
    PAR_CalcImpAConfOrd = Abs(Me.chkCalcImpConfOrd.Value)
    PAR_NonVisMsgImpZeroConfOrd = Abs(Me.chkNonVisMsgImpZeroConfOrd.Value)
    PAR_AssNewPedDaAssSingola = Abs(Me.chkAssPedInSpostSing.Value)
    PAR_NonCalcPrezzoDueRefArtOrd = Abs(Me.chkNonCalPrzDueRifDaOrd.Value)
    ATT_MULT_SEL_IV_GAMMA = Abs(Me.chkAttSelMult.Value)
    CHIUDI_CONF_QTASEL_ZERO = Abs(Me.chkChiudiConf.Value)
    GG_DATA_SCADENZA_CONTR = Me.txtGGDataScadContr.Value
    FOCUS_IVGAMMA_LAV = Me.cboFocusIVGammaLav.CurrentID
    FOCUS_IVGAMMA_CONF = Me.cboFocusIVGammaConf.CurrentID
    LINK_UM_COOP_SEL_AUT_IVGAMMA = Me.cboUMCoopSelAut.CurrentID
    
    PAR_VIS_NOTE_RIGA_ORD_ELENCO = Abs(Me.chkVisNoteLavElenco.Value)
    PAR_VIS_NOTE_RIGA_ORD = Abs(Me.chkVisNoteLavMsg.Value)
    DATA_SCADENZA_OBBL = Abs(Me.chkDataScadObbl.Value)
    PAR_IDSOCIO_RISC_PESO = Me.ACSSocio.IDAnagrafica
    
    GtinMigros = Me.txtGTIN.Text
    UrlMigros = Me.txtURLMigros.Text
    NomeUtenteMigros = Me.txtUserMigros.Text
    PasswordMigros = Me.txtPwdMigros.Text
    
    ChiaveFeed = Me.txtChiaveFeed.Text
    UrlFeed = Me.txtURLFeed.Text
    AttivaFeed = Me.chkAttivaFeed.Value
    
    ATTIVA_PROD_CONF_LAV = Me.chkAttConfInLav.Value
    ATTIVA_PROD_ORD_LAV = Me.chkAttOrdInLav.Value
    DISATTIVA_PROD_ORD_TESTA = Me.chkAttSoloRigheOrd.Value
    ATTIVA_PROD_CALC_COLLI_PED = Me.chkAttPesoColloPesoPed.Value
    ATTIVA_SEL_SUB_LOTTO_PROD = Me.chkAttivaSelSubLottoProd.Value
    ATTIVA_SUB_CONTATTI_FEED = Me.Check1.Value
    ATTIVA_SEL_MULT_LOTTI_FEED = Me.Check2.Value
    ATTIVA_PROD_FIFO_ACCESSI_CONF = Me.Check3.Value
    
    CodiceAnnullaOperazione = txtCodiceAnnullaOperazione.Text
    CodiceConfermaOrdine = txtCodiceConfermaOrdine.Text
    CodiceGestioneErrori = txtCodiceGestioneErrori.Text
    CodiceConfermaViaggio = txtCodiceConfermaViaggio.Text
    
    
    ATTIVA_GESTIONE_CARICO_MERCE = Me.Check4.Value
    NonVisMsgCreaListaPrelAutPedUscita = Me.Check5.Value
    ATTIVA_VisAndGrigliaConferimento = Me.Check6.Value
    ATTIVA_VisAndGrigliaRigaOdine = Me.Check7.Value
    ATTIVA_RICALCOLO_COLLI_PREL_CHIUSO = Me.Check8.Value
    IDTipoCollegRigaOrdRigaConf = Me.cboTipoCollFiltroConf.CurrentID
    
    NonVisRigheOrdiniComplInSelLav = Me.Check9.Value
    ModificaLavorazioneSenzaRicalcolo = Me.Check10.Value
    RISCONTRO_PESO_PER_CONF = Me.Check11.Value
    ATTIVA_ELIMINAZIONE_ACCESSI = Me.Check12.Value
    
    IDSOCIO_PRE_CONF = Me.ACSSocioPerConf.IDAnagrafica
        
    ATTIVA_COMMISSIONI_ORDINI = Me.Check13.Value
    RIC_COMM_TIPO_PED_EVAS_ORD = Me.Check14.Value
    VIS_ELENCO_RIGHE_ORD = Me.Check15.Value
    ATTIVA_CALCOLO_N_PED_BI_VENDITE = Me.Check16.Value
    RISCONTRO_PESO_VAL_SOCIO = Me.Check17.Value
    RISCONTRO_PESO_VAL_FORNITORE = Me.Check18.Value
    OBBL_N_DOC_PES_CONF = Me.Check19.Value
    ATTIVA_OBBL_SFALCIO_CONF = Me.Check20.Value
    ATTIVA_SEQ_SFALCIO = Me.Check21.Value
    RIPORTA_RIF_DETTAGLIO_XML_NC = Me.Check22.Value
    RIPORTA_RIF_DETTAGLIO_XML_ND = Me.Check23.Value
    CONF_AUT_CONTR_PRESA_VISIONE = Me.Check24.Value
    NO_RIP_AGENTE_IN_DOC_EVASIONE = Me.Check25.Value
    DISATTIVA_SCALATA_COMM_TRASP = Me.Check26.Value
    CONFERMA_ALTRI_PARAMETRI = True
    ATTIVA_FEED_MULTI_LIVELLO = Me.Check27.Value
    IDFEED_AZIENDA = Me.Text1.Text
    SincronizzaAutFeed = Me.Check28.Value
    IvaArticoloDaDocColl = Me.Check30.Value
    LetteraIntentoDaDocColl = Me.Check31.Value
    RifComuneDaConfigSocio = Me.Check29.Value
    DocAnaDestUgualeAnaCoop = Me.Check32.Value
    MsgInDocSeRigaMerceSenzaImballo = Me.Check33.Value
    CodiceCampoIDFeedPerClass01 = Me.Text2.Text
    CodiceCampoIDFeedPerClass02 = Me.Text3.Text
    CodiceCampoIDFeedPerAcquisto = Me.Text4.Text
    NonInviareCodiceForInFeedentity = Me.Check34.Value
    PrendiVarietaDaTipologiaFeedentity = Me.Check35.Value
    AttivaPaginazioneContattiFeed = Me.Check36.Value
    NumeroElementiContattiPerPagina = Me.txtNEleContattiPerPag.Value
    AttivaRicercaFatturaAccontoBIVendite = Me.Check38.Value
    AttivaRicercaFatturaAccontoBIFatturato = Me.Check37.Value
    AttivaGestioneUltimaSincronizzazioneFeed = Me.Check39.Value
    AttivaControlloEsistenzaLottiInFeedAutInSync = Me.Check40.Value
    NonEliminareLottiDefinitivamenteDaFeed = Me.Check41.Value
    NonEliminareLottiProvvDefinitivamenteDaFeed = Me.Check42.Value
    NonAggiornareRifLottoInFeed = Me.Check43.Value
    
    
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    INIT_CONTROLLI
    
    Me.chkNonCalcImpDaOrdVel.Value = PAR_NonCalcImpDaAssVeloce
    Me.chkCalcImpConfOrd.Value = PAR_CalcImpAConfOrd
    Me.chkNonVisMsgImpZeroConfOrd.Value = PAR_NonVisMsgImpZeroConfOrd
    Me.chkAssPedInSpostSing.Value = PAR_AssNewPedDaAssSingola
    Me.chkNonCalPrzDueRifDaOrd.Value = PAR_NonCalcPrezzoDueRefArtOrd
    Me.chkAttSelMult.Value = ATT_MULT_SEL_IV_GAMMA
    Me.chkChiudiConf.Value = CHIUDI_CONF_QTASEL_ZERO
    Me.txtGGDataScadContr.Value = GG_DATA_SCADENZA_CONTR
    Me.cboFocusIVGammaLav.WriteOn FOCUS_IVGAMMA_LAV
    Me.cboFocusIVGammaConf.WriteOn FOCUS_IVGAMMA_CONF
    Me.cboUMCoopSelAut.WriteOn LINK_UM_COOP_SEL_AUT_IVGAMMA
    Me.chkVisNoteLavElenco.Value = PAR_VIS_NOTE_RIGA_ORD_ELENCO
    Me.chkVisNoteLavMsg.Value = PAR_VIS_NOTE_RIGA_ORD
    Me.chkDataScadObbl.Value = DATA_SCADENZA_OBBL
    Me.ACSSocio.sbLoadCFByIDAnagrafica 7, PAR_IDSOCIO_RISC_PESO
    
    Me.txtGTIN.Text = GtinMigros
    Me.txtURLMigros.Text = UrlMigros
    Me.txtUserMigros.Text = NomeUtenteMigros
    Me.txtPwdMigros.Text = PasswordMigros
    
    Me.txtChiaveFeed.Text = ChiaveFeed
    Me.txtURLFeed.Text = UrlFeed
    Me.chkAttivaFeed.Value = AttivaFeed
    
    Me.chkAttConfInLav.Value = ATTIVA_PROD_CONF_LAV
    Me.chkAttOrdInLav.Value = ATTIVA_PROD_ORD_LAV
    Me.chkAttSoloRigheOrd.Value = DISATTIVA_PROD_ORD_TESTA
    Me.chkAttPesoColloPesoPed.Value = ATTIVA_PROD_CALC_COLLI_PED
    

    Me.chkAttivaSelSubLottoProd.Value = ATTIVA_SEL_SUB_LOTTO_PROD
    
    txtCodiceAnnullaOperazione.Text = CodiceAnnullaOperazione
    txtCodiceConfermaOrdine.Text = CodiceConfermaOrdine
    txtCodiceGestioneErrori.Text = CodiceGestioneErrori
    txtCodiceConfermaViaggio.Text = CodiceConfermaViaggio
    
    Me.Check1.Value = ATTIVA_SUB_CONTATTI_FEED
    Me.Check2.Value = ATTIVA_SEL_MULT_LOTTI_FEED
    Me.Check3.Value = ATTIVA_PROD_FIFO_ACCESSI_CONF
    Me.Check4.Value = ATTIVA_GESTIONE_CARICO_MERCE
    Me.Check5.Value = NonVisMsgCreaListaPrelAutPedUscita
    Me.Check6.Value = ATTIVA_VisAndGrigliaConferimento
    Me.Check7.Value = ATTIVA_VisAndGrigliaRigaOdine
    Me.Check8.Value = ATTIVA_RICALCOLO_COLLI_PREL_CHIUSO
    Me.cboTipoCollFiltroConf.WriteOn IDTipoCollegRigaOrdRigaConf
    
    Me.Check9.Value = NonVisRigheOrdiniComplInSelLav
    Me.Check10.Value = ModificaLavorazioneSenzaRicalcolo
    Me.Check11.Value = RISCONTRO_PESO_PER_CONF
    Me.Check12.Value = ATTIVA_ELIMINAZIONE_ACCESSI
    
    Me.ACSSocioPerConf.sbLoadCFByIDAnagrafica 7, IDSOCIO_PRE_CONF
    
    Me.Check13.Value = ATTIVA_COMMISSIONI_ORDINI
    Me.Check14.Value = RIC_COMM_TIPO_PED_EVAS_ORD
    Me.Check15.Value = VIS_ELENCO_RIGHE_ORD
    
    Me.Check16.Value = ATTIVA_CALCOLO_N_PED_BI_VENDITE
    
    Me.Check17.Value = RISCONTRO_PESO_VAL_SOCIO
    Me.Check18.Value = RISCONTRO_PESO_VAL_FORNITORE
    Me.Check19.Value = OBBL_N_DOC_PES_CONF
    Me.Check20.Value = ATTIVA_OBBL_SFALCIO_CONF
    Me.Check21.Value = ATTIVA_SEQ_SFALCIO
    Me.Check22.Value = RIPORTA_RIF_DETTAGLIO_XML_NC
    Me.Check23.Value = RIPORTA_RIF_DETTAGLIO_XML_ND
    Me.Check24.Value = CONF_AUT_CONTR_PRESA_VISIONE
    Me.Check25.Value = NO_RIP_AGENTE_IN_DOC_EVASIONE
    Me.Check26.Value = DISATTIVA_SCALATA_COMM_TRASP
    Me.Check27.Value = ATTIVA_FEED_MULTI_LIVELLO
    Me.Text1.Text = IDFEED_AZIENDA
    Me.Check28.Value = SincronizzaAutFeed
    Me.Check30.Value = IvaArticoloDaDocColl
    Me.Check31.Value = LetteraIntentoDaDocColl
    Me.Check29.Value = RifComuneDaConfigSocio
    Me.Check32.Value = DocAnaDestUgualeAnaCoop
    Me.Check33.Value = MsgInDocSeRigaMerceSenzaImballo
    Me.Text2.Text = CodiceCampoIDFeedPerClass01
    Me.Text3.Text = CodiceCampoIDFeedPerClass02
    Me.Text4.Text = CodiceCampoIDFeedPerAcquisto
    Me.Check34.Value = NonInviareCodiceForInFeedentity
    Me.Check35.Value = PrendiVarietaDaTipologiaFeedentity
    Me.Check36.Value = AttivaPaginazioneContattiFeed
    Me.txtNEleContattiPerPag.Value = NumeroElementiContattiPerPagina
    Me.Check38.Value = AttivaRicercaFatturaAccontoBIVendite
    Me.Check37.Value = AttivaRicercaFatturaAccontoBIFatturato
    Me.Check39.Value = AttivaGestioneUltimaSincronizzazioneFeed
    
    Me.Check40.Value = AttivaControlloEsistenzaLottiInFeedAutInSync
    Me.Check41.Value = NonEliminareLottiDefinitivamenteDaFeed
    Me.Check42.Value = NonEliminareLottiProvvDefinitivamenteDaFeed
    Me.Check43.Value = NonAggiornareRifLottoInFeed
    
End Sub
Private Sub INIT_CONTROLLI()
    
    With Me.cboFocusIVGammaLav
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POFocusControlIVGamma"
        .DisplayField = "Descrizione"
        .SQL = "SELECT * FROM RV_POFocusControlIVGamma "
        .SQL = .SQL & "WHERE Lavorazione=1"
        .Fill
    End With

    With Me.cboFocusIVGammaConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POFocusControlIVGamma"
        .DisplayField = "Descrizione"
        .SQL = "SELECT * FROM RV_POFocusControlIVGamma "
        .SQL = .SQL & "WHERE Conferimento=1"
        .Fill
    End With

    With Me.cboFocusIVGammaConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POFocusControlIVGamma"
        .DisplayField = "Descrizione"
        .SQL = "SELECT * FROM RV_POFocusControlIVGamma "
        .SQL = .SQL & "WHERE Conferimento=1"
        .Fill
    End With

    With Me.cboUMCoopSelAut
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POUnitaDiMisuraCoop"
        .DisplayField = "UnitaDiMisuraCoop"
        .SQL = "SELECT * FROM RV_POUnitaDiMisuraCoop "
        .Fill
    End With
    
    With Me.cboTipoCollFiltroConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoCollegamentoRigaOrdRigaConf"
        .DisplayField = "TipoCollegamentoRigaOrdRigaConf"
        .SQL = "SELECT * FROM RV_POTipoCollegamentoRigaOrdRigaConf "
        .Fill
    End With
    
    
    Set Me.ACSSocio.Connection = TheApp.Database.Connection
    ACSSocio.ApplicationName = App.Title
    ACSSocio.Client = App.EXEName
    ACSSocio.IDFirm = TheApp.IDFirm
    ACSSocio.IDUser = TheApp.IDUser
    ACSSocio.UserName = TheApp.User
    ACSSocio.SearchType = DmtSearchSuppliers
    ACSSocio.HwndContainer = Me.hwnd
    
    Set Me.ACSSocioPerConf.Connection = TheApp.Database.Connection
    ACSSocioPerConf.ApplicationName = App.Title
    ACSSocioPerConf.Client = App.EXEName
    ACSSocioPerConf.IDFirm = TheApp.IDFirm
    ACSSocioPerConf.IDUser = TheApp.IDUser
    ACSSocioPerConf.UserName = TheApp.User
    ACSSocioPerConf.SearchType = DmtSearchSuppliers
    ACSSocioPerConf.HwndContainer = Me.hwnd
    
End Sub

