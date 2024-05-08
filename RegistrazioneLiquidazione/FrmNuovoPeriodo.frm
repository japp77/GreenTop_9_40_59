VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{F95AA20B-3F80-11D3-A741-00105A2E9BAF}#2.1#0"; "DmtSearchAccount2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmNuovoPeriodo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creazione liquidazione (2 di 4)"
   ClientHeight    =   12270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17070
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmNuovoPeriodo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12270
   ScaleWidth      =   17070
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7320
      Width           =   16935
   End
   Begin VB.CommandButton cmdEliminaLiquidazione 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   12000
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERIODO DI LIQUIDAZIONE"
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
      Height          =   6855
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   16935
      Begin VB.Frame Frame5 
         Caption         =   "ALTRE IMPOSTAZIONI"
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
         Left            =   120
         TabIndex        =   69
         Top             =   5640
         Width           =   9975
         Begin VB.CheckBox chkConguaglio 
            Alignment       =   1  'Right Justify
            Caption         =   "LIQUIDAZIONE DI CONGUAGLIO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   15
            Top             =   600
            Width           =   3495
         End
         Begin DMTEDITNUMLib.dmtCurrency txtTrattenutaImporto_Nuovo 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   " 0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            CurrencySymbol  =   ""
            AllowEmpty      =   0   'False
            DecFinalZeros   =   -1  'True
         End
         Begin DMTEDITNUMLib.dmtCurrency txtTrattenutaPercentuale_Nuovo 
            Height          =   315
            Left            =   1560
            TabIndex        =   14
            Top             =   600
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   " 0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            CurrencySymbol  =   ""
            AllowEmpty      =   0   'False
            DecFinalZeros   =   -1  'True
         End
         Begin DMTEDITNUMLib.dmtCurrency txtDecQtaLiq 
            Height          =   315
            Left            =   3000
            TabIndex        =   79
            Top             =   600
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   " 0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            CurrencyDecimalPlaces=   0
            CurrencySymbol  =   ""
            AllowEmpty      =   0   'False
            DecFinalZeros   =   -1  'True
         End
         Begin DMTEDITNUMLib.dmtCurrency txtDecImpUni 
            Height          =   315
            Left            =   4440
            TabIndex        =   81
            Top             =   600
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   " 0"
            BackColor       =   16777215
            Appearance      =   1
            UseSeparator    =   -1  'True
            CurrencyDecimalPlaces=   0
            CurrencySymbol  =   ""
            AllowEmpty      =   0   'False
            DecFinalZeros   =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Dec. imp. uni."
            Height          =   255
            Index           =   20
            Left            =   4440
            TabIndex        =   82
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Dec. q.tà liq."
            Height          =   255
            Index           =   16
            Left            =   3000
            TabIndex        =   80
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tratt. Importo"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tratt. Perc."
            Height          =   255
            Index           =   18
            Left            =   1560
            TabIndex        =   70
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "PARAMETRI DI LIQUIDAZIONE"
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
         Height          =   6615
         Left            =   10200
         TabIndex        =   60
         Top             =   120
         Width           =   6615
         Begin VB.CheckBox chkNonCalcPreLiqInND 
            Caption         =   "Non calcolare le trattenute pre liquidazione nelle note di debito"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   6120
            Width           =   6255
         End
         Begin VB.CheckBox chkNonCalcPreLiqInNC 
            Caption         =   "Non calcolare le trattenute pre liquidazione nelle note di credito"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   5760
            Width           =   6255
         End
         Begin VB.CheckBox chkNoRipScarti 
            Caption         =   "Non riportare gli scarti nella liquidazione"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   5400
            Width           =   6255
         End
         Begin VB.CheckBox chkRicalcolaTutto 
            Caption         =   "Al calcolo dell'importo unitario di liquidazione ricalcola la quantità di liquidazione e l'incidenza imballo"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   5040
            Width           =   6375
         End
         Begin VB.CheckBox chkAttivaQtaAbb 
            Caption         =   "Attiva il calcolo della quantità da abbattere"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   4680
            Width           =   6255
         End
         Begin VB.CheckBox chkNoLiqVendUff 
            Caption         =   "Non liquidare vendita collegata ad una liquidazione ufficiale"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   4320
            Width           =   6255
         End
         Begin VB.CheckBox chkNonCalcIncidenzaImb 
            Caption         =   "Non calcolare incidenza imballo nelle note di credito e note di debito"
            Height          =   375
            Left            =   120
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   3960
            Width           =   6255
         End
         Begin VB.CheckBox chkNonCalcComm 
            Caption         =   "Non calcolare le commissioni nelle note di credito e note di debito"
            Height          =   375
            Left            =   120
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   3600
            Width           =   6255
         End
         Begin VB.CheckBox chkRicalcolaPrezziLiq 
            Caption         =   "Ricalcola valori di liquidazione"
            Height          =   375
            Left            =   120
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   3240
            Width           =   6255
         End
         Begin VB.CheckBox chkAggCostoConfezInPrezzoLiq 
            Caption         =   "Aggiungi il costo della confezione al prezzo di liquidazione"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   2520
            Width           =   6255
         End
         Begin VB.CheckBox ChkAggCostoKitInPrezzoLiq 
            Caption         =   "Aggiungi il costo del KIT al prezzo di liquidazione"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2880
            Width           =   6255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Non calcolare le trattenute su articoli di quadratura"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   2160
            Width           =   6255
         End
         Begin VB.CheckBox chkCalcolaPMCamp 
            Caption         =   "Aggiorna importi delle righe di campionature con i prezzi medi"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   360
            Width           =   6255
         End
         Begin VB.CheckBox chkAggPMCamp 
            Caption         =   "Aggiorna solo campionature non bloccate"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   720
            Width           =   6255
         End
         Begin VB.CheckBox chkCalcTrattCamp 
            Caption         =   "Calcola trattenute per le righe di campionatura"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   1080
            Width           =   6255
         End
         Begin VB.CheckBox chkPrzMedioIVGamma 
            Caption         =   "Calcola prezzo medio in IV Gamma"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1440
            Width           =   6255
         End
         Begin VB.CheckBox chkCollegNotaCodiceLotto 
            Caption         =   "Collegamento per codice lotto con la Nota di credito e nota di debito"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1800
            Width           =   6255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "FILTRI DI LIQUIDAZIONE"
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
         Height          =   4095
         Left            =   120
         TabIndex        =   57
         Top             =   1680
         Width           =   9975
         Begin VB.CheckBox chkLiquidazionePerFor 
            Caption         =   "Liquida fornitori"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
         Begin DMTDataCmb.DMTCombo cboListino 
            Height          =   315
            Left            =   3480
            TabIndex        =   12
            Top             =   1680
            Width           =   3255
            _ExtentX        =   5741
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
         Begin DmtSearchAccount2.DmtSearchACS2 ACS 
            Height          =   600
            Left            =   2160
            TabIndex        =   9
            Top             =   240
            Width           =   7620
            _ExtentX        =   13441
            _ExtentY        =   1058
            WidthCode       =   1500
            WidthDescription=   4500
            WidthSecondDescription=   1500
            Object.Visible         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
            CaptionDescription=   "Socio/Fornitore"
            CaptionCode     =   "Codice:"
            OnlyAccounts    =   -1  'True
         End
         Begin DMTDataCmb.DMTCombo cboCategoriaMerc 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   1680
            Width           =   3255
            _ExtentX        =   5741
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
         Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
            Height          =   615
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   795
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   1085
            PropCodice      =   $"FrmNuovoPeriodo.frx":4781A
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmNuovoPeriodo.frx":47872
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmNuovoPeriodo.frx":478E2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label Label1 
            Caption         =   "Listino"
            Height          =   255
            Index           =   15
            Left            =   3480
            TabIndex        =   59
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Categoria liquidazione"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   1440
            Width           =   3255
         End
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroProtocolloInterno 
         Height          =   315
         Left            =   7440
         TabIndex        =   4
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboTipoLiquidazione 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
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
      Begin DMTDataCmb.DMTCombo cboTipoQuantita 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
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
      Begin DMTDataCmb.DMTCombo cboTipoImportoDocumento 
         Height          =   315
         Left            =   8880
         TabIndex        =   39
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
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
      Begin DMTDataCmb.DMTCombo cboTipoImportoArticolo 
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
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
      Begin DMTEDITNUMLib.dmtNumber txtNumeroLiquidazione 
         Height          =   315
         Left            =   6120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataFine 
         Height          =   315
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataInizio 
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTDataCmb.DMTCombo cboTipoLiqConf 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.Label Label3 
         Caption         =   "Tipo di liquidazione conferimenti"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "N° interno"
         Height          =   255
         Index           =   5
         Left            =   7440
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Calcolo trattenute a % su:"
         Height          =   255
         Index           =   4
         Left            =   6840
         TabIndex        =   38
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Calcolo trattenute a valore su:"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   36
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "N° Liq."
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   35
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Data fine"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Data inizio"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdAvanti 
      Caption         =   "Avanti"
      Height          =   375
      Left            =   15840
      TabIndex        =   16
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   14640
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton cmdIndietro 
      Caption         =   "Indietro"
      Height          =   375
      Left            =   13440
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   11160
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   11880
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modifica periodo di liquidazione"
      Height          =   4935
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   2655
      Begin DMTEDITNUMLib.dmtCurrency txtTrattenutaGeneralePercentuale 
         Height          =   315
         Left            =   1560
         TabIndex        =   31
         Top             =   3720
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "€ 0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencySymbol  =   "€"
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin DMTEDITNUMLib.dmtCurrency txtImportoTrattenutaGeneralePerRiga 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   253
         Text            =   "€ 0"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         UseSeparator    =   -1  'True
         CurrencySymbol  =   "€"
         AllowEmpty      =   0   'False
         DecFinalZeros   =   -1  'True
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Non calcolare le trattenute su articoli di quadratura"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   32
         Top             =   3720
         Width           =   3615
      End
      Begin DMTDataCmb.DMTCombo cboTipoLiquidazione_PerModifica 
         Height          =   315
         Left            =   4440
         TabIndex        =   42
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTDataCmb.DMTCombo cboPeriodo 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
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
      Begin DMTEDITNUMLib.dmtNumber TxtNumeroLiquidazione_PerModifica 
         Height          =   315
         Left            =   3240
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataInizio_PerModifica 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
      End
      Begin DMTDATETIMELib.dmtDate txtDataFinePerModifica 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
      End
      Begin DmtSearchAccount2.DmtSearchACS2 ACS_PerModifica 
         Height          =   600
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   1058
         WidthDescription=   4000
         WidthSecondDescription=   2000
         Object.Visible         =   0   'False
         VisibleCode     =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         CaptionDescription=   "Socio/Fornitore"
         CaptionCode     =   "Codice:"
         OnlyAccounts    =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboTipoImportoArticolo_PerModifica 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   2880
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
      Begin DMTDataCmb.DMTCombo cboTipoQuantita_PerModifica 
         Height          =   315
         Left            =   3360
         TabIndex        =   29
         Top             =   2880
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
      Begin DMTDataCmb.DMTCombo cboTipoImportoDocumento_PerModifica 
         Height          =   315
         Left            =   3120
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin DMTEDITNUMLib.dmtNumber txtNumeroProtInt_perModifica 
         Height          =   315
         Left            =   4800
         TabIndex        =   26
         Top             =   1320
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   253
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "% trattenuta"
         Height          =   255
         Index           =   14
         Left            =   1560
         TabIndex        =   54
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Importo trattenuta"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   53
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "N° di liq."
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   49
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Data inizio"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Data fine"
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   47
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Calcolo trattenute a % su:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   46
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Calcolo trattenute a valore su:"
         Height          =   255
         Index           =   7
         Left            =   3360
         TabIndex        =   45
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "N° di prot. int."
         Height          =   255
         Index           =   6
         Left            =   4800
         TabIndex        =   44
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Label lblInfoStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   52
      Top             =   11640
      Width           =   17055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ELABORAZIONE LIQUIDAZIONE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   51
      Top             =   6960
      Width           =   16935
   End
End
Attribute VB_Name = "FrmNuovoPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub InitVariabili()
On Error GoTo ERR_InitVariabili
If Nuova_Liquidazione = 1 Then
    DATA_INIZIO = Me.txtDataInizio.Text
    DATA_FINE = Me.txtDataFine.Text
    NUMERO_LIQUIDAZIONE = Me.txtNumeroLiquidazione.Value
    LINK_SOCIO = Me.ACS.IDAnagrafica
    LINK_SOCIO_SEL = Me.ACS.IDAnagrafica
    TIPO_IMPORTO_ARTICOLO = Me.cboTipoImportoArticolo.CurrentID
    TIPO_IMPORTO_DOCUMENTO = Me.cboTipoImportoDocumento.CurrentID
    TIPO_QUANTITA = Me.cboTipoQuantita.CurrentID
    ARTICOLI_DI_QUAD = Me.Check1.Value
    TIPO_LIQUIDAZIONE = Me.cboTipoLiquidazione.CurrentID
    LINK_PERIODO = fnGetNewKey("RV_POLiquidazionePeriodo", "IDRV_POLiquidazionePeriodo")
    NUMERO_PROTOCOLLO = Me.txtNumeroProtocolloInterno.Value
    TRATTENUTA_PER_IMPORTO = Me.txtTrattenutaImporto_Nuovo.Value
    TRATTENUTA_PER_PERCENTUALE = Me.txtTrattenutaPercentuale_Nuovo.Value
    LINK_ARTICOLO = Me.CDArticolo.KeyFieldID
    LIQUIDA_FORNITORE = Me.chkLiquidazionePerFor.Value
    RICALCOLA_VALORI_LIQ = Me.chkRicalcolaPrezziLiq.Value
        

    LINK_TIPO_LIQ_CONF = Me.cboTipoLiqConf.CurrentID    'Utilizzare per liquidare il tipo di stato della riga di conferimento
    CALCOLA_PM_CAMP = Me.chkCalcolaPMCamp.Value         'Indica se bisogna calcolare il prezzo medio nella campionatura e aggiornare la riga
    AGGIORNA_PM_CAMP = Me.chkAggPMCamp.Value            'Aggiorna i prezzi medi delle campionature non bloccati ad ogni elaborazione
    CALCOLA_TRATT_CAMP = Me.chkCalcTrattCamp.Value      'Indica che deve calcolare le trattenute nella campionature
    LIQ_CONGUAGLIO = Me.chkConguaglio.Value             'Indica che è una liquidazione di conguaglio
    
    LINK_LISTINO = Me.cboListino.CurrentID
    COLLEGAMENTO_NOTA_PER_LOTTO = Abs(Me.chkCollegNotaCodiceLotto.Value)
    LINK_CAT_MERCE = Me.cboCategoriaMerc.CurrentID
    AGG_COSTO_CONFEZ_PRZ_LIQ = Abs(Me.chkAggCostoConfezInPrezzoLiq.Value)
    AGG_COSTO_KIT_PRZ_LIQ = Abs(Me.ChkAggCostoKitInPrezzoLiq.Value)
    NON_CALC_COMM = Abs(Me.chkNonCalcComm.Value)
    NON_CALC_INCIDENZA_IMB = Abs(Me.chkNonCalcIncidenzaImb.Value)
    NO_LIQ_VEND_UFF = Abs(Me.chkNoLiqVendUff.Value)
    ATTIVA_CALCOLO_QTA_DA_ABB = Abs(Me.chkAttivaQtaAbb.Value)
    RICALCOLA_TUTTO = Abs(Me.chkRicalcolaTutto.Value)
    DEC_QTA_LIQ = Me.txtDecQtaLiq.Value
    DEC_IMP_UNI_LIQ = Me.txtDecImpUni.Value
    NO_CALC_PRELIQ_NC = Me.chkNonCalcPreLiqInNC.Value
    NO_CALC_PRELIQ_ND = Me.chkNonCalcPreLiqInND.Value
End If

If Nuova_Liquidazione = 0 Then
    DATA_INIZIO = Me.txtDataInizio_PerModifica.Text
    DATA_FINE = Me.txtDataFinePerModifica.Text
    NUMERO_LIQUIDAZIONE = Me.TxtNumeroLiquidazione_PerModifica.Value
    LINK_PERIODO = Me.cboPeriodo.CurrentID
    LINK_SOCIO = Me.ACS_PerModifica.IDAnagrafica
    LINK_SOCIO_SEL = Me.ACS.IDAnagrafica
    TIPO_IMPORTO_ARTICOLO = Me.cboTipoImportoArticolo_PerModifica.CurrentID
    TIPO_IMPORTO_DOCUMENTO = Me.cboTipoImportoDocumento_PerModifica.CurrentID
    TIPO_QUANTITA = Me.cboTipoQuantita_PerModifica.CurrentID
    ARTICOLI_DI_QUAD = Abs(Me.Check2.Value)
    TIPO_LIQUIDAZIONE = Me.cboTipoLiquidazione_PerModifica.CurrentID
    NUMERO_PROTOCOLLO = Me.txtNumeroProtInt_perModifica.Value
    TRATTENUTA_PER_IMPORTO = Me.txtImportoTrattenutaGeneralePerRiga.Value
    TRATTENUTA_PER_PERCENTUALE = Me.txtTrattenutaGeneralePercentuale.Value
    RICALCOLA_VALORI_LIQ = Abs(Me.chkRicalcolaPrezziLiq.Value)

    LINK_TIPO_LIQ_CONF = Me.cboTipoLiqConf.CurrentID    'Utilizzare per liquidare il tipo di stato della riga di conferimento
    CALCOLA_PM_CAMP = Abs(Me.chkCalcolaPMCamp.Value)         'Indica se bisogna calcolare il prezzo medio nella campionatura e aggiornare la riga
    AGGIORNA_PM_CAMP = Abs(Me.chkAggPMCamp.Value)            'Aggiorna i prezzi medi delle campionature non bloccati ad ogni elaborazione
    CALCOLA_TRATT_CAMP = Abs(Me.chkCalcTrattCamp.Value)      'Indica che deve calcolare le trattenute nella campionature
    LIQ_CONGUAGLIO = Abs(Me.chkConguaglio.Value)            'Indica che è una liquidazione di conguaglio
    
    LINK_LISTINO = Abs(Me.cboListino.CurrentID)
    LINK_CAT_MERCE = Abs(Me.cboCategoriaMerc.CurrentID)
    COLLEGAMENTO_NOTA_PER_LOTTO = Abs(Me.chkCollegNotaCodiceLotto.Value)
    NON_CALC_COMM = Abs(Me.chkNonCalcComm.Value)
    NON_CALC_INCIDENZA_IMB = Abs(Me.chkNonCalcIncidenzaImb.Value)
    NO_LIQ_VEND_UFF = Abs(Me.chkNoLiqVendUff.Value)
    ATTIVA_CALCOLO_QTA_DA_ABB = Abs(Me.chkAttivaQtaAbb.Value)
    RICALCOLA_TUTTO = Abs(Me.chkRicalcolaTutto.Value)
    DEC_QTA_LIQ = Me.txtDecQtaLiq.Value
    DEC_IMP_UNI_LIQ = Me.txtDecImpUni.Value
    NO_CALC_PRELIQ_NC = Me.chkNonCalcPreLiqInNC.Value
    NO_CALC_PRELIQ_ND = Me.chkNonCalcPreLiqInND.Value
    
End If

TIPO_QUADRATURA = ParametroTipoQuadratura

ParametroSocio

Exit Sub
ERR_InitVariabili:
    MsgBox Err.Description, vbCritical, "InitVariabili"
End Sub



Private Sub cboPeriodo_Click()
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT RV_POLiquidazionePeriodo.IDRV_POLiquidazionePeriodo, RV_POLiquidazionePeriodo.Periodo, RV_POLiquidazionePeriodo.NumeroLiquidazione, "
sSQL = sSQL & "RV_POLiquidazionePeriodo.DataInizio, RV_POLiquidazionePeriodo.DataFine, RV_POLiquidazionePeriodo.IDTipoImportoArticolo,"
sSQL = sSQL & "RV_POLiquidazionePeriodo.IDSocio , Anagrafica.Anagrafica, Anagrafica.Nome, "
sSQL = sSQL & "RV_POLiquidazionePeriodo.IDTipoImportoDocumento, RV_POLiquidazionePeriodo.IDTipoQuantita, "
sSQL = sSQL & "RV_POLiquidazionePeriodo.ArticoliDiQuadratura, RV_POLiquidazionePeriodo.IDTipoLiquidazione, RV_POLiquidazionePeriodo.NumeroProtInt, "
sSQL = sSQL & "RV_POLiquidazionePeriodo.TrattenutaRigaImporto, RV_POLiquidazionePeriodo.TrattenutaRigaPercentuale "
sSQL = sSQL & "FROM RV_POLiquidazionePeriodo LEFT OUTER JOIN "
sSQL = sSQL & "Anagrafica ON RV_POLiquidazionePeriodo.IDSocio = Anagrafica.IDAnagrafica "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection


If rs.EOF = False Then
    Me.txtDataInizio_PerModifica.Text = fnNotNull(rs!DataInizio)
    Me.txtDataFinePerModifica.Text = fnNotNull(rs!DataFine)
    Me.TxtNumeroLiquidazione_PerModifica.Value = fnNotNullN(rs!NumeroLiquidazione)
    Me.cboTipoImportoArticolo_PerModifica.WriteOn fnNotNullN(rs!IDTipoImportoArticolo)
    Me.cboTipoImportoDocumento_PerModifica.WriteOn fnNotNullN(rs!IDTipoImportoDocumento)
    Me.cboTipoQuantita_PerModifica.WriteOn fnNotNullN(rs!IDTipoQuantita)
    Me.Check2.Value = fnNormBoolean(fnNotNullN(rs!ArticoliDiQuadratura))
    Me.cboTipoLiquidazione_PerModifica.WriteOn fnNotNullN(rs!IDTipoLiquidazione)
    Me.txtNumeroProtInt_perModifica.Value = fnNotNullN(rs!NumeroProtInt)
    Me.ACS_PerModifica.Description = fnNotNull(rs!Anagrafica)
    Me.ACS_PerModifica.SecondDescription = fnNotNull(rs!Nome)
    Me.ACS_PerModifica.IDAnagrafica = fnNotNullN(rs!IDSocio)
    Me.txtTrattenutaGeneralePercentuale = fnNotNullN(rs!TrattenutaRigaPercentuale)
    Me.txtImportoTrattenutaGeneralePerRiga.Value = fnNotNullN(rs!TrattenutaRigaImporto)
    
Else
    Me.txtDataInizio_PerModifica.Text = Date
    Me.txtDataFinePerModifica.Text = Date
    Me.ACS_PerModifica.IDAnagrafica = 0
    Me.ACS_PerModifica.Description = ""
    Me.ACS_PerModifica.SecondDescription = ""
    Me.TxtNumeroLiquidazione_PerModifica.Value = 0
    Me.cboTipoImportoArticolo_PerModifica.WriteOn 0
    Me.cboTipoImportoDocumento_PerModifica.WriteOn 0
    Me.cboTipoQuantita_PerModifica.WriteOn 0
    Me.Check2.Value = Unchecked
    Me.cboTipoLiquidazione_PerModifica.WriteOn 0
    Me.txtNumeroProtInt_perModifica.Value = 0
    Me.txtTrattenutaGeneralePercentuale = 0
    Me.txtImportoTrattenutaGeneralePerRiga.Value = 0
    
End If

rs.Close
Set rs = Nothing

InitVariabili
End Sub

Private Sub cboTipoLiquidazione_Click()

    If Me.cboTipoLiquidazione.CurrentID = 3 Then
        Me.chkLiquidazionePerFor.Enabled = False
        Me.ACS.Enabled = False
        Me.CDArticolo.Enabled = False
    Else
        Me.chkLiquidazionePerFor.Enabled = True
        Me.ACS.Enabled = True
        Me.CDArticolo.Enabled = True
    End If
    
End Sub

Private Sub cboTipoQuantita_Click()

If Me.cboTipoQuantita.CurrentID = 3 Then
    Me.Check1.Value = Checked
    'Me.Enabled = False
Else
    Me.Check1.Value = Unchecked
    'Me.Check1.Enabled = True
End If
End Sub

Private Sub cmdAnnulla_Click()
    If MsgBox("Vuoi abbandonare il wizard per la creazione della liquidazione?", vbQuestion + vbYesNo, "Creazione liquidazione") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdAvanti_Click()
'On Error GoTo ERR_cmdAvanti_Click
Dim Stringa As String

Stringa = Permesso

If Stringa = "" Then
    InitVariabili
    'LINK_PERIODO = SalvaPeriodo
    Me.List1.Clear
    If LINK_PERIODO > 0 Then
        Me.Enabled = False
        EsecuzioneElaborazione Me.ProgressBar1
        If TIPO_LIQUIDAZIONE = 3 Then
            If CONFERMA_SEL_DOCUMENTI = 0 Then
                Me.Enabled = True
                Me.SetFocus
            Else
                Me.Enabled = True
                Unload Me
            End If
        Else
            Me.Enabled = True
            Unload Me
        End If
    Else
        MsgBox "Ci sono stati degli errori da non permettere la prosecuzione del'elaborazione", vbCritical, "Creazione documenti"
    End If
Else
    MsgBox Stringa, vbInformation, "Elaborazione liquidazione"
    Me.Enabled = True
End If
Exit Sub
ERR_cmdAvanti_Click:
    MsgBox Err.Description, vbCritical, "cmdAvanti_Click"
End Sub

Private Sub cmdEliminaLiquidazione_Click()
Dim TESTO As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


TESTO = "ATTENZIONE!!!" & vbCrLf
TESTO = TESTO & "Con questo comando si eliminano tutte le liquidazioni riferite al periodo in questione" & vbCrLf
TESTO = TESTO & "Continuare?"

If MsgBox(TESTO, vbQuestion + vbYesNo, "Eliminazione periodo di liquidazione") = vbYes Then
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Max = 1000

'''''''''''''''''''''''ELIMINAZIONE RIGHE AGGIUNTIVE'''''''''''''''''''''''
    sSQL = "SELECT IDRV_POLiquidazione FROM RV_POLiquidazione "
    sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        CnDMT.Execute "DELETE FROM RV_POLiquidazioneRighe WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POLiquidazione)
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''ELIMINAZIONE RIGHE ELABORATE''''''''''''''''''''''''
    
    CnDMT.Execute "DELETE FROM RV_POLiquidazioneRigheEla WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''ELIMINAZIONE TESTA''''''''''''''''''''''''
    
    CnDMT.Execute "DELETE FROM RV_POLiquidazione WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''ELIMINAZIONE PERIODO''''''''''''''''''''''''
    
    CnDMT.Execute "DELETE FROM RV_POLiquidazionePeriodo WHERE IDRV_POLiquidazionePeriodo=" & Me.cboPeriodo.CurrentID
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 250
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Unload Me
End If
End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub



Private Sub Form_Load()

    'Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fncTipoImportoArticolo
    fncSocio
    fncArticolo
    fncTipoImportoDocumento
    fncTipoQuantita
    fncTipoLiquidazione
    fncPeriodo
    fncListino
    fncCatMerc
    Me.chkRicalcolaPrezziLiq.Value = 1
    
If Nuova_Liquidazione = 1 Then
    Me.Frame1.Visible = True
    Me.Frame2.Visible = False
    Me.cmdEliminaLiquidazione.Visible = False
    'fncTipoImportoDaLiquidare
    Me.txtNumeroLiquidazione.Value = fncNumeroLiquidazione
    Me.txtNumeroProtocolloInterno.Value = fncNumeroProtocolloInterno
    Me.txtDataInizio.Text = Date
    Me.txtDataFine.Text = Date
    RecuperaDatiFiliale
Else
    Me.Frame1.Visible = False
    Me.Frame2.Visible = True
    Me.cmdEliminaLiquidazione.Visible = True
End If
    
    
End Sub
Private Sub RecuperaDatiFiliale()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDSocio=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboTipoImportoArticolo.WriteOn 0
    Me.cboTipoImportoDocumento.WriteOn 0
    Me.cboTipoQuantita.WriteOn 0
    Me.Check1.Value = Unchecked
    Me.cboTipoLiquidazione.WriteOn 0
    Me.cboTipoLiqConf.WriteOn 0
    Me.chkCalcolaPMCamp.Value = 0
    Me.chkAggPMCamp.Value = 0
    Me.chkCalcTrattCamp.Value = 0
    Me.chkPrzMedioIVGamma.Value = 0
    Me.chkCollegNotaCodiceLotto.Value = 0
    Me.chkAggCostoConfezInPrezzoLiq.Value = 0
    Me.ChkAggCostoKitInPrezzoLiq.Value = 0
    Me.chkNonCalcComm.Value = 0
    Me.chkNonCalcIncidenzaImb.Value = 0
    Me.chkNoLiqVendUff.Value = 0
    Me.chkAttivaQtaAbb.Value = 0
    Me.chkRicalcolaTutto.Value = 0
    TIPO_RAGGR_ABB = 0
    NO_RIP_SCARTI_IN_LIQ = 0
    NO_CALC_PRELIQ_NC = 0
    NO_CALC_PRELIQ_ND = 0
    DEC_IMP_UNI_LIQ = 0
    DEC_QTA_LIQ = 0
    
Else
    Me.cboTipoImportoArticolo.WriteOn fnNotNullN(rs!IDTipoImportoArticolo)
    Me.cboTipoImportoDocumento.WriteOn fnNotNullN(rs!IDTipoImportoDocumento)
    Me.cboTipoQuantita.WriteOn fnNotNullN(rs!IDTipoQuantita)
    Me.cboTipoLiquidazione.WriteOn fnNotNullN(rs!IDTipoLiquidazione)
    Me.Check1.Value = Abs(fnNotNullN(rs!ArticoliDiQuadratura))
    Me.cboTipoLiqConf.WriteOn fnNotNullN(rs!IDRV_POTipoLiqConf)
    Me.chkCalcolaPMCamp.Value = fnNotNullN(rs!PrezziMediInCampionatura)
    Me.chkAggPMCamp.Value = fnNotNullN(rs!AggiornaPMNonBloccatiCamp)
    Me.chkCalcTrattCamp.Value = fnNotNullN(rs!CalcoloTrattenuteSuCamp)
    Me.chkPrzMedioIVGamma.Value = fnNotNullN(rs!CalcolaPrezzoMedioIVGamma)
    TIPO_CALCOLO_PREZZO_MEDIO = fnNotNullN(rs!IDTipoPrezzoMedio)
    Me.chkCollegNotaCodiceLotto.Value = fnNotNullN(rs!CollegamentoNotaPerCodiceLotto)
    Me.chkAggCostoConfezInPrezzoLiq.Value = fnNotNullN(rs!AggiungiCostoConfezInPrezzoLiq)
    Me.ChkAggCostoKitInPrezzoLiq = fnNotNullN(rs!AggiungiCostoKitInPrezzoLiq)
    Me.chkNonCalcComm.Value = fnNotNullN(rs!NonCalcolareCommissioniInNDeNC)
    Me.chkNonCalcIncidenzaImb.Value = fnNotNullN(rs!NonCalcolareIncidenzaImballoInNDeNC)
    Me.chkNoLiqVendUff.Value = fnNotNullN(rs!NonLiquidareVenditaLiqUff)
    Me.chkAttivaQtaAbb.Value = fnNotNullN(rs!AttivaCalcoloQtaDaAbbattere)
    Me.chkRicalcolaTutto.Value = fnNotNullN(rs!RicalcolaTutto)
    TIPO_RAGGR_ABB = fnNotNullN(rs!IDRV_POTipoRaggruppamentoPerAbb)
    NO_RIP_SCARTI_IN_LIQ = fnNotNullN(rs!NonRipScartiInFatt)
    Me.chkNonCalcPreLiqInNC.Value = fnNotNullN(rs!NonCalcolarePreLiqInNC)
    Me.chkNonCalcPreLiqInND.Value = fnNotNullN(rs!NonCalcolarePreLiqInND)
    Me.txtDecImpUni.Value = fnNotNullN(rs!DecimaliPerCalcoloImpUniLiq)
    Me.txtDecQtaLiq.Value = fnNotNullN(rs!DecimaliPerCalcoloQtaLiq)
End If

rs.CloseResultset
Set rs = Nothing

If Me.cboTipoLiqConf.CurrentID = 2 Then
    Me.chkConguaglio.Enabled = True
Else
    Me.chkConguaglio.Enabled = False
End If


LINK_TIPO_AUMENTO_PESO = 0

sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    LINK_TIPO_AUMENTO_PESO = fnNotNullN(rs!IDTipoAumentoPeso)
End If

rs.CloseResultset
Set rs = Nothing


End Sub
Private Sub fncTipoImportoArticolo()
    With Me.cboTipoImportoArticolo
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoImportoArticolo"
        .DisplayField = "TipoImportoArticolo"
        .Sql = "SELECT * FROM RV_POTipoImportoArticolo WHERE IDRV_POTipoImportoArticolo<>2"
        .Fill
    End With

    With Me.cboTipoImportoArticolo_PerModifica
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoImportoArticolo"
        .DisplayField = "TipoImportoArticolo"
        .Sql = "SELECT * FROM RV_POTipoImportoArticolo WHERE IDRV_POTipoImportoArticolo<>2"
        .Fill
    End With
End Sub
Private Sub fncTipoImportoDocumento()
    With Me.cboTipoImportoDocumento
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoImportoDocumento"
        .DisplayField = "TipoImportoDocumento"
        .Sql = "SELECT * FROM RV_POTipoImportoDocumento"
        .Fill
    End With
  
    With Me.cboTipoImportoDocumento_PerModifica
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoImportoDocumento"
        .DisplayField = "TipoImportoDocumento"
        .Sql = "SELECT * FROM RV_POTipoImportoDocumento"
        .Fill
    End With
  
End Sub

Private Sub fncTipoQuantita()
    With Me.cboTipoQuantita
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoQuantita"
        .DisplayField = "TipoQuantita"
        .Sql = "SELECT * FROM RV_POTipoQuantita"
        .Fill
    End With
  
    With Me.cboTipoQuantita_PerModifica
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoQuantita"
        .DisplayField = "TipoQuantita"
        .Sql = "SELECT * FROM RV_POTipoQuantita"
        .Fill
    End With
  
End Sub

Private Sub fncTipoLiquidazione()
    With Me.cboTipoLiquidazione
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoLiquidazione"
        .DisplayField = "TipoLiquidazione"
        .Sql = "SELECT * FROM RV_POTipoLiquidazione"
        .Fill
    End With

    With Me.cboTipoLiquidazione_PerModifica
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoLiquidazione"
        .DisplayField = "TipoLiquidazione"
        .Sql = "SELECT * FROM RV_POTipoLiquidazione"
        .Fill
    End With
    
    'Tipo liquidazione conferimento
    With Me.cboTipoLiqConf
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POTipoLiqConf"
        .DisplayField = "TipoLiqConf"
        .Sql = "SELECT * FROM RV_POTipoLiqConf"
        .Fill
    End With
    

End Sub
Private Sub fncArticolo()

    With Me.CDArticolo
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hWnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
End Sub
Private Sub fncSocio()

    'Socio
    With Me.ACS
        'Imposta la connessione attiva al controllo
        Set .Connection = CnDMT
        'Imposta il nome dell'applicazione
        'Set .ApplicationName = TheApp
        'Imposta il nome dell'eseguibile dell'applicazione
        .Client = App.EXEName
        'Imposta l'identificativo dell'azienda corrente
        .IDFirm = TheApp.IDFirm
        'Imposta l'identificativo dell'utente corrente
        .IDUser = TheApp.IDUser
        .UserName = "Amministratore"
        'Impostare con la proprietà Hwnd del form che contiene
        'il controllo. Serve per l'esegui gestione
        .HwndContainer = Me.hWnd
    End With

    'Socio
    With Me.ACS_PerModifica
        'Imposta la connessione attiva al controllo
        Set .Connection = CnDMT
        'Imposta il nome dell'applicazione
        '.ApplicationName = m_App.
        'Imposta il nome dell'eseguibile dell'applicazione
        .Client = App.EXEName
        'Imposta l'identificativo dell'azienda corrente
        .IDFirm = TheApp.IDFirm
        'Imposta l'identificativo dell'utente corrente
        .IDUser = TheApp.IDUser
        .UserName = "Amministratore"
        'Impostare con la proprietà Hwnd del form che contiene
        'il controllo. Serve per l'esegui gestione
        .HwndContainer = Me.hWnd
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Me.cmdAvanti.Value = True Then
        FrmVisualizzaLiquidazione.Show
        Exit Sub
    End If
    
    If Me.cmdIndietro.Value = True Then
        FrmMain.Show
        Exit Sub
    End If
    
    If Me.cmdEliminaLiquidazione.Value = True Then
        FrmMain.Show
        Exit Sub
    End If
    

End Sub
Private Function fncNumeroLiquidazione() As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT MAX(NumeroLiquidazione) AS NumeroLiquidazione "
sSQL = sSQL & "FROM RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection
If rs.EOF = True Then
    fncNumeroLiquidazione = 1
Else
    fncNumeroLiquidazione = fnNotNullN(rs!NumeroLiquidazione) + 1
End If

rs.Close
Set rs = Nothing
End Function
Private Function fncNumeroProtocolloInterno() As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT MAX(NumeroProtInt) AS NumeroProtocolloInterno "
sSQL = sSQL & "From RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch


Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection
If rs.EOF = True Then
    fncNumeroProtocolloInterno = 1
Else
    fncNumeroProtocolloInterno = fnNotNullN(rs!NumeroProtocolloInterno) + 1
End If

rs.Close
Set rs = Nothing
End Function

Private Function ControlloEsistenzaNumeroLiquidazione() As Boolean
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT NumeroLiquidazione "
sSQL = sSQL & "From RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE NumeroLiquidazione=" & Me.txtNumeroLiquidazione.Value
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection

If rs.EOF = True Then
    ControlloEsistenzaNumeroLiquidazione = False
Else
    ControlloEsistenzaNumeroLiquidazione = True
End If

rs.Close
Set rs = Nothing

End Function
Private Function Permesso() As String
Permesso = ""
Dim IDCategoriaAna As Long
Dim TESTO As String

If Nuova_Liquidazione = 1 Then
    If ControlloEsistenzaNumeroLiquidazione = True Then
        Permesso = "Il numero di liquidazione è esistente"
        Exit Function
    End If
    
    If Me.txtDataInizio.Text = "" Then
        Permesso = "Inserire la data di inizio elaborazione del documento di liquidazione" & vbCrLf
        
        Exit Function
    End If
    If Me.txtDataFine.Text = "" Then
        Permesso = "Inserire la data di fine elaborazione del documento di liquidazione" & vbCrLf
        
        Exit Function
    End If
    If DateDiff("d", Me.txtDataInizio.Text, Me.txtDataFine.Text) < 0 Then
        Permesso = "Risulta una incongruenza tra le date di eleborazione" & vbCrLf
        
        Exit Function
    End If
    
    If Me.cboTipoImportoArticolo.CurrentID = 0 Then
        Permesso = "Manca il tipo di importo dell'articolo da utilizzare per il calcolo delle trattenute"
        Exit Function
    End If
    
    If Me.cboTipoImportoDocumento.CurrentID = 0 Then
        Permesso = "Manca il tipo di importo totale del documento da utilizzare per il calcolo delle trattenute"
        Exit Function
    End If
    
    If Me.cboTipoQuantita.CurrentID = 0 Then
        Permesso = "Manca il tipo di quantita da utilizzare per il calcolo delle trattenute"
        Exit Function
    End If
    
'    If Me.chkLiquidazionePerFor.Value = vbChecked Then
'        If Me.ACS.IDAnagrafica = 0 Then
'            Permesso = "Manca il fornitore"
'            Exit Function
'        End If
'    End If
    If Me.chkLiquidazionePerFor.Value = vbChecked Then
        IDCategoriaAna = GET_CATEGORIA_ANA_SOCIO(Me.ACS.IDAnagrafica)
        If IDCategoriaAna = LINK_TIPO_CATEGORIA_SOCIO Then
            TESTO = "ATTENZIONE!!!" & vbCrLf
            TESTO = TESTO & "È stato selezionato un socio e vistato 'Liquida un fornitore'." & vbCrLf
            TESTO = TESTO & "Questa configurazione potrebbe comportare un calcolo errato della liquidazione." & vbCrLf
            TESTO = TESTO & "Vuoi continuare?" & vbCrLf
            
            If MsgBox(TESTO, vbQuestion + vbYesNo, "Controllo parametri") = vbNo Then
                Permesso = TESTO
                Exit Function
            End If
        End If
    End If
    
    
End If
If Nuova_Liquidazione = 0 Then
    If Me.cboPeriodo.CurrentID = 0 Then
        Permesso = "Inserire il periodo di elaborazione del documento di liquidazione" & vbCrLf
        Exit Function
    End If
    
    If Me.txtDataInizio_PerModifica.Text = "" Then
        Permesso = "Inserire la data di inizio elaborazione del documento di liquidazione" & vbCrLf
        Exit Function
    End If
    
    If Me.txtDataFinePerModifica.Text = "" Then
        Permesso = "Inserire la data di fine elaborazione del documento di liquidazione" & vbCrLf
        Exit Function
    End If
    
    If DateDiff("d", Me.txtDataInizio_PerModifica.Text, Me.txtDataFinePerModifica.Text) < 0 Then
        Permesso = "Risulta una incongruenza tra le date di eleborazione" & vbCrLf
        Exit Function
    End If
    
    If Me.cboTipoImportoArticolo_PerModifica.CurrentID = 0 Then
        Permesso = "Manca il tipo di importo dell'articolo da utilizzare per il calcolo delle trattenute"
        Exit Function
    End If
    
    If Me.cboTipoImportoDocumento_PerModifica.CurrentID = 0 Then
        Permesso = "Manca il tipo di importo totale del documento da utilizzare per il calcolo delle trattenute"
        Exit Function
    End If
    
    If Me.cboTipoQuantita_PerModifica.CurrentID = 0 Then
        Permesso = "Manca il tipo di quantita da utilizzare per il calcolo delle trattenute"
        Exit Function
    End If
    
    If Me.cboTipoLiquidazione_PerModifica.CurrentID = 0 Then
        Permesso = "Manca il tipo di liquidazione da utilizzare per il calcolo delle trattenute"
        Exit Function
    End If

    
End If

End Function
Private Function GET_CATEGORIA_ANA_SOCIO(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica, IDCategoriaAnagrafica "
sSQL = sSQL & "FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CATEGORIA_ANA_SOCIO = fnNotNullN(rs!IDCategoriaAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Sub fncPeriodo()
    With Me.cboPeriodo
        Set .Database = CnDMT
        .AddFieldKey "IDRV_POLiquidazionePeriodo"
        .DisplayField = "Periodo"
        .Sql = "SELECT * FROM RV_POLiquidazionePeriodo ORDER BY NumeroLiquidazione DESC"
        .Fill
    End With
End Sub
Private Sub fncListino()
    
    With cboListino
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .Sql = "SELECT IDListino, Listino FROM Listino"
        .Sql = .Sql & " WHERE IDAzienda = " & TheApp.IDFirm
        .Sql = .Sql & " AND TipoListino = 0"
        .Sql = .Sql & " ORDER BY Listino"
    End With

End Sub
Private Sub fncCatMerc()
    
    With cboCategoriaMerc
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POCategoriaLiquidazione"
        .DisplayField = "CategoriaLiquidazione"
        .Sql = "SELECT IDRV_POCategoriaLiquidazione, CategoriaLiquidazione FROM RV_POCategoriaLiquidazione "
        .Sql = .Sql & " ORDER BY CategoriaLiquidazione"
    End With

End Sub

