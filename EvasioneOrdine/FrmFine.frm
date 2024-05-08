VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{FCA49525-5F72-11D2-B9EB-00201880103B}#18.1#0"; "DMTPrinterDialog.OCX"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmFine 
   Caption         =   "Fatturazione ordini"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmFine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10125
   ScaleWidth      =   21150
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   47
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   10095
      Left            =   0
      ScaleHeight     =   10065
      ScaleWidth      =   20985
      TabIndex        =   0
      Top             =   0
      Width           =   21015
      Begin VB.Frame Frame4 
         Caption         =   "Annotazioni documento"
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
         Height          =   3975
         Left            =   17400
         TabIndex        =   78
         Top             =   5520
         Width           =   3495
         Begin VB.CommandButton cmdNota3 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   2960
            TabIndex        =   88
            Top             =   2680
            Width           =   420
         End
         Begin VB.CommandButton cmdNota2 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   2960
            TabIndex        =   87
            Top             =   1480
            Width           =   420
         End
         Begin VB.CommandButton cmdNote1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   2960
            TabIndex        =   86
            Top             =   280
            Width           =   420
         End
         Begin VB.TextBox txtAnnotazione03 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Top             =   3000
            Width           =   3255
         End
         Begin VB.TextBox txtAnnotazione02 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   80
            Top             =   1800
            Width           =   3255
         End
         Begin VB.TextBox txtAnnotazione01 
            Height          =   855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   79
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label Label4 
            Caption         =   "Annotazione 3"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   85
            Top             =   2760
            Width           =   3135
         End
         Begin VB.Label Label4 
            Caption         =   "Annotazione 2"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   84
            Top             =   1560
            Width           =   3135
         End
         Begin VB.Label Label4 
            Caption         =   "Annotazione 1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraOrdineSel 
         Caption         =   "PARAMETRI ORDINE SELEZIONATO"
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
         Height          =   6495
         Left            =   8760
         TabIndex        =   57
         Top             =   3480
         Width           =   4695
         Begin VB.CommandButton cmdConfModOrdSel 
            Caption         =   "CONFERMA MODIFICHE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   6000
            Width           =   4455
         End
         Begin VB.TextBox txtIstruzioniMittenteOrdSel 
            Height          =   1035
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Top             =   2880
            Width           =   4455
         End
         Begin VB.TextBox txtTargaAutomezzoOrdSel 
            Height          =   315
            Left            =   120
            TabIndex        =   61
            Top             =   2280
            Width           =   4455
         End
         Begin VB.TextBox txtCausaleDocOrdSel 
            Height          =   315
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   4455
         End
         Begin DMTDataCmb.DMTCombo cboVettoreOrdSel 
            Height          =   315
            Left            =   120
            TabIndex        =   60
            Top             =   1680
            Width           =   4455
            _ExtentX        =   7858
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
         Begin DMTDataCmb.DMTCombo cboTipoTrasportoOrdSel 
            Height          =   315
            Left            =   120
            TabIndex        =   59
            Top             =   1050
            Width           =   4455
            _ExtentX        =   7858
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
         Begin DMTDataCmb.DMTCombo cboVettoreSuccOrdSel 
            Height          =   315
            Left            =   120
            TabIndex        =   63
            Top             =   4200
            Width           =   4455
            _ExtentX        =   7858
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
         Begin DMTDataCmb.DMTCombo cboLuogoPresaMerceOrdSel 
            Height          =   315
            Left            =   120
            TabIndex        =   64
            Top             =   4800
            Width           =   2175
            _ExtentX        =   3836
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
         Begin DMTDATETIMELib.dmtTime txtOraArrivoLuogoOrdSel 
            Height          =   315
            Left            =   3840
            TabIndex        =   66
            Top             =   4800
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDataArrivoLuogoOrdSel 
            Height          =   315
            Left            =   2400
            TabIndex        =   65
            Top             =   4800
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboAspettoEsterioreOrdSel 
            Height          =   315
            Left            =   120
            TabIndex        =   67
            Top             =   5400
            Width           =   4455
            _ExtentX        =   7858
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
            Caption         =   "Aspetto esteriore"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   77
            Top             =   5160
            Width           =   2415
         End
         Begin VB.Label Label1 
            Caption         =   "Data e ora di arrivo"
            Height          =   255
            Index           =   16
            Left            =   2400
            TabIndex        =   76
            Top             =   4560
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Luogo di presa merce"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   75
            Top             =   4560
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Vettore successivo"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   74
            Top             =   3960
            Width           =   3495
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo trasporto"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Istruzioni del mittente"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   72
            Top             =   2640
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Targa automezzo"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   71
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Vettore"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   70
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Causale del documento"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.ListBox lstReport 
         Appearance      =   0  'Flat
         Height          =   5205
         Left            =   17400
         Style           =   1  'Checkbox
         TabIndex        =   50
         Top             =   240
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Inserimento parametri di fatturazione"
         ForeColor       =   &H00FF0000&
         Height          =   8535
         Left            =   13560
         TabIndex        =   1
         Top             =   120
         Width           =   3735
         Begin VB.TextBox txtCausaleDocumento 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   2400
            Width           =   3495
         End
         Begin VB.CommandButton cmdIstruzioni 
            Height          =   260
            Left            =   3300
            Picture         =   "FrmFine.frx":1084A
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Apri dati C.M.R."
            Top             =   6360
            Width           =   300
         End
         Begin VB.TextBox txtTargaAutomezzo 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   6000
            Width           =   3495
         End
         Begin VB.TextBox txtIstruzioniMittente 
            Height          =   315
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   6600
            Width           =   3495
         End
         Begin VB.CheckBox chkCMR 
            Alignment       =   1  'Right Justify
            Caption         =   "C.M.R."
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
            Left            =   120
            TabIndex        =   2
            Top             =   6960
            Width           =   975
         End
         Begin DMTEDITNUMLib.dmtNumber txtNumeroDocumento 
            Height          =   315
            Left            =   2040
            TabIndex        =   3
            Top             =   1800
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo CboTipoDocumento 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   600
            Width           =   3495
            _ExtentX        =   6165
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
         Begin DMTDataCmb.DMTCombo CboPagamento 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   4200
            Width           =   3495
            _ExtentX        =   6165
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
         Begin DMTDATETIMELib.dmtDate txtDataDoc 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo CboValuta 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   3000
            Width           =   3495
            _ExtentX        =   6165
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
         Begin DMTDataCmb.DMTCombo CboSezionale 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   3495
            _ExtentX        =   6165
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
         Begin DMTDataCmb.DMTCombo cboMagazzino 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   3600
            Width           =   3495
            _ExtentX        =   6165
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
         Begin DMTDataCmb.DMTCombo cboVettore 
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   5400
            Width           =   3495
            _ExtentX        =   6165
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
         Begin DMTDataCmb.DMTCombo cboSezionaleCMR 
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   7515
            Width           =   3495
            _ExtentX        =   6165
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
         Begin DMTDATETIMELib.dmtDate txtDataDocCMR 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   8130
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTEDITNUMLib.dmtNumber txtNumeroDocCMR 
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Top             =   8130
            Width           =   1770
            _Version        =   65536
            _ExtentX        =   3122
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboTipoTrasporto 
            Height          =   315
            Left            =   120
            TabIndex        =   51
            Top             =   4770
            Width           =   3495
            _ExtentX        =   6165
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
         Begin VB.Label Label1 
            Caption         =   "Causale del documento"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   56
            Top             =   2160
            Width           =   3495
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo trasporto"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   52
            Top             =   4560
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo di pagamento"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   3960
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Data documento"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Valuta azienda"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Sezionale"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Magazzino"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   3360
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Vettore"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   5160
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo documento"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "Targa automezzo"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   21
            Top             =   5760
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Istruzioni del mittente"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   20
            Top             =   6360
            Width           =   3495
         End
         Begin VB.Label Label1 
            Caption         =   "N° Documento"
            Height          =   255
            Index           =   7
            Left            =   2040
            TabIndex        =   19
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lblDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero"
            Height          =   195
            Index           =   40
            Left            =   1800
            TabIndex        =   18
            Top             =   7920
            Width           =   1395
         End
         Begin VB.Label lblDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Doc."
            Height          =   195
            Index           =   41
            Left            =   120
            TabIndex        =   17
            Top             =   7920
            Width           =   1080
         End
         Begin VB.Label lblDocument 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sezionale"
            Height          =   195
            Index           =   42
            Left            =   120
            TabIndex        =   16
            Top             =   7320
            Width           =   2235
         End
      End
      Begin VB.CommandButton CmdAvvia 
         Caption         =   "Fine"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   19320
         TabIndex        =   40
         Top             =   9600
         Width           =   1575
      End
      Begin VB.CommandButton CmdFine 
         Caption         =   "Annulla"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   17280
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1575
      End
      Begin VB.CommandButton CmdIndietro 
         Caption         =   "Indietro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   13560
         TabIndex        =   32
         Top             =   8640
         Width           =   3735
         Begin VB.CheckBox chkStampaNumeroCopie 
            Caption         =   "Stampa numero copie"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            ToolTipText     =   $"FrmFine.frx":10DD4
            Top             =   480
            Width           =   2775
         End
         Begin VB.CheckBox chkStampaFattura 
            Caption         =   "Stampa documento"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   2415
         End
         Begin DMTEDITNUMLib.dmtNumber txtNumeroCopie 
            Height          =   255
            Left            =   3000
            TabIndex        =   34
            Top             =   480
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   450
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            DecimalPlaces   =   0
            AllowEmpty      =   0   'False
         End
      End
      Begin VB.CheckBox chkChiudiOrdini 
         Caption         =   "Chiudi ordini elaborati alla fine del processo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   9000
         Width           =   8535
      End
      Begin VB.Frame Frame3 
         Caption         =   "ANNOTAZIONI INTERNE"
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
         Height          =   2535
         Left            =   8760
         TabIndex        =   29
         Top             =   960
         Width           =   4695
         Begin VB.TextBox txtAnnotazioniOrdine 
            Appearance      =   0  'Flat
            ForeColor       =   &H00C00000&
            Height          =   2175
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   4455
         End
      End
      Begin DMTLblLinkCtl.LabelLink LabelLink1 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   8520
         Visible         =   0   'False
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   450
         Caption         =   "Visualizza e modifica il [TipoOggetto] [Numero documento] del [data documento]"
         Name            =   "LabelLink"
      End
      Begin MSComctlLib.ListView LVFattureCreate 
         Height          =   1815
         Left            =   120
         TabIndex        =   36
         Top             =   6360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin DmtGridCtl.DmtGrid GrigliaOrdini 
         Height          =   4815
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8493
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
      Begin DMTPrinterDialog.DMTDialog DMTDialog 
         Left            =   120
         Top             =   6480
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   9360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   4935
         Left            =   17400
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   8705
         BackColor       =   -2147483643
         ForeColor       =   -2147483630
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ORDINI DA EVADERE"
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
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   8535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "CHIUDERE GLI ORDINI INTERESSATI ALLA FATTURAZIONE NELLE ALTRE POSTAZIONI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   13335
      End
      Begin VB.Label LblFine 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   9720
         Width           =   8535
      End
      Begin VB.Label lblInfoStatus 
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   8160
         Width           =   8535
      End
      Begin VB.Line Line1 
         DrawMode        =   1  'Blackness
         X1              =   8640
         X2              =   120
         Y1              =   8880
         Y2              =   8880
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DOCUMENTI CREATI"
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
         Left            =   120
         TabIndex        =   42
         Top             =   6120
         Width           =   8535
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Annotazione 1"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "FrmFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




''''CARTELLA DATA APPLICAZIONI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Const CSIDL_COMMON_APPDATA = &H23
Private Const CSIDL_LOCAL_APPDATA = &H1C&
Private Const MAX_PATH = 260
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''Numero documento della fattura per conto del socio'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private NumeroDocumento As Long
Private Link_TipoOggetto As Long
Private Link_TipoOggetto_Coop As Long
Private StringaRiferimento As String

Private ObjDoc As DmtDocs.cDocument
Private cDefault As Collection
Private oReport As dmtReportLib.dmtReport
'Private cPerfoming As New CPerforming
Private Const NOMETABELLAPIANA = "ValoriOggettoPerTipo"
Private Const NomeTabellaDettaglio = "ValoriOggettoDettaglio"

Private ArrayCli(0, 12) As String

Private NumeroRecord As Long

Private rsGrigliaOrdini As ADODB.Recordset

Private Link_IVAArticolo As Long
Private AliquotaIvaArticolo As Double

Private Unita_Progresso As Double

'Protocollo ICE
Private Link_Prog_Protocollo_ICE As Long
Private NumeroProgressivoICE As Long


Private TIPO_SCONTO_CLIENTE As Long
Private SCONTO_ABITUALE As Double

Private FLAG_IVA_UGUALE As Long
Private FLAG_IVA_IMBALLO_A_RENDERE As Long


Private FLAG_INTRASTAT_DOC As Long

'----- Oggetti e variabili per la gestione del riquadro attività -----------
'***Reports                                                                -
Private WithEvents oReportsActivity As DmtActBoxLib.ReportsActivity       '-
Attribute oReportsActivity.VB_VarHelpID = -1
'***Filtri                                                                 -
Private WithEvents oFiltersActivity As DmtActBoxLib.FiltersActivity       '-
Attribute oFiltersActivity.VB_VarHelpID = -1
'***Viste tabellari                                                        -
Private WithEvents oTableViewsActivity As DmtActBoxLib.TableViewsActivity '-
Attribute oTableViewsActivity.VB_VarHelpID = -1
'***Esportazioni                                                           -
Private oExportActivity As DmtActBoxLib.ExportActivity                    '-
'***Supporto tecnico                                                       -
Private oSupportActivity As DmtActBoxLib.SupportActivity                  '-
'***Nome dell'attività predefinita del riquadro attività                   -
Private m_DefaultActivity As String                                       '-
'---------------------------------------------------------------------------

Public TotaleNumeroColliOrdine As Double
Public TotalePesoOrdine As Double


'Variabile utilizzata per ottenere il nome della tabella di testata del documento
Private sTabellaTestata As String
'Variabile utilizzata per ottenere il nome della tabella di dettaglio del documento
Private sTabellaDettaglio As String
'Variabile utilizzata per ottenere il nome della tabella delle scadenze del documento
Private sTabellaScadenze As String
'Variabile utilizzata per ottenere il nome della tabella del castelletto IVA del documento
Private sTabellaIVA As String

Private Sub GET_NUMERO_ZERI_DOC_RIF()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT NumeroZeriRifDoc FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    NUMERO_ZERI_DOC_RIF = fnNotNullN(rs!NumeroZeriRifDoc)
Else
    NUMERO_ZERI_DOC_RIF = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub GET_DEFAULT_AZIENDA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoOggettoDocEvasione, IDPagamentoDocDefault "
sSQL = sSQL & "FROM PersonalizzazionePerFiliale "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    'Me.CboPagamento.WriteOn fnNotNullN(rs!IDPagamentoDocDefault)
    Me.CboTipoDocumento.WriteOn fnNotNullN(rs!IDTipoOggettoDocEvasione)
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_MAGAZZINO_DEFAULT()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDMagazzino_Vendita "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboMagazzino.WriteOn 0
Else
    Me.cboMagazzino.WriteOn fnNotNullN(rs!IDMagazzino_Vendita)

End If


rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub GET_VETTORE_DEFAULT()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


If Me.cboTipoTrasporto.CurrentID = 3 Then
    sSQL = "SELECT IDVettore "
    sSQL = sSQL & "FROM ConfigurazioneVendite "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        Me.cboVettore.WriteOn 0
    Else
        Me.cboVettore.WriteOn fnNotNullN(rs!IDVettore)
    End If

    rs.CloseResultset
    Set rs = Nothing
    
    Me.cboVettore.Enabled = True
    
Else
    Me.cboVettore.WriteOn 0
    Me.cboVettore.Enabled = False
End If


End Sub
Private Function GET_DEFAULT_VETTORE_X_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDVettoreDefault "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DEFAULT_VETTORE_X_CLIENTE = 0
Else
    GET_DEFAULT_VETTORE_X_CLIENTE = fnNotNullN(rs!IDVettoreDefault)

End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub GET_TIPO_OGGETTO_PER_FATT(IDOggetto_local As Long)

Select Case IDOggetto_local

    Case 2
        Link_TipoOggetto_Coop = GET_TIPO_OGGETTO("RV_PODDTL")
        StringaRiferimento = "Rif. D.d.t."
    Case 114
        Link_TipoOggetto_Coop = GET_TIPO_OGGETTO("RV_POFAL")
        StringaRiferimento = "Rif. F.A."
    Case 8
        Link_TipoOggetto_Coop = GET_TIPO_OGGETTO("RV_POSNFL")
        StringaRiferimento = "Rif. S.N.F."
    Case Else
        Link_TipoOggetto_Coop = 0
        StringaRiferimento = ""
End Select

End Sub

Private Function GET_PeriodoIVA(Anno As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDPeriodoIVA FROM PeriodoIVA "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Anno=" & Anno


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PeriodoIVA = 0
Else
    GET_PeriodoIVA = fnNotNullN(rs!IDPeriodoIva)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Sub ActivityBox_ItemSelected(ByVal Item As DmtActBoxTlb.Item, NeedRedraw As Boolean)

Me.txtNumeroCopie.Value = GET_NUMERO_COPIE(oReportsActivity.SelectedReportName, Link_TipoOggetto_Coop)


End Sub
Private Function GET_NUMERO_COPIE(Report As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_GET_NUMERO_COPIE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT NumeroCopie FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND ReportTipoOggetto=" & fnNormString(Report)


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_COPIE = 1
Else
    If fnNotNullN(rs!NumeroCopie) = 0 Then
        GET_NUMERO_COPIE = 1
    Else
        GET_NUMERO_COPIE = fnNotNullN(rs!NumeroCopie)
    End If
End If


rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_NUMERO_COPIE:
    GET_NUMERO_COPIE = 1
End Function
Private Sub cboSezionale_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_PERIODO_IVA  As Long

If Me.txtDataDoc.Value = 0 Then Me.txtDataDoc.Value = Date

LINK_PERIODO_IVA = GET_PeriodoIVA(DatePart("yyyy", Me.txtDataDoc.Text))

sSQL = "SELECT ProgressivoDisponibile FROM ProgressivoSezionale "
sSQL = sSQL & "WHERE IDSezionale=" & Me.cboSezionale.CurrentID
sSQL = sSQL & " AND IDTipoModulo=1"
sSQL = sSQL & " AND IDPeriodoIVA=" & LINK_PERIODO_IVA

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    sSQL = "INSERT INTO ProgressivoSezionale ("
    sSQL = sSQL & "IDProgressivoSezionale, IDTipoModulo, IDPeriodoIVA, IDSezionale, "
    sSQL = sSQL & "ProgressivoDisponibile, DataUltimaVariazione, IDUtenteUltimaVariazione, "
    sSQL = sSQL & "VirtualDelete) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("ProgressivoSezionale", "IDProgressivoSezionale") & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & LINK_PERIODO_IVA & ", "
    sSQL = sSQL & Me.cboSezionale.CurrentID & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & TheApp.IDUser & ", "
    sSQL = sSQL & 0 & ")"
    
    CnDMT.Execute sSQL
    
    Me.txtNumeroDocumento.Value = 1
Else
    Me.txtNumeroDocumento.Value = fnNotNullN(rs!ProgressivoDisponibile)
End If


rs.CloseResultset
Set rs = Nothing

Me.chkCMR.Value = GET_SEZ_PER_CMR(Me.cboSezionale.CurrentID)

End Sub

Private Sub cboSezionaleCMR_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_PERIODO_IVA  As Long

If Me.txtDataDocCMR.Value = 0 Then Me.txtDataDocCMR.Value = Date

LINK_PERIODO_IVA = GET_PeriodoIVA(DatePart("yyyy", Me.txtDataDocCMR.Text))

sSQL = "SELECT ProgressivoDisponibile FROM ProgressivoSezionale "
sSQL = sSQL & "WHERE IDSezionale=" & Me.cboSezionaleCMR.CurrentID
sSQL = sSQL & " AND IDTipoModulo=1"
sSQL = sSQL & " AND IDPeriodoIVA=" & LINK_PERIODO_IVA

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    sSQL = "INSERT INTO ProgressivoSezionale ("
    sSQL = sSQL & "IDProgressivoSezionale, IDTipoModulo, IDPeriodoIVA, IDSezionale, "
    sSQL = sSQL & "ProgressivoDisponibile, DataUltimaVariazione, IDUtenteUltimaVariazione, "
    sSQL = sSQL & "VirtualDelete) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("ProgressivoSezionale", "IDProgressivoSezionale") & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & LINK_PERIODO_IVA & ", "
    sSQL = sSQL & Me.cboSezionaleCMR.CurrentID & ", "
    sSQL = sSQL & 1 & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & TheApp.IDUser & ", "
    sSQL = sSQL & 0 & ")"
    
    CnDMT.Execute sSQL
    
    Me.txtNumeroDocCMR.Value = 1
Else
    Me.txtNumeroDocCMR.Value = fnNotNullN(rs!ProgressivoDisponibile)
End If


rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub CboTipoDocumento_Click()
On Error GoTo ERR_CboTipoDocumento_Click
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroOrdiniSel As Long
Dim IDAnagraficaSel As Long
Dim IDSezionaleCliente As Long

GET_TIPO_OGGETTO_PER_FATT Me.CboTipoDocumento.CurrentID

fncReport
DEFAULT_REPORT_LISTA

With Me.cboSezionale
    Set .Database = TheApp.Database.Connection
    .AddFieldKey "IDSezionale"
    .DisplayField = "Sezionale"
    .Sql = "SELECT Sezionale.IDSezionale, Sezionale.Sezionale "
    .Sql = .Sql & "FROM RegistroIvaPerTipoOggetto INNER JOIN "
    .Sql = .Sql & "Sezionale ON RegistroIvaPerTipoOggetto.IDRegistroIva = Sezionale.IDRegistroIva AND "
    .Sql = .Sql & "RegistroIvaPerTipoOggetto.IDFiliale = Sezionale.IDFiliale LEFT OUTER JOIN "
    .Sql = .Sql & "TipoOggetto ON RegistroIvaPerTipoOggetto.IDTipoOggetto = TipoOggetto.IDTipoOggetto "
    .Sql = .Sql & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & Me.CboTipoDocumento.CurrentID
    .Sql = .Sql & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & VarIDFiliale
End With


'Default SEZIONALE
sSQL = "SELECT IDSezionale "
sSQL = sSQL & "FROM DefaultFilialePerTipoOggetto "
sSQL = sSQL & "WHERE (IDTipoOggetto = " & Me.CboTipoDocumento.CurrentID & ") And (IDSezionale > 0) And (IDFiliale = " & TheApp.Branch & ")"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    Me.cboSezionale.WriteOn fnNotNullN(rs!IDSezionale)
Else
    Me.cboSezionale.WriteOn 0
End If

rs.CloseResultset
Set rs = Nothing


Me.txtAnnotazione01.Text = GET_NOTA_DOCUMENTO(Me.CboTipoDocumento.CurrentID, 1)
Me.txtAnnotazione02.Text = GET_NOTA_DOCUMENTO(Me.CboTipoDocumento.CurrentID, 2)
Me.txtAnnotazione03.Text = GET_NOTA_DOCUMENTO(Me.CboTipoDocumento.CurrentID, 3)

NumeroOrdiniSel = GET_CONTROLLO_NUM_ORD_DA_EVADERE
If NumeroOrdiniSel = 1 Then
    IDAnagraficaSel = GET_IDCLIENTE_DA_EVADERE
    If IDAnagraficaSel > 0 Then
        If Me.CboTipoDocumento.CurrentID > 0 Then
            IDSezionaleCliente = GET_SEZ_PER_CLIENTE(Me.CboTipoDocumento.CurrentID, IDAnagraficaSel)
            If IDSezionaleCliente > 0 Then Me.cboSezionale.WriteOn IDSezionaleCliente
        End If
    End If
End If
Exit Sub
ERR_CboTipoDocumento_Click:
    MsgBox Err.Description, vbCritical, "CboTipoDocumento_Click"
End Sub

Private Sub cboTipoTrasporto_Click()
    GET_VETTORE_DEFAULT
End Sub

Private Sub chkCMR_Click()
    If Me.chkCMR.Value = vbChecked Then
        Me.txtDataDocCMR.Value = Date
        Me.cboSezionaleCMR.WriteOn LINK_SEZIONALE_CMR
    Else
        Me.cboSezionaleCMR.WriteOn 0
        Me.txtDataDocCMR.Value = 0
        Me.txtNumeroDocCMR.Value = 0
    End If
End Sub
Private Sub CmdAvvia_Click()
'On Error GoTo ERR_CmdAvvia_Click
Dim sSQL As String
Dim rsSelectOrd As ADODB.Recordset
Dim rsEla As DmtOleDbLib.adoResultset
Dim I As Integer

Dim Cliente As String
Dim SitoPerAnagrafica As String
Dim Vettore As String
Dim IDSitoPerAnagrafica As Long
Dim IDVettore As Long
Dim ArrayOrd() As String
Dim RigaFile As String
Dim OLDCursor As Long
Dim Testo As String
Dim NomeCartellaALLUSER As String

If Link_TipoOggetto_Coop = 0 Then
    MsgBox "Il tipo di documento non è valido o non è stato selezionato", vbCritical, "Avvio procedura"
    Me.CboTipoDocumento.SetFocus
    Exit Sub
End If

If Me.CboTipoDocumento.CurrentID = 0 Then
    MsgBox "Inserire il tipo di documento", vbInformation, "Elaborazione dati"
    Me.CboTipoDocumento.SetFocus
    Exit Sub
End If
If Me.cboTipoTrasporto.CurrentID = 3 Then
    If Me.cboVettore.CurrentID = 0 Then
        MsgBox "Inserire il vettore", vbInformation, "Elaborazione dati"
        Me.cboVettore.SetFocus
        Exit Sub
    End If
End If


If Me.cboMagazzino.CurrentID = 0 Then
    MsgBox "Inserire il magazzino di riferimento", vbInformation, "Elaborazione dati"
    Me.cboMagazzino.SetFocus
    Exit Sub
End If

If Me.CboValuta.CurrentID = 0 Then
    MsgBox "Inserire La valuta del documento", vbInformation, "Elaborazione dati"
    Me.CboValuta.SetFocus
    Exit Sub
End If

GET_NUMERO_ZERI_DOC_RIF
GET_PESO_PEDANA_IN_VENDITA
ParametroICE

rsGrigliaOrdini.UpdateBatch


CnDMT.Execute "DELETE FROM RV_POTMPOrdiniFatturati WHERE IDUtente=" & TheApp.IDUser

fnGrigliaFattureCreate

Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100
NumeroRecord = 0

'*****************SELEZIONE DEL NUMERO DI RECORD*************************************************
sSQL = "SELECT COUNT(NumeroRiga) AS NumeroRecord "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE DaRegistrare=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rsSelectOrd = New ADODB.Recordset
rsSelectOrd.Open sSQL, CnDMT.InternalConnection
If Not rsSelectOrd.EOF Then
    NumeroRecord = fnNotNullN(rsSelectOrd!NumeroRecord)
End If
rsSelectOrd.Close
Set rsSelectOrd = Nothing

If NumeroRecord = 0 Then
    MsgBox "Non ci sono ordini da fatturare", vbInformation, "Elaborazioni ordini"
    Exit Sub
End If



AVVIA_FATTURAZIONE = 1
Me.CmdAvvia.Enabled = False
Me.CmdFine.Enabled = False
Me.CmdIndietro.Enabled = False

Unita_Progresso = Me.ProgressBar1.Max / NumeroRecord

rsGrigliaOrdini.Close
Set rsGrigliaOrdini = Nothing

'********************************************************
Me.CmdAvvia.Enabled = False

NomeCartellaALLUSER = TrovaCartella(CSIDL_LOCAL_APPDATA)

sSQL = "SELECT  IDCliente, IDSitoPerAnagrafica, IDoggetto, DaRegistrare, IDUtente, IDVettore, IDAzienda "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "GROUP BY IDCliente, IDSitoPerAnagrafica, DaRegistrare, IDOggetto, IDUtente, IDVettore, IDAzienda "
sSQL = sSQL & "HAVING DaRegistrare = " & fnNormBoolean(1)
sSQL = sSQL & " AND IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rsEla = CnDMT.OpenResultset(sSQL)
Open NomeCartellaALLUSER & "OrdiniDaFatturare" For Output As #1
While Not rsEla.EOF
    If GET_CONTROLLO_CLIENTE_BLOCCATO(fnNotNullN(rsEla!IDCliente)) = 0 Then
        Print #1, fnNotNullN(rsEla!IDCliente) & ";" & fnNotNullN(rsEla!IDSitoPerAnagrafica) & ";" & fnNotNullN(rsEla!IDOggetto) & ";" & fnNotNullN(rsEla!IDVettore)
    End If
rsEla.MoveNext
Wend
Close #1
rsEla.CloseResultset
Set rs = Nothing

Set ObjDoc = New DmtDocs.cDocument

OLDCursor = CnDMT.CursorLocation
CnDMT.CursorLocation = adUseClient

GET_PARAMETRI_IVA_IMBALLO

Open NomeCartellaALLUSER & "OrdiniDaFatturare" For Input As #1
    Do While Not EOF(1)
        Line Input #1, RigaFile
        ArrayOrd() = Split(RigaFile, ";")
             
        Testo = GET_CONTROLLO_STATO_ORDINE(CLng(ArrayOrd(2)))
        
        If Len(Testo) > 0 Then
            MsgBox GET_RIFERIMENTO_ORDINE(CLng(ArrayOrd(2))) & Testo & vbCrLf & "Impossibile creare documento di vendita", vbCritical, "Creazione documenti di vendita"
        Else
             
            Cliente = GET_CLIENTE(CLng(ArrayOrd(0)))
            
            TIPO_SCONTO_CLIENTE = GET_TIPO_SCONTO(CLng(ArrayOrd(0)))
            SCONTO_ABITUALE = GET_SCONTO_ABITUALE(CLng(ArrayOrd(0)))
            
            SitoPerAnagrafica = GET_CLIENTE_DESTINAZIONE(CLng(ArrayOrd(1)))
            Vettore = GET_VETTORE(CLng(ArrayOrd(3)))
            
            'GET_DOCUMENTO_DEFAULF_CLIENTE CLng(ArrayOrd(0))
            
            Me.lblInfoStatus.Caption = "ELABORAZIONI ORDINI DEL CLIENTE " & Cliente & " - " & SitoPerAnagrafica
            DoEvents
            
            Settaggio
            
            fncTestata fnNotNullN(ArrayOrd(0)), fnNotNullN(ArrayOrd(1)), fnNotNullN(ArrayOrd(2)), fnNotNullN(ArrayOrd(3))
            
            fncRighe fnNotNullN(ArrayOrd(0)), fnNotNullN(ArrayOrd(1)), fnNotNullN(ArrayOrd(2))
            
            Me.lblInfoStatus.Caption = "CREAZIONE DOCUMENTO DI VENDITA......"
            DoEvents
            
            InserimentoDMT fnNotNullN(ArrayOrd(0)), Cliente, fnNotNullN(ArrayOrd(1)), SitoPerAnagrafica, fnNotNullN(ArrayOrd(2))
            
            fnGrigliaFattureCreate
            DoEvents
            
            fnGrigliaOrdini
            DoEvents
        End If
    Loop
    
    Close #1

CnDMT.CursorLocation = OLDCursor

    'If Me.chkChiudiOrdini.Value = vbChecked Then
    '    Me.lblInfoStatus.Caption = "Chiusura ordini elaborati......."
    '    DoEvents
    '    fnChiudiOrdiniElaborati
    'End If

Me.lblInfoStatus.Caption = ""
Me.LblFine.Caption = "OPERAZIONE COMPLETATA"

fnGrigliaOrdini

AVVIA_FATTURAZIONE = 0
Me.CmdAvvia.Enabled = True
Me.CmdFine.Enabled = True
Me.CmdIndietro.Enabled = True
Exit Sub
ERR_CmdAvvia_Click:
    MsgBox Err.Description, vbCritical, "CmdAvvia_Click"
    AVVIA_FATTURAZIONE = 0
    Me.CmdAvvia.Enabled = True
    Me.CmdFine.Enabled = True
    Me.CmdIndietro.Enabled = True
End Sub

Private Sub cmdConfModOrdSel_Click()
On Error GoTo ERR_cmdConfModOrdSel_Click
Dim IDOggetto As Long

If ((rsGrigliaOrdini.EOF) And (rsGrigliaOrdini.BOF)) Then Exit Sub
IDOggetto = fnNotNullN(rsGrigliaOrdini!IDOggetto)
If (IDOggetto > 0) Then
    AGGIORNA_ORDINE_SELEZIONATO IDOggetto
End If
Exit Sub
ERR_cmdConfModOrdSel_Click:
    MsgBox Err.Description, vbCritical, "cmdConfModOrdSel_Click"
End Sub

Private Sub CmdFine_Click()
    If MsgBox("Vuoi abbandonare il wizard per la creazione della fattura?", vbQuestion + vbYesNo, "Creazione liquidazione") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdIndietro_Click()
    Unload Me
End Sub



Private Sub cmdIstruzioni_Click()
If Me.cboVettore.CurrentID > 0 Then
    If Me.txtIstruzioniMittente.Height = 285 Then
        Me.txtIstruzioniMittente.Height = 1845
        Me.txtIstruzioniMittente.SetFocus
    Else
        Me.txtIstruzioniMittente.Height = 285
        Me.txtIstruzioniMittente.SetFocus
    End If
Else
    Me.txtIstruzioniMittente.Height = 285
End If
    
End Sub



Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    Me.Caption = Me.Caption & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    Me.WindowState = 2
    
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
    

    Link_TipoOggetto = 0
    Link_TipoOggetto_Coop = 0
    
    
    fncSezionale
    fncSezionaleCMR
    fncPagamento
    fncValuta
    fncTipoOggetto
    fncMagazzino
    fncTipoTrasporto
    fncVettore
    ParametroAggiornaPrezzoMedioDaConf
    ParametroAggiornaTipoLavorazioneDaConf
    GET_LINK_PORTO_PER_COMM_TRASP
    GET_DEFAULT_AZIENDA
    GET_MAGAZZINO_DEFAULT
    GET_TIPO_CALCOLO_COMM_TRASP
    GET_STAMPA_DOC_ATT
    GET_PARAMETRI_LIQ TheApp.Branch
    RECUPERA_CONFIG_CAUS_XML
    
    'GET_VETTORE_DEFAULT
    
    Me.txtDataDoc.Text = Date
    
    Me.CboValuta.WriteOn 9
    
    fncReport
    
    'Me.cboReport.WriteOn fnDefaultReport
    DEFAULT_REPORT_LISTA
    
    fnGrigliaOrdini
    
    Me.chkChiudiOrdini.Value = vbChecked
    
    If STAMPA_DOCUMENTO_NON_ATTIVO = 0 Then
         Me.chkStampaFattura.Value = vbChecked
    Else
        Me.chkStampaFattura.Value = vbUnchecked
    End If
    
    Me.chkStampaNumeroCopie.Value = STAMPA_DOCUMENTO_ATTIVO
    
    Me.txtNumeroCopie.Value = 1
    
    
    'Inizializza la ListView contenente la ricerca
    With Me.LVFattureCreate
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "IDOggetto", 100
        .ColumnHeaders.Add , , "IDTipoDocumento", 100
        .ColumnHeaders.Add , , "Data", 1000
        .ColumnHeaders.Add , , "N° Doc", 1000
        .ColumnHeaders.Add , , "Cliente", 1500
        .ColumnHeaders.Add , , "Destinazione", 1500
        .ColumnHeaders.Add , , "Vettore", 1500
    End With

    'Inizializzazione della LabelLink
    '----------------------------
    Set LabelLink1.Application = TheApp    'L’oggetto Application
    LabelLink1.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    'Viene disabilitata la voce "Ricerca" del menu popup
    LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False
End Sub

Private Sub fncPagamento()
    Dim sSQL As String

    
    
    sSQL = "SELECT IDPagamento, Pagamento"
    sSQL = sSQL & " FROM Pagamento"
    sSQL = sSQL & " ORDER BY Pagamento"

    With Me.cboPagamento
        Set .Database = CnDMT
        .DisplayField = "Pagamento"
        .AddFieldKey "IDPagamento"
        .Sql = sSQL
        .Refresh
    End With

End Sub
Private Sub fncTipoTrasporto()
    Dim sSQL As String

    
    
    sSQL = "SELECT IDTipoSpedizione, TipoSpedizione"
    sSQL = sSQL & " FROM TipoSpedizione"
    sSQL = sSQL & " ORDER BY TipoSpedizione"

    With Me.cboTipoTrasporto
        Set .Database = CnDMT
        .DisplayField = "TipoSpedizione"
        .AddFieldKey "IDTipoSpedizione"
        .Sql = sSQL
        .Refresh
    End With

    With Me.cboTipoTrasportoOrdSel
        Set .Database = CnDMT
        .DisplayField = "TipoSpedizione"
        .AddFieldKey "IDTipoSpedizione"
        .Sql = sSQL
        .Refresh
    End With
End Sub

Private Sub fncValuta()
    Dim sSQL As String
    'Dim sSQLValuta As String
    'Dim rs As dmtoledblib.adoResultset
    
    
    sSQL = "SELECT IDValuta, Valuta"
    sSQL = sSQL & " FROM Valuta"
    'sSQL = sSQL & " WHERE ((IDFiliale=" & VarIDFiliale & ") AND (IDRegistroIva = 1))"
    sSQL = sSQL & " ORDER BY Valuta"

    With Me.CboValuta
        Set .Database = CnDMT
        .DisplayField = "Valuta"
        .AddFieldKey "IDValuta"
        .Sql = sSQL
        .Refresh
    End With
    
End Sub
Private Sub fncMagazzino()
    Dim sSQL As String
    'Dim sSQLValuta As String
    'Dim rs As dmtoledblib.adoResultset
    
    
    sSQL = "SELECT IDMagazzino, Magazzino"
    sSQL = sSQL & " FROM Magazzino"
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Magazzino"

    With Me.cboMagazzino
        Set .Database = CnDMT
        .DisplayField = "Magazzino"
        .AddFieldKey "IDMagazzino"
        .Sql = sSQL
        .Refresh
    End With
    
End Sub
Private Sub fncVettore()
    Dim sSQL As String
    
    sSQL = "SELECT IDVettore, Vettore"
    sSQL = sSQL & " FROM Vettore"
    sSQL = sSQL & " ORDER BY Vettore"

    With Me.cboVettore
        Set .Database = CnDMT
        .DisplayField = "Vettore"
        .AddFieldKey "IDVettore"
        .Sql = sSQL
        .Refresh
    End With
    
    With Me.cboVettoreOrdSel
        Set .Database = CnDMT
        .DisplayField = "Vettore"
        .AddFieldKey "IDVettore"
        .Sql = sSQL
        .Refresh
    End With
    
    With Me.cboVettoreSuccOrdSel
        Set .Database = CnDMT
        .DisplayField = "Vettore"
        .AddFieldKey "IDVettore"
        .Sql = sSQL
        .Refresh
    End With
    
    With Me.cboLuogoPresaMerceOrdSel
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT * FROM SitoPerAnagrafica  "
        .Sql = .Sql & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .Sql = .Sql & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With
    
    With Me.cboAspettoEsterioreOrdSel
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDAspettoEsterioreArticolo"
        .DisplayField = "AspettoEsterioreArticolo"
        .Sql = "SELECT * FROM AspettoEsterioreArticolo"
        .Sql = .Sql & " ORDER BY AspettoEsterioreArticolo"
    End With
    
End Sub
Private Sub fncSezionale()
    Dim sSQL As String
    
    sSQL = "SELECT IDSezionale, Sezionale "
    sSQL = sSQL & " FROM Sezionale "
    sSQL = sSQL & " WHERE ((IDFiliale=" & TheApp.Branch & ") AND (IDRegistroIva <> 2))"
    sSQL = sSQL & " ORDER BY Sezionale"

    With Me.cboSezionale
        Set .Database = CnDMT
        .DisplayField = "Sezionale"
        .AddFieldKey "IDSezionale"
        .Sql = sSQL
        .Refresh
    End With
    
End Sub
Private Sub fncSezionaleCMR()
    Dim sSQL As String
    
    sSQL = "SELECT IDSezionale, Sezionale "
    sSQL = sSQL & " FROM Sezionale "
    sSQL = sSQL & " WHERE ((IDFiliale=" & TheApp.Branch & ") AND (IDRegistroIva <> 2))"
    sSQL = sSQL & " ORDER BY Sezionale"

    With Me.cboSezionaleCMR
        Set .Database = CnDMT
        .DisplayField = "Sezionale"
        .AddFieldKey "IDSezionale"
        .Sql = sSQL
        .Refresh
    End With
    
End Sub


Private Sub fncTipoOggetto()
    Dim sSQL As String
    
    sSQL = "SELECT RV_POTipoDocDaFatt.IDTipoOggetto, TipoOggetto.Oggetto "
    sSQL = sSQL & "FROM RV_POTipoDocDaFatt INNER JOIN "
    sSQL = sSQL & "TipoOggetto ON RV_POTipoDocDaFatt.IDTipoOggetto = TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & " ORDER BY Oggetto"

    With Me.CboTipoDocumento
        Set .Database = CnDMT
        .DisplayField = "Oggetto"
        .AddFieldKey "RV_POTipoDocDaFatt.IDTipoOggetto"
        .Sql = sSQL
        .Refresh
    End With
End Sub



Private Sub Form_Resize()
On Error Resume Next
  If Me.WindowState <> 1 Then
    

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

Private Sub Form_Unload(Cancel As Integer)


    ELIMINA_RIF_UTENTE_TMP
    
    If Me.CmdIndietro.Value = True Then
        If Not (rsGrigliaOrdini Is Nothing) Then
                rsGrigliaOrdini.Close
            Set rsGrigliaOrdini = Nothing
        End If
        
        FrmMain.InitControlli
        FrmMain.SettaggioIniziale
        FrmMain.Show
        Exit Sub
    End If
    
    
End Sub
Private Sub ELIMINA_RIF_UTENTE_TMP()
Dim sSQL As String

sSQL = "UPDATE RV_POTMPEvasioneOrdini SET "
sSQL = sSQL & "IDUtente=0, "
sSQL = sSQL & "Utente=" & fnNormString("") & ", "
sSQL = sSQL & "DaRegistrare=" & fnNormBoolean(0) & " "

sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser

CnDMT.Execute sSQL

End Sub

Private Sub GrigliaOrdini_DblClick()
On Error GoTo ERR_GrigliaOrdini_DblClick
    
'    If AVVIA_FATTURAZIONE = 1 Then Exit Sub
'
'    If (rsGrigliaOrdini.EOF) And (rsGrigliaOrdini.BOF) Then Exit Sub
'
'    AGGIORNA_GRIGLIA_ORDINI_PER_EVASIONE = False
'
'    frmAnnotazioniOrdine.Show vbModal
'
'    If AGGIORNA_GRIGLIA_ORDINI_PER_EVASIONE = True Then
'        rsGrigliaOrdini!DescrizioneCorpoDocEv = DESCRIZIONE_CORPO_PER_EVASIONE
'        rsGrigliaOrdini!IDLuogoPresaMerce = LINK_LUOGO_MERCE_PER_EVASIONE
'        rsGrigliaOrdini!IDVettoreSuccessivo = LINK_VETTORE_SUCCESSIVO_PER_EVASIONE
'
'        rsGrigliaOrdini.UpdateBatch
'
'        GrigliaOrdini.Refresh
'
'    End If
Exit Sub
ERR_GrigliaOrdini_DblClick:
    MsgBox Err.Description, vbCritical, "GrigliaOrdini_DblClick"
End Sub

Private Sub GrigliaOrdini_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_GrigliaOrdini_MouseUp
Dim IDTipoDocumento As Long
Dim IDSezionaleCliente As Long

    If AVVIA_FATTURAZIONE = 1 Then Exit Sub
    
    If (rsGrigliaOrdini.EOF) And (rsGrigliaOrdini.BOF) Then Exit Sub

    'rsGrigliaOrdini.Requery
    'rsGrigliaOrdini.Move Me.GrigliaOrdini.ListIndex - 1
    'GrigliaOrdini.Refresh
    'Controlla se l'utente ha cliccato su una riga valida
    If Me.GrigliaOrdini.HitTest(X, Y) > 0 Then
        'Controlla se le coordinate del cursore corrispondono alla colonna "Selezionato"
        If X > 0 And (X * Screen.TwipsPerPixelX) < GrigliaOrdini.ColumnsHeader(2).Width Then
            'Se non siamo in modalità filtri
            If GrigliaOrdini.GuiMode = dgNormal Then
                'Abilitiamo o disabilitiamo il check in base allo stato corrente
                sbSelectSelectedRow Not CBool(rsGrigliaOrdini.Fields("DaRegistrare").Value), 2
                
                NumeroOrdiniDaEvadere = GET_NUMERO_ORDINI_DA_EVADERE(TheApp.IDUser)
                
                If NumeroOrdiniDaEvadere > 1 Then
                    Me.cboPagamento.WriteOn 0
                    Me.cboTipoTrasporto.WriteOn 0
                    Me.cboVettore.WriteOn 0
                    Me.txtTargaAutomezzo.Text = ""
                    Me.txtIstruzioniMittente.Text = ""
                    Me.cboPagamento.Enabled = False
                    Me.cboTipoTrasporto.Enabled = False
                    Me.cboVettore.Enabled = False
                    Me.txtTargaAutomezzo.Enabled = False
                    Me.txtIstruzioniMittente.Enabled = False
                    
                    IDTipoDocumento = GET_TIPO_DOCUMENTO_CLIENTE(TheApp.IDUser, Me.CboTipoDocumento.CurrentID)
                    If IDTipoDocumento = 0 Then
                        GET_DEFAULT_AZIENDA
                    Else
                        Me.CboTipoDocumento.WriteOn IDTipoDocumento
                    End If
                Else
                    If NumeroOrdiniDaEvadere = 1 Then
                        Me.cboPagamento.Enabled = True
                        Me.cboTipoTrasporto.Enabled = True
                        Me.cboVettore.Enabled = True
                        Me.txtTargaAutomezzo.Enabled = True
                        Me.txtIstruzioniMittente.Enabled = True
                        IDTipoDocumento = GET_TIPO_DOCUMENTO_CLIENTE(TheApp.IDUser, Me.CboTipoDocumento.CurrentID)
                        If IDTipoDocumento = 0 Then
                            GET_DEFAULT_AZIENDA
                        Else
                            Me.CboTipoDocumento.WriteOn IDTipoDocumento
                        End If
                    End If
                    
                    If NumeroOrdiniDaEvadere = 0 Then
                        Me.cboPagamento.Enabled = True
                        Me.cboTipoTrasporto.Enabled = True
                        Me.cboVettore.Enabled = True
                        Me.txtTargaAutomezzo.Enabled = True
                        Me.txtIstruzioniMittente.Enabled = True
                        
                        GET_DEFAULT_AZIENDA
                    End If
                End If
                
                CboTipoDocumento_Click
            End If
        End If
    End If
Exit Sub
ERR_GrigliaOrdini_MouseUp:
    MsgBox Err.Description, vbCritical, "GrigliaOrdini_MouseUp"
End Sub

Private Sub GrigliaOrdini_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
On Error GoTo ERR_GrigliaOrdini_Reposition
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtAnnotazioniOrdine.Text = ""
Me.txtCausaleDocOrdSel.Text = ""
Me.cboTipoTrasportoOrdSel.WriteOn 0
Me.cboVettoreOrdSel.WriteOn 0
Me.txtTargaAutomezzoOrdSel.Text = ""
Me.txtIstruzioniMittenteOrdSel.Text = ""
Me.cboVettoreSuccOrdSel.WriteOn 0
Me.cboLuogoPresaMerceOrdSel.WriteOn 0
Me.txtDataArrivoLuogoOrdSel.Value = 0
Me.txtOraArrivoLuogoOrdSel.Value = 0
Me.cboAspettoEsterioreOrdSel.WriteOn 0



sSQL = "SELECT * FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(Me.GrigliaOrdini.AllColumns("IDOggetto").Value)

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtAnnotazioniOrdine.Text = fnNotNull(rs!RV_POAnnotazioniInterna)
    Me.txtCausaleDocOrdSel.Text = fnNotNull(rs!Doc_causale_documento)
    Me.cboTipoTrasportoOrdSel.WriteOn fnNotNullN(rs!Link_Doc_spedizione)
    Me.cboVettoreOrdSel.WriteOn fnNotNullN(rs!Link_Vet_vettore)
    Me.txtTargaAutomezzoOrdSel.Text = fnNotNull(rs!RV_POTargaAutomezzo)
    Me.txtIstruzioniMittenteOrdSel.Text = fnNotNull(rs!RV_POIstruzioniMittente)
    Me.cboVettoreSuccOrdSel.WriteOn fnNotNullN(rs!RV_POIDTrasportatoreSuccessivo)
    Me.cboLuogoPresaMerceOrdSel.WriteOn fnNotNullN(rs!RV_POIDLuogoPresaMerce)
    Me.txtDataArrivoLuogoOrdSel.Text = fnNotNull(rs!RV_PODataArrivoMerceLuogo)
    Me.txtOraArrivoLuogoOrdSel.Text = fnNotNull(rs!RV_POOraArrivoMerceLuogo)
    Me.cboAspettoEsterioreOrdSel.WriteOn fnNotNullN(rs!Link_Doc_aspetto_esteriore)
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GrigliaOrdini_Reposition:
    MsgBox Err.Description, vbCritical, "GrigliaOrdini_Reposition"
End Sub

Private Sub LabelLink1_BeforeRunServerApplication()
On Error GoTo ERR_LabelLink1_BeforeRunServerApplication

    Me.LabelLink1.DisableDoEvents = True
    
    Select Case Me.LVFattureCreate.ListItems(Me.LVFattureCreate.SelectedItem.Index).SubItems(1)
        Case 2
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_PODDTL"))
            Me.LabelLink1.IDReturn = CLng(Me.LVFattureCreate.SelectedItem.Text)
        Case 114
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POFAL"))
            Me.LabelLink1.IDReturn = CLng(Me.LVFattureCreate.SelectedItem.Text)
        Case 8
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POSNFL"))
            Me.LabelLink1.IDReturn = CLng(Me.LVFattureCreate.SelectedItem.Text)
    End Select
    
Exit Sub
ERR_LabelLink1_BeforeRunServerApplication:
    MsgBox Err.Description, vbCritical, "LabelLink1_BeforeRunServerApplication"
End Sub

Private Sub LVFattureCreate_Click()
On Error GoTo ERR_LVFattureCreate_Click
    'Inizializzazione della LabelLink
    '----------------------------
    'Set LabelLink1.Application = TheApp    'L’oggetto Application
    'LabelLink1.WindowHandleClient = Me.hWnd   'L’Handle del Form che contiene la LabelLink
    
    If Me.LVFattureCreate.ListItems.Count > 0 Then
        Me.LabelLink1.Visible = True
        Me.LabelLink1.Caption = "Visualizza "
        Select Case Me.LVFattureCreate.ListItems(Me.LVFattureCreate.SelectedItem.Index).SubItems(1)
            Case 2
                Me.LabelLink1.Caption = Me.LabelLink1.Caption & " il documento di trasporto "
            Case 114
                Me.LabelLink1.Caption = Me.LabelLink1.Caption & " la fattura accompagnatoria "
            Case 8
                Me.LabelLink1.Caption = Me.LabelLink1.Caption & " il documento corrispettivo "
        End Select
        
        Me.LabelLink1.Caption = Me.LabelLink1.Caption & "N° " & Me.LVFattureCreate.ListItems(Me.LVFattureCreate.SelectedItem.Index).SubItems(3)
        Me.LabelLink1.Caption = Me.LabelLink1.Caption & " del " & Me.LVFattureCreate.ListItems(Me.LVFattureCreate.SelectedItem.Index).SubItems(2)
        
    End If
    
    
    
    'Viene disabilitata la voce "Ricerca" del menu popup
    'LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False

Exit Sub
ERR_LVFattureCreate_Click:
    MsgBox Err.Description, vbCritical, "LVFattureCreate_Click"
End Sub

Private Function GET_SEZIONALE_PER_DEFAULT() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale "
sSQL = sSQL & "FROM DefaultFilialePerTipoOggetto "
sSQL = sSQL & "WHERE (IDTipoOggetto =" & Link_TipoOggettoPerDoc & ") And (IDFiliale = " & VarIDFiliale & ") And (IDReportTipoOggetto Is Null)"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_SEZIONALE_PER_DEFAULT = 0
Else
    GET_SEZIONALE_PER_DEFAULT = fnNotNullN(rs!IDSezionale)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub Settaggio()


    Set ObjDoc = New cDocument
    With ObjDoc
        fncEsercizio Me.txtDataDoc.Text
        Link_TipoOggetto = Me.CboTipoDocumento.CurrentID
        Set .Connection = CnDMT
        .IDAzienda = TheApp.IDFirm
        .IDAttivitaAzienda = VarIDAttivitaAzienda
        .IDFiliale = TheApp.Branch
        .SetTipoOggetto Link_TipoOggetto
        .TablesNames ObjDoc.IDTipoOggetto, sTabellaTestata, sTabellaDettaglio, sTabellaIVA, sTabellaScadenze
        .IDFunzione = GET_FUNZIONE(Me.CboTipoDocumento.CurrentID)
        .UseAutomation = True
        .IDEsercizio = fncEsercizio(Me.txtDataDoc.Text)
        .IDSezionale = Me.cboSezionale.CurrentID
        .IDTipoAnagrafica = 2
        .IDUtente = TheApp.IDUser
        .Descrizione = GET_DESCRIZIONE_TIPOOGGETTO(Me.CboTipoDocumento.CurrentID)
        .DataEmissione = Me.txtDataDoc.Text
        .Numero = Me.txtNumeroDocumento.Value
        
        If .Tables.Count = 0 Then
        'Se Tables.Count = 0 vuol dire che l'oggetto
        'DmtDocs non è mai stato inizializzato
            .Clear
            .SetTipoOggetto Me.CboTipoDocumento.CurrentID
        Else
            .ClearValues
        End If
    
    End With
End Sub
Private Function fncTestata(IDSocio As Long, IDDestinazione As Long, IDOggettoOrdine As Long, IDVettore As Long) As Boolean
'On Error GoTo ERR_fncTestata
Dim IDListinoDefault As Long
Dim Link_Pagamento As Long
Dim Link_Valuta_Cliente As Long

Dim LINK_VALUTA_NAZIONALE As Long
Dim SPESE_TRASPORTO_CLIENTE As Double


VARErroreFunzione = "fncTestata"
         
         With ObjDoc.Tables
        
            'Imposta la riga attiva per la tabella di testata
            
            ObjDoc.Tables(NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)).SetActiveRetail 1
            
            ObjDoc.ReadDataFromCliFo IDSocio, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            
            ObjDoc.ReadDataFromCliFoSite IDDestinazione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            
            Link_Valuta_Cliente = .Field("Link_Val_valuta", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))
                        
            '.Field "Doc_intra_cessione", GET_CESSIONE_INTRA_CLIENTE(IDSocio), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            
            If Abs(fnNotNullN(.Field("Doc_intra_cessione", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))) = 1 Then
                GET_CAMPI_INSTRASTAT_AZIENDA IDSocio, TheApp.IDFirm, TheApp.Branch, ObjDoc.DataEmissione
            End If
            
            SCRIVI_RIFERIMENTO_ORDINE IDOggettoOrdine
            
            If Me.cboTipoTrasporto.CurrentID > 0 Then
                ObjDoc.Field "Link_Doc_spedizione", Me.cboTipoTrasporto.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
                If Me.cboTipoTrasporto.CurrentID = 3 Then
                    ObjDoc.ReadDataFromCarrier Me.cboVettore.CurrentID, MainCarrier, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
                End If
            End If
            
            If Me.cboVettore.CurrentID > 0 Then
                ObjDoc.Field "Link_Doc_spedizione", 3, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
                ObjDoc.ReadDataFromCarrier Me.cboVettore.CurrentID, MainCarrier, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            End If
            
            'LISTINO
            If .Field("Link_Doc_Listino", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) = 0 Then

                If IDDestinazione > 0 Then
                    IDListinoDefault = GET_LISTINO_PER_DESTINAZIONE(IDSocio, IDDestinazione)
                End If
                
                If IDListinoDefault = 0 Then
                    IDListinoDefault = GET_LISTINO_DEFAULT(IDSocio)
                    If IDListinoDefault = 0 Then
                        IDListinoDefault = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
                    End If
                End If
                
                .Field "Link_Doc_Listino", IDListinoDefault, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            End If
                        
            SPESE_TRASPORTO_CLIENTE = GET_IMPORTO_SPESE_TRASPORTO(IDSocio, IDOggettoOrdine, IDDestinazione)
            .Field "Doc_causale_trasporto", fnSetCausaleDocumento, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Link_Doc_Magazzino", Me.cboMagazzino.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Link_Doc_sezionale", Me.cboSezionale.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_prefisso", GET_PREFISSO_SEZ(Me.cboSezionale.CurrentID), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_data", ObjDoc.DataEmissione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_data_inizio_trasporto", ObjDoc.DataEmissione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_ora_inizio_trasporto", Time, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Spe_trasporto_neutro", SPESE_TRASPORTO_CLIENTE, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Doc_crea_scadenze", fnNormBoolean(1), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "RV_PODataCompetenzaLiq", ObjDoc.DataEmissione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            
            'PAGAMENTO DOPO I DATI DELL'ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Me.cboPagamento.CurrentID > 0 Then
                ObjDoc.ReadDataFromPayment Me.cboPagamento.CurrentID
            Else
                If fnNotNullN(.Field("Link_Doc_pagamento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) = 0 Then
                    Link_Pagamento = GET_MODALITA_PAGAMENTO(ObjDoc.DataEmissione, IDSocio)
                        
                    If Link_Pagamento > 0 Then
                        ObjDoc.ReadDataFromPayment Link_Pagamento
                    Else
                        ObjDoc.ReadDataFromPayment fnNotNullN(.Field("Link_Doc_pagamento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
                    End If
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Len(Trim(Me.txtCausaleDocumento.Text)) > 0 Then
                ObjDoc.Field "Doc_causale_documento", Me.txtCausaleDocumento.Text, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            End If


            'AGENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If NO_RIP_AGENTE_IN_DOC_EVASIONE = 0 Then
                If .Field("Link_Doc_agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) = 0 Then
                    ObjDoc.ReadDataFromAgent GET_LINK_AGENTE_CLIENTE(IDSocio, TheApp.IDFirm), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            'VALUTA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If .Field("Link_Val_valuta", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) = 0 Then
                If Link_Valuta_Cliente = 0 Then
                    .Field "Link_Val_valuta", Me.CboValuta.CurrentID, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
                Else
                    .Field "Link_Val_valuta", ObjDoc.DBDefaults.Link_Val_valuta_nazionale, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
                End If
            End If
            
            GET_CAMBIO_VALUTA Me.CboValuta.CurrentID, .Field("Link_Val_valuta", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
            If Me.chkCMR.Value = vbChecked Then
                GET_DATI_CMR Me.cboSezionaleCMR.CurrentID, GET_PeriodoIVA(Year(Me.txtDataDocCMR.Text)), Me.txtNumeroDocCMR.Value, Me.txtDataDocCMR.Text
            End If
            
            'If TIPO_SCONTO_CLIENTE = 1 Then
            '    .Field "Sco_percentuale_fine_documento", SCONTO_ABITUALE, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            'Else
            '    .Field "Sco_percentuale_fine_documento", 0, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            'End If
            
            .Field "Doc_data_plafond", ObjDoc.DataEmissione, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Link_Spe_esenti_art_10_IVA", ObjDoc.DBDefaults.Link_Spe_esenti_art_10_IVA, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "Link_Spe_bolli_eff_art_15_IVA", ObjDoc.DBDefaults.Link_Spe_bolli_eff_art_15_IVA, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)

            .Field "RV_POAnnotazioni1", Me.txtAnnotazione01.Text, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "RV_POAnnotazioni2", Me.txtAnnotazione02.Text, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            .Field "RV_POAnnotazioni3", Me.txtAnnotazione03.Text, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
            
            fn_ProtocolloICE ObjDoc
        End With
        
        fncTestata = True
     
Exit Function
ERR_fncTestata:
    fncTestata = False
    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento
    
End Function
Private Sub fn_ProtocolloICE(oDoc As DmtDocs.cDocument)
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    If USA_PROT_ICE_PERIODO = 0 Then
    
        sSQL = "SELECT RV_POSchemaCoop.IDFiliale, RV_POSchemaCoop.IDUtente, RV_POProgProtocolloICE.IDRV_POProtocolloICE, "
        sSQL = sSQL & "RV_POProgProtocolloICE.IDRV_POProgProtocolloICE, RV_POProgProtocolloICE.Progressivo , RV_POProgProtocolloICE.Predefinito, RV_POProtocolloICE.ProtocolloICE "
        sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
        sSQL = sSQL & "RV_POProgProtocolloICE ON "
        sSQL = sSQL & "RV_POSchemaCoop.IDRV_POSchemaCoop = dbo.RV_POProgProtocolloICE.IDRV_POSchemaCoop LEFT OUTER JOIN "
        sSQL = sSQL & "RV_POProtocolloICE ON RV_POProgProtocolloICE.IDRV_POProtocolloICE = dbo.RV_POProtocolloICE.IDRV_POProtocolloICE "
        sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale = " & oDoc.IDFiliale & ") "
        sSQL = sSQL & " AND (RV_POSchemaCoop.IDUtente = 0) "
        sSQL = sSQL & " AND (RV_POProgProtocolloICE.Predefinito = " & fnNormBoolean(1) & ")"
        
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            oDoc.Field "RV_POIndicaProtICE", fnNormBoolean(0)
            oDoc.Field "RV_POIDProtocolloICE", 0
            oDoc.Field "RV_POProtICE", ""
            oDoc.Field "RV_PONumeroProtICE", 0
            Link_Prog_Protocollo_ICE = 0
            NumeroProgressivoICE = 0
            oDoc.Field "RV_POIDSchemaProtICE", 0
            oDoc.Field "RV_POProtocolloICEPeriodo", 0
            
        Else
            oDoc.Field "RV_POIndicaProtICE", fnNormBoolean(1)
            oDoc.Field "RV_POIDProtocolloICE", fnNotNullN(rs!IDRV_POProtocolloICE)
            oDoc.Field "RV_POProtICE", fnNotNull(rs!ProtocolloICE)
            oDoc.Field "RV_PONumeroProtICE", fnNotNullN(rs!Progressivo)
            Link_Prog_Protocollo_ICE = fnNotNullN(rs!IDRV_POProgProtocolloICE)
            NumeroProgressivoICE = fnNotNullN(rs!Progressivo)

            oDoc.Field "RV_POIDSchemaProtICE", fnNotNull(rs!IDRV_POProgProtocolloICE)
            oDoc.Field "RV_POProtocolloICEPeriodo", USA_PROT_ICE_PERIODO
        End If
        
        
        rs.CloseResultset
        Set rs = Nothing
    End If
    
    If USA_PROT_ICE_PERIODO = 1 Then
        sSQL = "SELECT RV_POProgProtocolloICEPeriodo.IDRV_POProgProtocolloICEPeriodo, RV_POProgProtocolloICEPeriodo.IDFiliale, RV_POProgProtocolloICEPeriodo.IDAzienda, "
        sSQL = sSQL & "RV_POProgProtocolloICEPeriodo.IDRV_POProtocolloICE, RV_POProgProtocolloICEPeriodo.DaData, RV_POProgProtocolloICEPeriodo.AData, RV_POProgProtocolloICEPeriodo.Progressivo,"
        sSQL = sSQL & "RV_POProtocolloICE.ProtocolloICE "
        sSQL = sSQL & "FROM RV_POProgProtocolloICEPeriodo INNER JOIN "
        sSQL = sSQL & "RV_POProtocolloICE ON RV_POProgProtocolloICEPeriodo.IDRV_POProtocolloICE = RV_POProtocolloICE.IDRV_POProtocolloICE "
        sSQL = sSQL & " WHERE IDFiliale=" & oDoc.IDFiliale
        sSQL = sSQL & " AND DaData<=" & fnNormDate(oDoc.DataEmissione)
        sSQL = sSQL & " AND AData>=" & fnNormDate(oDoc.DataEmissione)
        
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            oDoc.Field "RV_POIndicaProtICE", fnNormBoolean(0)
            oDoc.Field "RV_POIDProtocolloICE", 0
            oDoc.Field "RV_POProtICE", ""
            oDoc.Field "RV_PONumeroProtICE", 0
            Link_Prog_Protocollo_ICE = 0
            NumeroProgressivoICE = 0
            oDoc.Field "RV_POIDSchemaProtICE", 0
            oDoc.Field "RV_POProtocolloICEPeriodo", 0
            
        Else
            oDoc.Field "RV_POIndicaProtICE", fnNormBoolean(1)
            oDoc.Field "RV_POIDProtocolloICE", fnNotNullN(rs!IDRV_POProtocolloICE)
            oDoc.Field "RV_POProtICE", fnNotNull(rs!ProtocolloICE)
            oDoc.Field "RV_PONumeroProtICE", fnNotNullN(rs!Progressivo)
            Link_Prog_Protocollo_ICE = fnNotNullN(rs!IDRV_POProgProtocolloICEPeriodo)
            NumeroProgressivoICE = fnNotNullN(rs!Progressivo)

            oDoc.Field "RV_POIDSchemaProtICE", fnNotNull(rs!IDRV_POProgProtocolloICEPeriodo)
            oDoc.Field "RV_POProtocolloICEPeriodo", USA_PROT_ICE_PERIODO
        End If
        
        
        rs.CloseResultset
        Set rs = Nothing
    End If
End Sub
Private Sub GET_CAMBIO_VALUTA(IDValutaRiferimento As Long, IDValutaDocumento As Long, oDoc As DmtDocs.cDocument, TabellaTestata As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDCambio, DataCambio, Valore FROM Cambio "
sSQL = sSQL & "WHERE IDValuta=" & IDValutaDocumento
sSQL = sSQL & " AND IDValutaDiRiferimento=" & IDValutaRiferimento
sSQL = sSQL & " ORDER BY DataCambio DESC"

Set rs = CnDMT.OpenResultset(sSQL)

If Not (rs.EOF) Then
    oDoc.Field "Link_Val_cambio", fnNotNullN(rs!IDCambio), TabellaTestata
    oDoc.Field "Val_valore_cambio", fnNotNullN(rs!Valore), TabellaTestata
    oDoc.Field "Val_data_cambio", fnNotNull(rs!DataCambio), TabellaTestata
End If
       
rs.CloseResultset
Set rs = Nothing

End Sub
Private Function fncRighe(IDCliente As Long, IDSitoPerAnagrafica As Long, IDOggettoOrdine As Long) As Boolean
On Error GoTo ERR_fncRighe
VARErroreFunzione = "fncRighe"
Dim I As Integer
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ProgressivoArticolo As Long
Dim Link_Riga As Long
Dim ImportoPedana As Double
Dim IDListinoDefault As Long
Dim ImportoLiquidazione As Double
Dim Sconto1 As Double
Dim Sconto2 As Double
Dim ImportoImballo As Double
Dim DescrizioneRigaDaOrdine As String
Dim Link_IVA_Articolo_Riga As Long
Dim Aliquota_IVA_Articolo_Riga As Long
Dim ImballoARendere As Long
Dim rsImballiANoleggio As ADODB.Recordset
Dim ImportoUnitarioArticoloMerceNetta As Double
Dim ImportoUnitarioArticoloMerceImballo As Double

Dim LINK_REGOLA_PROVV As Long


'creazione recordset per per noleggi''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rsImballiANoleggio = New ADODB.Recordset
rsImballiANoleggio.CursorLocation = adUseClient
rsImballiANoleggio.Fields.Append "IDArticoloImballo", adInteger, , adFldIsNullable
rsImballiANoleggio.Fields.Append "Quantita", adDouble, , adFldIsNullable
rsImballiANoleggio.Fields.Append "IDAgente", adInteger, , adFldIsNullable
rsImballiANoleggio.Fields.Append "IDListino", adInteger, , adFldIsNullable
rsImballiANoleggio.Open , , adOpenKeyset, adLockBatchOptimistic

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    I = 1
    Link_Riga = 1
    ProgressivoArticolo = 0
    
    IDListinoDefault = GET_LISTINO_DEFAULT(IDCliente)
    If IDListinoDefault = 0 Then
        IDListinoDefault = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
    End If
    
    sSQL = "SELECT * FROM RV_POIEEvasioneOrdine "
    sSQL = sSQL & "WHERE DaRegistrare=" & fnNormBoolean(1)
    sSQL = sSQL & " AND IDOggetto=" & IDOggettoOrdine
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
    Me.lblInfoStatus.Caption = "Elaborazione ordine numero " & fnNotNullN(rs!NumeroOrdine) & " del " & fnNotNull(rs!DataOrdine)
    DoEvents
        ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
            
            GET_IVA_ARTICOLO fnNotNullN(rs!IDArticolo)
            
            If (RIP_IVA_DA_DOC_COLL = 1) Then
                If (fnNotNullN(rs!IDValoriOggettoDettaglioRigaOrd) > 0) Then
                    GET_IVA_ART_DA_ORDINE fnNotNullN(rs!IDValoriOggettoDettaglioRigaOrd)
                End If
            End If
            
            
            GET_PARAMETRI_LIQUIDAZIONE_ARTICOLO fnNotNullN(rs!IDArticolo)
                    
            ImportoLiquidazione = 0
            
            ObjDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_codice", fnNotNull(rs!CodiceArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_descrizione", fnNotNull(rs!Articolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Qta_UM), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)

            ObjDoc.Field "Art_sco_in_percentuale_1", fnNotNullN(rs!Sconto1), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_sco_in_percentuale_2", fnNotNullN(rs!Sconto2), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "Art_importo_totale_lordo_IVA", (fnNotNullN(rs!ImportoUnitarioArticolo) * fnNotNullN(rs!Qta_UM)) + (((fnNotNullN(rs!ImportoUnitarioArticolo) * fnNotNullN(rs!Qta_UM)) / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_importo_totale_netto_IVA", fnNotNullN(rs!ImportoUnitarioArticolo) * fnNotNullN(rs!Qta_UM), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_netto_IVA", fnNotNullN(rs!ImportoUnitarioArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_lordo_IVA", fnNotNullN(rs!ImportoUnitarioArticolo) + ((fnNotNullN(rs!ImportoUnitarioArticolo) / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_pre_uni_net_sco_net_IVA", fnNotNullN(rs!ImportoUnitarioArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_pre_uni_net_sco_lor_IVA", fnNotNullN(rs!ImportoUnitarioArticolo) + ((fnNotNullN(rs!ImportoUnitarioArticolo) / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_Importo_totale_neutro", (fnNotNullN(rs!ImportoUnitarioArticolo) * fnNotNullN(rs!Qta_UM)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_neutro", fnNotNullN(rs!ImportoUnitarioArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_Importo_netto_IVA", fnNotNullN(rs!ImportoUnitarioArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_importo_net_sconto_lor_IVA", (fnNotNullN(rs!ImportoUnitarioArticolo) * fnNotNullN(rs!Qta_UM)) + (((fnNotNullN(rs!ImportoUnitarioArticolo) * fnNotNullN(rs!Qta_UM)) / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_importo_net_sconto_net_IVA", (fnNotNullN(rs!ImportoUnitarioArticolo) * fnNotNullN(rs!Qta_UM)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA_ARTICOLO(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            Else
                ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            End If
            
            Link_IVA_Articolo_Riga = ObjDoc.Field("Link_art_IVA", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
            Aliquota_IVA_Articolo_Riga = ObjDoc.Field("Art_aliquota_IVA", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
            
            ObjDoc.Field "Art_numero_colli", fnNotNullN(rs!Colli), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_Peso", fnNotNullN(rs!PesoLordo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_tara", fnNotNullN(rs!Tara), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_quantita_pezzi", fnNotNullN(rs!Pezzi), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "Link_Art_unita_di_misura", fnNotNullN(rs!IDUnitaDiMisura), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(fnNotNullN(rs!IDUnitaDiMisura)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
           
            LINK_UM_COOP = fnGetUMCoop(ObjDoc.Field("Link_Art_unita_di_misura", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))
            
            ''''AGENTE
            If ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) > 0 Then
                LINK_REGOLA_PROVV = GET_LINK_REGOLA_PROVV_AGE(ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), TheApp.IDFirm)
                
                If LINK_REGOLA_PROVV > 0 Then
                    ObjDoc.Field "Link_Art_agente", ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_nome", ObjDoc.Field("Doc_age_nome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_codice", ObjDoc.Field("Doc_age_codice", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_ragione_sociale", ObjDoc.Field("Doc_age_ragione_sociale", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    ObjDoc.Field "Link_Art_age_regola_provv", LINK_REGOLA_PROVV, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_regola_provv", GET_DESCRIZIONE_REGOLA_PROVV(LINK_REGOLA_PROVV), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    If fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))) > 0 Then
                        Sconto1 = fnNotNullN(rs!Sconto1)
                        Sconto2 = fnNotNullN(rs!Sconto2)
                        
                        If GESTIONE_ORDINE_VIVAIO = 1 Then
                            If ((fnNotNullN(rs!RV_POImportoUnitarioListino) > fnNotNullN(rs!ImportoUnitarioArticolo)) And (fnNotNullN(rs!Sconto1) = 0)) And (fnNotNullN(rs!Sconto2 = 0)) Then
                                Sconto2 = 0
                                Sconto1 = (1 - (fnNotNullN(rs!ImportoUnitarioArticolo) / fnNotNullN(rs!RV_POImportoUnitarioListino))) * 100
                                Sconto1 = fnRoundDown(Sconto1)
                            End If
                        End If
                        
                        ObjDoc.Field "Link_Art_age_tipo_ordine", GET_LINK_TIPO_ORDINE(LINK_REGOLA_PROVV, Sconto1, Sconto2), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
            End If
            
            'If FLAG_INTRASTAT_DOC = 1 Then
            If Abs(fnNotNullN(ObjDoc.Field("Doc_intra_cessione", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))) = 1 Then
                GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDArticolo), TheApp.IDFirm, TheApp.Branch, ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
            End If
            
            ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POTipoRiga", 1, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_POIDCalibro", fnNotNullN(rs!IDRV_POCalibro), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoCategoria", fnNotNullN(rs!IDRV_POTipoCategoria), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoLavorazione", fnNotNullN(rs!IDTipoLavorazione), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PODataConferimento", fnNotNull(rs!DataConferimento), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDConferimentoRighe", fnNotNullN(rs!IDRV_POCaricoMerceRighe), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDAssegnazioneMerce", fnNotNullN(rs!IDRV_POAssegnazioneMerce), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDProcessoIVGamma", fnNotNullN(rs!IDRV_POProcessoIVGamma), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDSocio", fnNotNullN(rs!IDSocio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDAnagraficaFatturazione", fnNotNullN(rs!IDAnagraficaFatturazione), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceSocio", fnNotNull(rs!CodiceSocio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POSocio", fnNotNull(rs!Socio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PONomeSocio", fnNotNull(rs!NomeSocio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POLottoCampagna", fnNotNull(rs!LottoDiConferimento), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceLotto", fnNotNull(rs!CodiceLottoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POImportoImballoInArticolo", rs!MerceInclusoImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_PODataLavorazione", fnNotNull(rs!DataDocumento), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POVariazionePrezzoManuale", GET_IMPORTO_ART_TRATT_CLI(IDCliente, fnNotNullN(rs!IDArticolo), (fnNotNullN(rs!Qta_UM) * Moltiplicatore)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)

            ObjDoc.Field "RV_POIDPedana", fnNotNull(rs!IDRV_POPedana), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodicePedana", fnNotNull(rs!CodicePedana), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoPedana", GET_LINK_TIPO_PEDANA(fnNotNull(rs!IDRV_POPedana)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POPesoPedana", GET_PESO_PEDANA(fnNotNull(rs!IDRV_POPedana)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_POImportoUnitarioListino", fnNotNullN(rs!RV_POImportoUnitarioListino), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_PONotaRigaOrdRaggr", fnNotNull(rs!NotaRigaOrdRaggr), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDValoriOggettoDettaglio0010", fnNotNullN(rs!IDValoriOggettoDettaglioRigaOrd), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POImportoImballoSel", fnNotNullN(rs!ImportoUnitarioImballo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POAnnotazioniAggiuntiveLav", fnNotNull(rs!AnnotazioniAggiuntiveLav), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_PODataOrdineCliente", rs!DataOrdineCliente, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PONumeroOrdineCliente", rs!NumeroOrdineCliente, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PODataOrdineInterno", rs!DataOrdine, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PONumeroOrdineInterno", rs!NumeroOrdine, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            If AGGIORNA_PREZZO_MEDIO = 0 Then
                ObjDoc.Field "RV_POPrezzoMedioInLiq", GET_PREZZO_MEDIO_CLIENTE(IDCliente, fnNotNullN(rs!IDArticolo)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            Else
                ObjDoc.Field "RV_POPrezzoMedioInLiq", GET_DATI_RIGA_CONFERIMENTO(fnNotNullN(rs!IDRV_POCaricoMerceRighe), "PrezzoMedio"), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            End If
            
            ObjDoc.Field "RV_POIDTipoImportoVenditaLiq", GET_FORZATURA_PREZZO_LIQ_CLIENTE(IDCliente, fnNotNullN(rs!IDArticolo)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoDocumentoCoop", fnNotNullN(rs!IDTipoDocumentoCoop), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PO01_IDSezionaleRighe", fnNotNullN(rs!RV_PO01_IDSezionaleRighe), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PO01_NumeroPassaportoRighe", fnNotNullN(rs!RV_PO01_NumeroPassaportoRighe), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POQuantitaLiq", (fnNotNullN(rs!Qta_UM) * Moltiplicatore), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            If (LINK_UM_LIQ > 0) Then
                Select Case LINK_UM_LIQ
                    Case 1
                        ObjDoc.Field "RV_POQuantitaLiq", (fnNotNullN(rs!Colli) * Moltiplicatore), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    Case 2
                        ObjDoc.Field "RV_POQuantitaLiq", (fnNotNullN(rs!PesoLordo) * Moltiplicatore), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    Case 3
                        ObjDoc.Field "RV_POQuantitaLiq", (fnNotNullN(rs!PesoNetto) * Moltiplicatore), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    Case 4
                        ObjDoc.Field "RV_POQuantitaLiq", (fnNotNullN(rs!Tara) * Moltiplicatore), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    Case 5
                        ObjDoc.Field "RV_POQuantitaLiq", (fnNotNullN(rs!Pezzi) * Moltiplicatore), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                End Select
            End If
            If fnNotNullN(rs!MerceInclusoImballo) = False Then
                ObjDoc.Field "RV_POImportoDaLiq", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                ImportoUnitarioArticoloMerceNetta = fnNotNullN(rs!ImportoUnitarioArticolo)
                ImportoUnitarioArticoloMerceNetta = ImportoUnitarioArticoloMerceNetta - ((ImportoUnitarioArticoloMerceNetta / 100) * fnNotNullN(rs!Sconto1))
                ImportoUnitarioArticoloMerceNetta = ImportoUnitarioArticoloMerceNetta - ((ImportoUnitarioArticoloMerceNetta / 100) * fnNotNullN(rs!Sconto2))
                
                'ObjDoc.Field "RV_POImportoMerceNetta", fnNotNullN(rs!ImportoUnitarioArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POImportoMerceNetta", ImportoUnitarioArticoloMerceNetta, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POVariazionePrezzoImballo", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            Else
                If fnNotNullN(rs!IDImballoVendita) > 0 Then
                    ImportoLiquidazione = 0
                    ImportoLiquidazione = ImportoLiquidazione - sbCalcolaImportoLiquidazione(LINK_UM_COOP, IDListinoDefault, fnNotNullN(rs!Qta_UM), fnNotNullN(rs!Colli), fnNotNullN(rs!IDImballoVendita), fnNotNullN(rs!ImportoUnitarioArticolo), fnNotNullN(rs!ImportoUnitarioImballo), ObjDoc.Field("RV_POQuantitaLiq", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))
                    
                    'DA METTERE SOTTO PARAMETRO QUESTO CALCOLO
    '                ImportoLiquidazione = ImportoLiquidazione - ((ImportoLiquidazione / 100) * fnNotNullN(rs!Sconto1))
    '                ImportoLiquidazione = ImportoLiquidazione - ((ImportoLiquidazione / 100) * fnNotNullN(rs!Sconto2))
                    
                    ObjDoc.Field "RV_POImportoDaLiq", ImportoLiquidazione, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ImportoUnitarioArticoloMerceNetta = fnNotNullN(rs!ImportoUnitarioArticolo)
                    ImportoUnitarioArticoloMerceNetta = ImportoUnitarioArticoloMerceNetta - ((ImportoUnitarioArticoloMerceNetta / 100) * fnNotNullN(rs!Sconto1))
                    ImportoUnitarioArticoloMerceNetta = ImportoUnitarioArticoloMerceNetta - ((ImportoUnitarioArticoloMerceNetta / 100) * fnNotNullN(rs!Sconto2))
    
                    ImportoUnitarioArticoloMerceImballo = GET_VARIAZIONE_PREZZO_IMBALLO(LINK_UM_COOP, IDListinoDefault, fnNotNullN(rs!Qta_UM), fnNotNullN(rs!Colli), fnNotNullN(rs!IDImballoVendita), fnNotNullN(rs!ImportoUnitarioArticolo), fnNotNullN(rs!ImportoUnitarioImballo), ObjDoc.Field("RV_POQuantitaLiq", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))
                    
                    'DA METTERE SOTTO PARAMETRO QUESTO CALCOLO
    '                ImportoUnitarioArticoloMerceImballo = ImportoUnitarioArticoloMerceImballo - ((ImportoUnitarioArticoloMerceImballo / 100) * fnNotNullN(rs!Sconto1))
    '                ImportoUnitarioArticoloMerceImballo = ImportoUnitarioArticoloMerceImballo - ((ImportoUnitarioArticoloMerceImballo / 100) * fnNotNullN(rs!Sconto2))
                    
                    'ObjDoc.Field "RV_POVariazionePrezzoImballo", GET_VARIAZIONE_PREZZO_IMBALLO(LINK_UM_COOP, IDListinoDefault, fnNotNullN(rs!Qta_UM), fnNotNullN(rs!Colli), fnNotNullN(rs!IDImballoVendita), fnNotNullN(rs!ImportoUnitarioArticolo), fnNotNullN(rs!ImportoUnitarioImballo), ObjDoc.Field("RV_POQuantitaLiq", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    'ObjDoc.Field "RV_POImportoMerceNetta", fnNotNullN(rs!ImportoUnitarioArticolo) - ObjDoc.Field("RV_POVariazionePrezzoImballo", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
                    ObjDoc.Field "RV_POVariazionePrezzoImballo", ImportoUnitarioArticoloMerceImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "RV_POImportoMerceNetta", ImportoUnitarioArticoloMerceNetta - ImportoUnitarioArticoloMerceImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    ObjDoc.Field "RV_POImportoImballoInArticolo", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "RV_POImportoDaLiq", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    ImportoUnitarioArticoloMerceNetta = fnNotNullN(rs!ImportoUnitarioArticolo)
                    ImportoUnitarioArticoloMerceNetta = ImportoUnitarioArticoloMerceNetta - ((ImportoUnitarioArticoloMerceNetta / 100) * fnNotNullN(rs!Sconto1))
                    ImportoUnitarioArticoloMerceNetta = ImportoUnitarioArticoloMerceNetta - ((ImportoUnitarioArticoloMerceNetta / 100) * fnNotNullN(rs!Sconto2))
                    
                    'ObjDoc.Field "RV_POImportoMerceNetta", fnNotNullN(rs!ImportoUnitarioArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "RV_POImportoMerceNetta", ImportoUnitarioArticoloMerceNetta, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "RV_POVariazionePrezzoImballo", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                End If
            End If
  
            ObjDoc.Field "RV_POIDAssegnazioneMerce", fnNotNullN(rs!IDRV_POAssegnazioneMerce), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDProcessoIVGamma", fnNotNullN(rs!IDRV_POProcessoIVGamma), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            'ProgressivoArticolo = ProgressivoArticolo + 1
            
            ObjDoc.Field "Art_volume", fnNotNullN(rs!VolumeImballo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)

            ObjDoc.Field "RV_POIDImballoPrim", rs!IDArticoloImballoPrimario, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceImballoPrim", rs!CodiceArticoloImbPrim, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PODescrizioneImballoPrim", rs!ArticoloImbPrim, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PONumeroConfezioniPerImballo", rs!NumeroConfezioniPerImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POTaraConfezioneImballo", rs!TaraConfezioneImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCostoConfezioneImballo", rs!CostoConfezioneImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POQuantitaTotaleConfImballo", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCostoKit", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POQuantitaPerCollo", fnNotNullN(rs!QuantitaPerCollo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POPesoPerCollo", fnNotNullN(rs!PesoPerCollo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POMoltiplicatorePerCollo", fnNotNullN(rs!MoltiplicatorePerCollo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "ID_Art_dettaglio_prog", ObjDoc.SetIDArtDettaglioProg, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(fnNotNullN(rs!IDArticolo), ObjDoc.Field("Link_Nom_anagrafica", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc.Field("Link_Nom_ult_sito", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDProcessoLavorazioneRighe", fnNotNullN(rs!IDRV_POProcessoLavorazioneRighe), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDProcessoLavorazione", fnNotNullN(rs!IDRV_POProcessoLavorazione), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDLineaProduzione", fnNotNullN(rs!IDRV_POLineaProduzione), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoUtilizzoLinea", fnNotNullN(rs!IDRV_POTipoUtilizzoLinea), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDLottoCampagnaLavorazione", fnNotNullN(rs!IDRV_PO01_LottoCampagna), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            If (ObjDoc.IDTipoOggetto <> 8) Then
                sbLoadElectronicInvoiceData4Article fnNotNullN(ObjDoc.Field("ID_Art_dettaglio_prog", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))), fnNotNullN(ObjDoc.Field("Link_Art_articolo", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))
            End If
            
            If fnNotNullN(rs!IDImballoVendita) > 0 Then
                GET_IVA_ARTICOLO fnNotNullN(rs!IDImballoVendita)
                ImballoARendere = GET_TIPO_IMBALLO_A_RENDERE(fnNotNullN(rs!IDImballoVendita))

                If (ImballoARendere = 1) And (FLAG_IVA_IMBALLO_A_RENDERE = 1) Then
                    ObjDoc.Field "RV_POIDIvaImballo", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    If FLAG_IVA_UGUALE = 1 Then
                        ObjDoc.Field "RV_POIDIvaImballo", Link_IVA_Articolo_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                            ObjDoc.Field "RV_POIDIvaImballo", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        Else
                            ObjDoc.Field "RV_POIDIvaImballo", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        End If
                    End If
                End If
                
                ObjDoc.Field "RV_PORigaCompleta", 1, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDImballo", fnNotNullN(rs!IDImballoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POCodiceImballo", fnNotNull(rs!CodiceImballoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_PODescrizioneImballo", fnNotNull(rs!ImballoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                If fnNotNullN(rs!MerceInclusoImballo) = False Then
                    ObjDoc.Field "RV_POImportoUnitarioImballo", fnNotNullN(rs!ImportoUnitarioImballo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    ObjDoc.Field "RV_POImportoUnitarioImballo", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                End If
            Else
                ObjDoc.Field "RV_PORigaCompleta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            End If
            
            
            If fnNotNullN(rs!IDImballoVendita) > 0 Then
                GET_IVA_ARTICOLO fnNotNullN(rs!IDImballoVendita)
                
                If (RIP_IVA_DA_DOC_COLL = 1) Then
                    
                    GET_IVA_IMB_DA_ORDINE fnNotNullN(rs!IDValoriOggettoDettaglioRigaOrd), fnNotNullN(rs!IDOggettoOrdine)
                    
                End If
                
                ImballoARendere = GET_TIPO_IMBALLO_A_RENDERE(fnNotNullN(rs!IDImballoVendita))
                
                I = I + 1
                ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                
                ObjDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDImballoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_codice", fnNotNull(rs!CodiceImballoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_descrizione", fnNotNull(rs!ImballoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Colli), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_tara", fnNotNullN(rs!TaraUnitaria), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                If (ImballoARendere = 1) And (FLAG_IVA_IMBALLO_A_RENDERE = 1) Then
                    ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    If FLAG_IVA_UGUALE = 1 Then
                        ObjDoc.Field "Link_art_IVA", Link_IVA_Articolo_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_aliquota_IVA", Aliquota_IVA_Articolo_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                            ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                            ObjDoc.Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA_ARTICOLO(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        Else
                            ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                            ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        End If
                    End If
                End If
                If fnNotNullN(rs!MerceInclusoImballo) = False Then
                    ImportoImballo = fnNotNullN(rs!ImportoUnitarioImballo)
                    If fnNotNullN(rs!Colli) = 0 Then
                        ImportoImballo = 0
                    End If
                Else
                    ImportoImballo = 0
                End If
                
                ObjDoc.Field "Art_prezzo_unitario_netto_IVA", ImportoImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_prezzo_unitario_lordo_IVA", ImportoImballo + ((ImportoImballo / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                ObjDoc.Field "Art_pre_uni_net_sco_net_IVA", ImportoImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_pre_uni_net_sco_lor_IVA", ImportoImballo + ((ImportoImballo / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                ObjDoc.Field "Art_Importo_totale_neutro", (ImportoImballo * fnNotNullN(rs!Colli)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_prezzo_unitario_neutro", ImportoImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                ObjDoc.Field "Art_Importo_netto_IVA", (ImportoImballo * fnNotNullN(rs!Colli)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_importo_totale_netto_IVA", (ImportoImballo * fnNotNullN(rs!Colli)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    
                ObjDoc.Field "Link_Art_unita_di_misura", GET_LINK_UM_ART(fnNotNullN(rs!IDImballoVendita)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(fnNotNullN(ObjDoc.Field("Link_Art_unita_di_misura", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)

                ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POTipoRiga", 2, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_PORigaCompleta", 1, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDImballo", fnNotNullN(rs!IDImballoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                'ProgressivoArticolo = ProgressivoArticolo + 1
                ObjDoc.Field "ID_Art_dettaglio_prog", ObjDoc.SetIDArtDettaglioProg, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(fnNotNullN(rs!IDImballoVendita), ObjDoc.Field("Link_Nom_anagrafica", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc.Field("Link_Nom_ult_sito", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                If (ObjDoc.IDTipoOggetto <> 8) Then
                    sbLoadElectronicInvoiceData4Article fnNotNullN(ObjDoc.Field("ID_Art_dettaglio_prog", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))), fnNotNullN(ObjDoc.Field("Link_Art_articolo", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))
                End If
                ''''AGENTE
                If ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) > 0 Then
                    LINK_REGOLA_PROVV = GET_LINK_REGOLA_PROVV_AGE(ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), TheApp.IDFirm)
        
                    If LINK_REGOLA_PROVV > 0 Then
                        ObjDoc.Field "Link_Art_agente", ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_nome", ObjDoc.Field("Doc_age_nome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_codice", ObjDoc.Field("Doc_age_codice", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_ragione_sociale", ObjDoc.Field("Doc_age_ragione_sociale", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        
                        ObjDoc.Field "Link_Art_age_regola_provv", LINK_REGOLA_PROVV, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_regola_provv", GET_DESCRIZIONE_REGOLA_PROVV(LINK_REGOLA_PROVV), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        If fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))) > 0 Then
                            ObjDoc.Field "Link_Art_age_tipo_ordine", GET_LINK_TIPO_ORDINE(LINK_REGOLA_PROVV, fnNotNullN(rs!Sconto1), fnNotNullN(rs!Sconto2)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        End If
                    End If
                End If

                If Abs(fnNotNullN(ObjDoc.Field("Doc_intra_cessione", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))) = 1 Then
                    GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDImballoVendita), TheApp.IDFirm, TheApp.Branch, ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
                End If
                
                ObjDoc.Field "RV_POIDCalibro", fnNotNullN(rs!IDRV_POCalibro), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDTipoCategoria", fnNotNullN(rs!IDRV_POTipoCategoria), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDTipoLavorazione", fnNotNullN(rs!IDTipoLavorazione), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_PODataConferimento", fnNotNull(rs!DataConferimento), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDConferimentoRighe", fnNotNullN(rs!IDRV_POCaricoMerceRighe), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDAssegnazioneMerce", fnNotNullN(rs!IDRV_POAssegnazioneMerce), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDProcessoIVGamma", fnNotNullN(rs!IDRV_POProcessoIVGamma), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDSocio", fnNotNullN(rs!IDSocio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POCodiceSocio", fnNotNull(rs!CodiceSocio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POSocio", fnNotNull(rs!Socio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_PONomeSocio", fnNotNull(rs!NomeSocio), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POLottoCampagna", fnNotNull(rs!LottoDiConferimento), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POCodiceLotto", fnNotNull(rs!CodiceLottoVendita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POImportoImballoInArticolo", rs!MerceInclusoImballo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDIvaImballo", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POPrezzoMedioInLiq", GET_PREZZO_MEDIO_CLIENTE(IDCliente, fnNotNullN(rs!IDArticolo)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDTipoDocumentoCoop", fnNotNullN(rs!IDTipoDocumentoCoop), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDPedana", fnNotNull(rs!IDRV_POPedana), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POCodicePedana", fnNotNull(rs!CodicePedana), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDTipoPedana", GET_LINK_TIPO_PEDANA(fnNotNull(rs!IDRV_POPedana)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POPesoPedana", GET_PESO_PEDANA(fnNotNull(rs!IDRV_POPedana)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                GET_IMBALLI_NOLEGGIO rsImballiANoleggio, ObjDoc.Field("Link_art_articolo", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)), ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)), ObjDoc.Field("Link_Art_agente", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)), IDCliente
                
            End If
    Link_Riga = Link_Riga + 1
    I = I + 1
        
    rs.MoveNext
    Wend
rs.CloseResultset
Set rs = Nothing

'''''''''''''''''''''''''''''''''INSERIMENTO PEDANE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If GESTIONE_ORDINE_VIVAIO = 0 Then
    sSQL = "SELECT IDArticoloImballo, Articolo, CodiceArticolo, COUNT(IDRV_POPedana) AS Quantita, NumeroPallet "
    sSQL = sSQL & "FROM RV_POIEArticoloPedanaOrdine "
    sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine
    sSQL = sSQL & "GROUP BY IDArticoloImballo, Articolo, CodiceArticolo, NumeroPallet "
        
        Set rs = CnDMT.OpenResultset(sSQL)
        DoEvents
        While Not rs.EOF
            If fnNotNullN(rs!IDArticoloImballo) > 0 Then
                
                GET_IVA_ARTICOLO fnNotNullN(rs!IDArticoloImballo)
                
                ImballoARendere = GET_TIPO_IMBALLO_A_RENDERE(fnNotNullN(rs!IDArticoloImballo))
                
                If IDListinoDefault > 0 Then
                    ImportoPedana = GET_PREZZO_IMBALLO(IDListinoDefault, fnNotNullN(rs!IDArticoloImballo))
                Else
                    ImportoPedana = 0
                End If
                
                ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                
                ObjDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDArticoloImballo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_codice", fnNotNull(rs!CodiceArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_descrizione", fnNotNull(rs!Articolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                If (fnNotNullN(rs!NumeroPallet) = 0) Then
                    ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Quantita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Quantita) * fnNotNullN(rs!NumeroPallet), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                End If
                
                ObjDoc.Field "Art_prezzo_unitario_netto_IVA", ImportoPedana, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_prezzo_unitario_lordo_IVA", ImportoPedana + ((ImportoPedana / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                ObjDoc.Field "Art_pre_uni_net_sco_net_IVA", ImportoPedana, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_pre_uni_net_sco_lor_IVA", ImportoPedana + ((ImportoPedana / 100) * AliquotaIvaArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                ObjDoc.Field "Art_Importo_totale_neutro", (ImportoPedana * fnNotNullN(rs!Quantita)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_prezzo_unitario_neutro", ImportoPedana, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                ObjDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    
                ObjDoc.Field "Art_Importo_netto_IVA", (ImportoPedana * fnNotNullN(rs!Quantita)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                'If PESO_PEDANA_IN_VENDITA = 1 Then
                ObjDoc.Field "Art_tara", GET_PESO_ARTICOLO(fnNotNullN(rs!IDArticoloImballo)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_peso", (ObjDoc.Field("Art_tara", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)) * fnNotNullN(ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                'End If
    
                If (ImballoARendere = 1) And (FLAG_IVA_IMBALLO_A_RENDERE = 1) Then
                    ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                        ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA_ARTICOLO(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    Else
                        ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
                
                
                ''''AGENTE
                If ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) > 0 Then
                    LINK_REGOLA_PROVV = GET_LINK_REGOLA_PROVV_AGE(ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), TheApp.IDFirm)
                    
                    If LINK_REGOLA_PROVV > 0 Then
                        ObjDoc.Field "Link_Art_agente", ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_nome", ObjDoc.Field("Doc_age_nome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_codice", ObjDoc.Field("Doc_age_codice", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_ragione_sociale", ObjDoc.Field("Doc_age_ragione_sociale", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Link_Art_age_regola_provv", LINK_REGOLA_PROVV, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        ObjDoc.Field "Art_age_regola_provv", GET_DESCRIZIONE_REGOLA_PROVV(LINK_REGOLA_PROVV), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        
                        If fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))) > 0 Then
                            ObjDoc.Field "Link_Art_age_tipo_ordine", GET_LINK_TIPO_ORDINE(LINK_REGOLA_PROVV, 0, 0), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                        End If
                    End If
                End If
                
                If FLAG_INTRASTAT_DOC = 1 Then
                    GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDArticoloImballo), TheApp.IDFirm, TheApp.Branch, ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
                End If
                
                'ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                'ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NOMETABELLADETTAGLIO & fnGetHex(ObjDoc.IDCorpo)
                
                'ObjDoc.Field "Art_tara", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_importo_totale_netto_IVA", (ImportoPedana * fnNotNullN(rs!Quantita)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                ObjDoc.Field "Link_Art_unita_di_misura", GET_LINK_UM_ART(fnNotNullN(rs!IDArticoloImballo)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POTipoRiga", 2, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_PORigaCompleta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDImballo", fnNotNullN(rs!IDArticoloImballo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                'ProgressivoArticolo = ProgressivoArticolo + 1
                ObjDoc.Field "ID_Art_dettaglio_prog", ObjDoc.SetIDArtDettaglioProg, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(fnNotNullN(rs!IDArticoloImballo), ObjDoc.Field("Link_Nom_anagrafica", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc.Field("Link_Nom_ult_sito", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                If (ObjDoc.IDTipoOggetto <> 8) Then
                    sbLoadElectronicInvoiceData4Article fnNotNullN(ObjDoc.Field("ID_Art_dettaglio_prog", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))), fnNotNullN(ObjDoc.Field("Link_Art_articolo", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))
                End If
                ObjDoc.Field "RV_POIDCalibro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDTipoCategoria", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDTipoLavorazione", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_PODataConferimento", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDConferimentoRighe", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POIDSocio", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POCodiceSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_PONomeSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POLottoCampagna", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_POCodiceLotto", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "RV_ArticoloPedana", 1, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
                Link_Riga = Link_Riga + 1
                I = I + 1
            End If
    
        rs.MoveNext
        Wend
    
    rs.CloseResultset
    Set rs = Nothing
End If


If GESTIONE_ORDINE_VIVAIO = 1 Then
    GET_RIGA_PEDANA IDOggettoOrdine, I, Link_Riga, ProgressivoArticolo
    GET_RIGA_PIANALE IDOggettoOrdine, I, Link_Riga, ProgressivoArticolo
    GET_RIGA_PROLUNGA IDOggettoOrdine, I, Link_Riga, ProgressivoArticolo
End If

'Imballi a noleggio''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not ((rsImballiANoleggio.EOF) And (rsImballiANoleggio.BOF)) Then
    rsImballiANoleggio.MoveFirst
    
    While Not rsImballiANoleggio.EOF
        GET_IVA_ARTICOLO fnNotNullN(rsImballiANoleggio!IDArticoloImballo)
        ImballoARendere = GET_TIPO_IMBALLO_A_RENDERE(fnNotNullN(rsImballiANoleggio!IDArticoloImballo))
        ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
        ObjDoc.ReadDataFromArticle fnNotNullN(rsImballiANoleggio!IDArticoloImballo)
        ObjDoc.Field "Art_quantita_totale", fnNotNullN(rsImballiANoleggio!Quantita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.ReadDataFromAgent fnNotNullN(rsImballiANoleggio!IDAgente), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_prezzo_unitario_neutro", GET_PREZZO_IMBALLO(fnNotNullN(rsImballiANoleggio!IDListino), fnNotNullN(rsImballiANoleggio!IDArticoloImballo)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        
        
        'If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
        '    ObjDoc.ReadDataFromIva fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        'End If
        
        If (ImballoARendere = 1) And (FLAG_IVA_IMBALLO_A_RENDERE = 1) Then
            ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        Else
            If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA_ARTICOLO(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            Else
                ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            End If
        End If
        
        ObjDoc.Field "ID_Art_dettaglio_prog", ObjDoc.SetIDArtDettaglioProg, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(fnNotNullN(rsImballiANoleggio!IDArticoloImballo), ObjDoc.Field("Link_Nom_anagrafica", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc.Field("Link_Nom_ult_sito", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        If (ObjDoc.IDTipoOggetto <> 8) Then
            sbLoadElectronicInvoiceData4Article fnNotNullN(ObjDoc.Field("ID_Art_dettaglio_prog", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))), fnNotNullN(ObjDoc.Field("Link_Art_articolo", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)))
        End If
        ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "RV_POTipoRiga", 2, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "RV_PORigaCompleta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        
        
    rsImballiANoleggio.MoveNext
    I = I + 1
    Wend
    

End If

rsImballiANoleggio.Close
Set rsImballiANoleggio = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

DescrizioneRigaDaOrdine = GET_DESCRIZIONE_RIGA_CORPO(IDOggettoOrdine)
If Len(Trim(DescrizioneRigaDaOrdine)) > 0 Then
    ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
    
    ObjDoc.Field "Link_Art_articolo", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_codice", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_descrizione", DescrizioneRigaDaOrdine, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_quantita_totale", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    
    ObjDoc.Field "Art_prezzo_unitario_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_prezzo_unitario_lordo_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        
    ObjDoc.Field "Link_art_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_aliquota_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "RV_POTipoRiga", 2, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "RV_PORigaCompleta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    If FLAG_INTRASTAT_DOC = 1 Then
        ObjDoc.Field "Art_intra_non_riporto", 1, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    End If
    ObjDoc.Field "ID_Art_dettaglio_prog", ObjDoc.SetIDArtDettaglioProg, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)

End If

If (Unita_Progresso + Me.ProgressBar1.Value) >= Me.ProgressBar1.Max Then
    Me.ProgressBar1.Value = Me.ProgressBar1.Max
Else
    Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
End If

fncRighe = True
Exit Function
ERR_fncRighe:
    fncRighe = False
    VARErroreIDIntervento = "GENERALITA':" & vbCrLf & "IDCliente : " & IDClienteOP
    'VARErroreIDArticolo = vbCrLf & "Articolo: " & rsArt!Articolo
    VARErroreGenerico = Err.Description & vbCrLf & VARErroreIDIntervento ' & VARErroreIDArticolo
    MsgBox Err.Description, vbCritical, "ERRORE"

End Function
Private Function InserimentoDMT(IDCliente As Long, Cliente As String, IDDestinazione As Long, SitoPerAnagrafica As String, IDOggettoOrdine As Long) As Boolean
On Error GoTo ERR_InserimentoDMT
Dim VarNumeroDoc As String
Dim Link_Oggetto As Long
Dim sSQL As String
Dim TestoMessaggio As String

Screen.MousePointer = vbHourglass
    If ObjDoc.IDTipoOggetto <> 8 Then
        ConsolidaDettaglioFatturaElettronica
    End If
   
    Set ObjDoc.Scadenze = Nothing
    ObjDoc.PerformDocument Nothing
    
    ObjDoc.AllowCreateMovements = True
     
    If GESTIONE_ORDINE_VIVAIO = 1 Then
        ObjDoc.Field "Tot_numero_colli", TotaleNumeroColliOrdine, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
        ObjDoc.Field "Tot_peso", TotalePesoOrdine, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    'CONTROLLO PLAFOND'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim sMsgPlafond As String
    If ObjDoc.PlafondExceed Then
        sMsgPlafond = ObjDoc.PlafondLastMessage
        If Len(sMsgPlafond) > 0 Then
            If ObjDoc.PlafondLastMessageStyle = vbCritical Then
                sbMsgError sMsgPlafond, TheApp.FunctionName
                Screen.MousePointer = 0
                Exit Function
            Else
                MsgBox sMsgPlafond, vbInformation, TheApp.FunctionName
                
                TestoMessaggio = "ATTENZIONE!!!" & vbCrLf
                TestoMessaggio = TestoMessaggio & "Il documento è stato emesso regolarmente, "
                TestoMessaggio = TestoMessaggio & "tuttavia a causa del plafond superato, "
                TestoMessaggio = TestoMessaggio & "è necessario visionarlo ed applicare le opportune modifiche"
                
                MsgBox TestoMessaggio, vbInformation, TheApp.FunctionName

            End If
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    VarNumeroDoc = ObjDoc.Insert
    
    If VarNumeroDoc > 0 Then
        'ObjDoc.GeneratePreviewPNC
        ' ObjDoc.Update
        
        Me.lblInfoStatus.Caption = "AGGIORNAMENTO DOCUMENTO NUMERO " & VarNumeroDoc
        DoEvents
        'Link_Oggetto = GET_IDOggetto(CLng(VarNumeroDoc))
        Link_Oggetto = ObjDoc.IDOggetto
        
        'AGGIORNAMENTO DESCRIZIONE DOCUMENTO
        'sSQL = "UPDATE " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " SET "
        'sSQL = sSQL & "RV_PODescrizioneDocumento=" & fnNormString(StringaRiferimento & " n: " & ObjDoc.Tables.Field("Doc_prefisso", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) & "/" & VarNumeroDoc & " del " & Me.txtDataDoc.Text)
        'sSQL = sSQL & " WHERE IDOggetto=" & Link_Oggetto
        'CnDMT.Execute sSQL
        If (NonRiportaInXMLRifVsNumOrd = 0) Then
            If Len(fnNotNull(ObjDoc.Tables.Field("Doc_numero_vs_ordine_di_rifer", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))) > 0 Then
                If ObjDoc.IDTipoOggetto <> 8 Then
                    SCRIVI_ORD_CLI_RIF_XML fnNotNull(ObjDoc.Tables.Field("Doc_data_vs_ordine_di_rifer", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), fnNotNull(ObjDoc.Tables.Field("Doc_numero_vs_ordine_di_rifer", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
                End If
            End If
        End If
        
        fnAggiornaDescrizioneDocumento ObjDoc.Tables.Field("Doc_prefisso", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        fnAggiornaDescrizioneOrdineCliente NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo), ObjDoc.Tables.Field("Doc_data_vs_ordine_di_rifer", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc.Tables.Field("Doc_numero_vs_ordine_di_rifer", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))
        fnAggiornaDescrizioneOrdineInterno NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo), ObjDoc.Tables.Field("Doc_data_ns_ordine_di_rifer", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc.Tables.Field("Doc_numero_ns_ordine_di_rifer", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))
        fnAggiornaDescrizioneProtocollo NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo), ObjDoc.Tables.Field("RV_POProtICE", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), ObjDoc.Tables.Field("RV_PONumeroProtICE", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))
        fnAggiornaDescrizioneDocumentoOrd ObjDoc.Tables.Field("Doc_prefisso", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        
        'INSERIMENTO COMMISSSIONE PER DOCUMENTO
        If (ATTIVA_COMMISSIONI_DA_ORDINE = 0) Then
            IMPOSTA_COMMISSIONI_PER_CLIENTE IDCliente, IDOggettoOrdine, Link_Oggetto, IDDestinazione
            IMPOSTA_COMMISSIONI_PER_TIPO_PEDANA IDCliente, IDOggettoOrdine, Link_Oggetto, IDDestinazione
        Else
            IMPOSTA_COMMISSIONI_PER_CLIENTE_DA_ORD IDCliente, IDOggettoOrdine, Link_Oggetto, IDDestinazione
            If (RIC_COMM_TIPO_PED_DA_ORD = 1) Then
                
            End If
        End If
        
        AGGIORNA_RIGHE_DOCUMENTO NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        AGGIORNA_RIGHE_COMMISSIONI ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo), ObjDoc.DataEmissione
        AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA_TOTALE NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        'AGGIORNAMENTO MOVIMENTI
        AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI 0, Link_Oggetto, IDDestinazione
        
        
        
        'AGGIORNAMENTO ORDINI FATTURATI IN TMP
        sSQL = "UPDATE  RV_POTMPEvasioneOrdini SET "
        sSQL = sSQL & "InFatturazione=" & fnNormBoolean(1)
        sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoOrdine
        CnDMT.Execute sSQL
        
        'AGGIORNAMENTO ORDINI FATTURATI
        sSQL = "INSERT INTO RV_POTMPOrdiniFatturati ("
        sSQL = sSQL & "IDAzienda, IDOggetto, IDTipoOggetto, NumeroDocumento, "
        sSQL = sSQL & "DataDocumento, Cliente, SitoPerAnagrafica, Vettore, IDUtente) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & TheApp.IDFirm & ", "
        sSQL = sSQL & Link_Oggetto & ", "
        sSQL = sSQL & Me.CboTipoDocumento.CurrentID & ", "
        sSQL = sSQL & fnNormNumber(VarNumeroDoc) & ", "
        sSQL = sSQL & fnNormDate(Me.txtDataDoc.Text) & ", "
        sSQL = sSQL & fnNormString(Cliente) & ", "
        sSQL = sSQL & fnNormString(SitoPerAnagrafica) & ", "
        sSQL = sSQL & fnNormString(Vettore) & ", "
        sSQL = sSQL & TheApp.IDUser
        sSQL = sSQL & ")"
        
        CnDMT.Execute sSQL
        
        'AGGIORNAMENTO PROTOCOLLO ICE
        If Link_Prog_Protocollo_ICE > 0 Then
            If USA_PROT_ICE_PERIODO = 0 Then
                sSQL = "UPDATE RV_POProgProtocolloICE SET "
                sSQL = sSQL & "Progressivo = " & fnNormNumber(NumeroProgressivoICE + 1) & " "
                sSQL = sSQL & "WHERE IDRV_POProgProtocolloICE = " & Link_Prog_Protocollo_ICE
                CnDMT.Execute sSQL
            End If
            If USA_PROT_ICE_PERIODO = 1 Then
                sSQL = "UPDATE RV_POProgProtocolloICEPeriodo SET "
                sSQL = sSQL & "Progressivo = " & fnNormNumber(NumeroProgressivoICE + 1) & " "
                sSQL = sSQL & "WHERE IDRV_POProgProtocolloICEPeriodo = " & Link_Prog_Protocollo_ICE
                CnDMT.Execute sSQL
            End If
        End If
        
        If Me.chkChiudiOrdini.Value = vbChecked Then
            'CHIUSURA ORDINE
            sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
            sSQL = sSQL & "Doc_ordine_chiuso=1 "
            sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoOrdine
            CnDMT.Execute sSQL
        End If
        
        'ELIMINAZIONE ORDINE
        sSQL = "DELETE FROM  RV_POTMPEvasioneOrdini  "
        sSQL = sSQL & "WHERE InFatturazione=" & fnNormBoolean(1)
        sSQL = sSQL & " AND IDOggetto=" & IDOggettoOrdine
        CnDMT.Execute sSQL
        
        AGGIORNA_ORDINE_PADRE IDOggettoOrdine
        SCRIVI_CAUSALI_DOC ObjDoc.IDOggetto
        
        If Me.chkStampaFattura.Value = vbChecked Then
            ObjDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, Link_Oggetto, Link_TipoOggetto
            StampaDocumento Link_Oggetto
        End If
        
    End If

Screen.MousePointer = vbDefault
    
Exit Function

ERR_InserimentoDMT:
    InserimentoDMT = False
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
    
End Function
Private Sub SetDefault()
    'Imposta i default per i campi relativi alla valuta corrente
    'al tipo di arrotondamento, alle spese e ai bolli
    'Questi default sono obbligatori per il calcolo del documento
    Set cDefault = New Collection
    'Valore di arrotondamento per la valuta corrente
    cDefault.Add 1, "Val_arrotondamento"
    'Tipo di arrotondomento per la valuta corrente
    cDefault.Add 1, "Link_Val_tipo_arrotondamento"
    'ID della valuta corrente
    cDefault.Add 0, "Link_Val_valuta"
    'Spesse incassi in percentuale
    cDefault.Add 0, "Nom_spese_incassi_perc"
    'Importo del bollo
    cDefault.Add 0, "Nom_bollo_esente"
    'Importo limite per il pagamento del bollo
    cDefault.Add 0, "Nom_bollo_esente_limite"
    'ID del contratto bancario azienda
    cDefault.Add 0, "Link_Doc_contratto_bancario_az"
    'ID della natura delle scadenze
    cDefault.Add 0, "IDNaturaScadenza"
End Sub
Private Sub TrovaAnagrafica(IDAna As Long)
    Dim sSQL As String
    Dim RSAna As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.CodiceFiscale, Anagrafica.PartitaIva, Anagrafica.Indirizzo, Anagrafica.Cap, Comune.Comune, Provincia.Provincia, Anagrafica.Telefono, Anagrafica.Fax, Fornitore.IDPagamentoDefault, Fornitore.Codice"
    sSQL = sSQL & " FROM ((Anagrafica LEFT JOIN Comune ON Anagrafica.IDComune = Comune.IDComune) LEFT JOIN Provincia ON Comune.IDProvincia = Provincia.IDProvincia) LEFT JOIN Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica"
    sSQL = sSQL & " WHERE (((Anagrafica.IDAnagrafica)=" & IDAna & "))"
    
    
    Set RSAna = CnDMT.OpenResultset(sSQL)
        If RSAna.EOF = False Then
            ArrayCli(0, 0) = RSAna!IDAnagrafica
            ArrayCli(0, 1) = RSAna!Anagrafica
            ArrayCli(0, 2) = fnNotNull(RSAna!Nome)
            ArrayCli(0, 3) = fnNotNull(RSAna!CodiceFiscale)
            ArrayCli(0, 4) = fnNotNull(RSAna!PartitaIva)
            ArrayCli(0, 5) = fnNotNull(RSAna!Indirizzo)
            ArrayCli(0, 6) = fnNotNull(RSAna!Cap)
            ArrayCli(0, 7) = fnNotNull(RSAna!Comune)
            ArrayCli(0, 8) = fnNotNull(RSAna!Provincia)
            ArrayCli(0, 9) = fnNotNull(RSAna!Fax)
            ArrayCli(0, 10) = fnNotNull(RSAna!Telefono)
            ArrayCli(0, 11) = IIf((IsNull(RSAna!IDPagamentoDefault)), FrmFine.cboPagamento.CurrentID, fnNotNull(RSAna!IDPagamentoDefault))
            ArrayCli(0, 12) = fnNotNull(RSAna!Codice)
        End If
        
    
    RSAna.CloseResultset
    Set RSAna = Nothing
End Sub
Private Function fncTrovaReport(NomeReport As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_fncTrovaReport
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE ((ReportTipoOggetto=" & fnNormString(NomeReport) & ") AND (IDTipoOggetto=" & IDTipoOggetto_coop & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaReport = rs!IDReportTipoOggetto
Else
    fncTrovaReport = 0
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_fncTrovaReport:
    MsgBox Err.Description, vbCritical, "Trova report per stampa"
    fncTrovaReport = 0

End Function
Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & Link_TipoOggetto_Coop
    sSQL = sSQL & " AND IDFiliale = " & TheApp.Branch
    
    CnDMT.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
Private Function fncIDTipoOggettoPrg(Gestore As String) As Long
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto, Gestore.Gestore"
    sSQL = sSQL & " FROM Gestore INNER JOIN TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore"
    sSQL = sSQL & " WHERE (((Gestore.Gestore)=" & fnNormString(Gestore) & "))"
    
    Set rs = CnDMT.OpenResultset(sSQL)
        
    If rs.EOF = False Then
        fncIDTipoOggettoPrg = rs!IDTipoOggetto
    Else
        fncIDTipoOggettoPrg = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub fncReport()
Dim oActivity As IActivity
Dim o As Activity
Dim oFilter As Filter
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'With Me.cboReport
'    Set .Database = CnDMT
'    .AddFieldKey "IDReportTipoOggetto"
'    .DisplayField = "ReportTipoOggetto"
'    .Sql = "SELECT * FROM ReportTipoOggetto WHERE IDTipoOggetto=" & Link_TipoOggetto_Coop
'    .Fill
'End With


sSQL = "SELECT IDReportTipoOggetto, ReportTipoOggetto "
sSQL = sSQL & "FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & Link_TipoOggetto_Coop

Set rs = CnDMT.OpenResultset(sSQL)
Me.lstReport.Clear
While Not rs.EOF
    Me.lstReport.AddItem fnNotNull(rs!ReportTipoOggetto)
    Me.lstReport.ItemData(Me.lstReport.NewIndex) = fnNotNullN(rs!IDReportTipoOggetto)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing


'With Me.ActivityBox
'    .Activities.Clear
'
'    'Aggiunge l'attività dei reports
'    Set oActivity = .Activities.Add("DmtActBoxLib.ReportsActivity", "Reports")
'    Set oActivity.Connection = CnDMT.InternalConnection
'
'    If Link_TipoOggetto_Coop > 0 Then
'        oActivity.Load Link_TipoOggetto_Coop, TheApp.IDFirm
'    Else
'        oActivity.Load 1, TheApp.IDFirm
'    End If
'
'    Set o = oActivity
'    Set oReportsActivity = o.InternalClass
'
'
'    'Imposta quale attività deve essere attivata per default
'    If m_DefaultActivity <> "" Then
'        Set .CurrentActivity = .Activities(m_DefaultActivity)
'    End If
    
'    'ridisegna il controllo
'    .Redraw = True
'
'    oReportsActivity.Is4DlgPrint = False
'End With

'Me.txtNumeroCopie.Value = GET_NUMERO_COPIE(oReportsActivity.DefaultReportName, Link_TipoOggetto_Coop)
End Sub
Private Sub DEFAULT_REPORT_LISTA()
Dim I As Integer
Dim IDReport As Long

IDReport = fnDefaultReport

For I = 0 To Me.lstReport.ListCount - 1
    If Me.lstReport.ItemData(I) = IDReport Then
        Me.lstReport.Selected(I) = True
    End If
Next

End Sub

Private Function fnDefaultReport() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDReportTipoOggetto FROM DefaultFilialePerTipoOggetto "
sSQL = sSQL & "WHERE IDFiliale=" & VarIDFiliale
sSQL = sSQL & " AND IDTipoOggetto=" & Link_TipoOggetto_Coop

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fnDefaultReport = fnNotNullN(rs!IDReportTipoOggetto)
Else
    fnDefaultReport = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub StampaDocumento(IDOggetto As Long)
On Error GoTo ERR_StampaDocumento
Dim I As Integer
Dim IDReportDefault As Long
Dim NumeroCopie As Long

Set oReport = New dmtReportLib.dmtReport
Set oReport.Connection = CnDMT
    
    If MenuOptions.DBType = 1 Then
        'parametri di accesso al database ACCESS
        oReport.Password = "dmt192981046"
        oReport.User = "admin"
    Else
        'parametri di accesso al database SQL Server
        oReport.Password = TheApp.Password '""
        oReport.User = TheApp.User '"sa"
    End If

        IDReportDefault = fnDefaultReport
    'Imposta l'idfiliale di appartenenza del documento da stampare
        oReport.BranchID = TheApp.Branch 'IDFiliale
    'Imposta l'identificativo del tipo di documento
        oReport.DocTypeID = Link_TipoOggetto_Coop
        'oReport.Where = "IDOggetto = 873" '& Val(Me.Txt_Reg_IDRegistro)
        oReport.Where = "ValoriOggettoPerTipo" & fnGetHex(Link_TipoOggetto) & ".IDOggetto = " & IDOggetto
        oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
        
        
    
        For I = 0 To Me.lstReport.ListCount - 1
            'oReport.DoPrint FrmFine.cboReport.Text
            If Me.lstReport.Selected(I) = True Then
                Me.lstReport.ListIndex = I
                If chkStampaNumeroCopie.Value = vbUnchecked Then
                    NumeroCopie = GET_NUMERO_COPIE(Me.lstReport.List(I), oReport.DocTypeID)
                Else
                    NumeroCopie = FrmFine.txtNumeroCopie.Text
                End If
                'If NumeroCopie > FrmFine.txtNumeroCopie.Text Then
                '    oReport.Copies = NumeroCopie
                'Else
                '    oReport.Copies = FrmFine.txtNumeroCopie.Text
                'End If
                oReport.Copies = NumeroCopie
                fncImpostaDefaultReport Me.lstReport.ItemData(I)
                oReport.DoPrint oReport.PrinterName
            End If
        Next
        
        fncImpostaDefaultReport IDReportDefault
    
Exit Sub
ERR_StampaDocumento:
    MsgBox Err.Description, vbCritical, "Stampa documento"
End Sub


Private Sub fnGrigliaOrdini()
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader
    
    
    
sSQL = "SELECT RV_POTMPEvasioneOrdini.IDOggetto, RV_POTMPEvasioneOrdini.DaRegistrare, RV_POTMPEvasioneOrdini.NumeroRiga, RV_POTMPEvasioneOrdini.NumeroOrdine, RV_POTMPEvasioneOrdini.DataOrdine, "
sSQL = sSQL & "RV_POTMPEvasioneOrdini.IDCliente, RV_POTMPEvasioneOrdini.IDSitoPerAnagrafica, RV_POTMPEvasioneOrdini.Cliente,"
sSQL = sSQL & "RV_POTMPEvasioneOrdini.SitoPerAnagrafica, RV_POTMPEvasioneOrdini.IDUtente, RV_POTMPEvasioneOrdini.Utente, "
sSQL = sSQL & "RV_POTMPEvasioneOrdini.IDVettore, RV_POTMPEvasioneOrdini.Vettore, RV_POTMPEvasioneOrdini.IDAzienda, RV_POTMPEvasioneOrdini.DescrizioneCorpoDocEv, "
sSQL = sSQL & "RV_POTMPEvasioneOrdini.IDLuogoPresaMerce, RV_POTMPEvasioneOrdini.IDVettoreSuccessivo, RV_POTMPEvasioneOrdini.NumeroListaPrelievo "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE InFatturazione=" & fnNormBoolean(0)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        Set rsGrigliaOrdini = New ADODB.Recordset
        rsGrigliaOrdini.CursorLocation = adUseClient
        rsGrigliaOrdini.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockBatchOptimistic
            'Set rsEvent = rsGriglia2.Data
    
        With Me.GrigliaOrdini
            .EnableMove = True
            .UpdatePosition = False
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectCell
            .ColumnsHeader.Clear
            
            .ColumnsHeader.Add "NumeroRiga", "ID", dgInteger, False, 500, dgAlignleft
            Set cl = .ColumnsHeader.Add("DaRegistrare", "Da registrare", dgBoolean, True, 1200, dgAligncenter)
                cl.Editable = True
            .ColumnsHeader.Add "IDAzienda", "IDAzienda", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDUtente", "IDUtente", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Utente", "Utente", dgchar, True, 1000, dgAlignleft
            .ColumnsHeader.Add "DataOrdine", "Data ordine", dgDate, True, 1100, dgAlignleft
            .ColumnsHeader.Add "NumeroOrdine", "N° ordine", dgNumeric, True, 1000, dgAlignRight
            .ColumnsHeader.Add "NumeroListaPrelievo", "N° lista", dgNumeric, True, 1000, dgAlignRight
            .ColumnsHeader.Add "IDCliente", "IDCliente", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDSitoPerAnagrafica", "IDSitoPerAnagrafica", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Cliente", "Cliente", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "IDVettore", "IDVettore", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "Vettore", "Vettore", dgchar, True, 2000, dgAlignleft
            .ColumnsHeader.Add "DescrizioneCorpoDocEv", "Descrizione finale del corpo documento di evasione", dgchar, True, 5000, dgAlignleft
            .ColumnsHeader.Add "IDLuogoPresaMerce", "IDLuogoPresaMerce", dgInteger, False, 500, dgAlignleft
            .ColumnsHeader.Add "IDVettoreSuccessivo", "IDVettoreSuccessivo", dgInteger, False, 500, dgAlignleft
    
            Set .Recordset = rsGrigliaOrdini
            .Refresh
        End With
    
    CnDMT.CursorLocation = OLDCursor
End Sub
Private Sub GrigliaOrdini_KeyPress(KeyAscii As Integer)
Dim Testo As String
Dim NumeroOrdiniDaEvadere As Long

    If AVVIA_FATTURAZIONE = 1 Then Exit Sub

    'Intercetta la pressione della barra spaziatrice sulla DmtGrid
    If (rsGrigliaOrdini.EOF) And (rsGrigliaOrdini.BOF) Then Exit Sub
    
    rsGrigliaOrdini.Requery
    rsGrigliaOrdini.Move Me.GrigliaOrdini.ListIndex - 1
    GrigliaOrdini.Refresh
    
    If KeyAscii = vbKeySpace Then
        'Se non siamo in modalità filtri
        If GrigliaOrdini.GuiMode = dgNormal Then
        'Abilitiamo o disabilitiamo il check in base allo stato corrente
            sbSelectSelectedRow Not CBool(rsGrigliaOrdini.Fields("DaRegistrare").Value), 2
        End If
    End If
    
    NumeroOrdiniDaEvadere = GET_NUMERO_ORDINI_DA_EVADERE(TheApp.IDUser)
    
    If NumeroOrdiniDaEvadere > 1 Then
        Me.cboPagamento.WriteOn 0
        Me.cboTipoTrasporto.WriteOn 0
        Me.cboVettore.WriteOn 0
        Me.txtTargaAutomezzo.Text = ""
        Me.txtIstruzioniMittente.Text = ""
        
        Me.cboPagamento.Enabled = False
        Me.cboTipoTrasporto.Enabled = False
        Me.cboVettore.Enabled = False
        Me.txtTargaAutomezzo.Enabled = False
        Me.txtIstruzioniMittente.Enabled = False
        
        IDTipoDocumento = GET_TIPO_DOCUMENTO_CLIENTE(TheApp.IDUser, Me.CboTipoDocumento.CurrentID)
        If IDTipoDocumento = 0 Then
            GET_DEFAULT_AZIENDA
        Else
            Me.CboTipoDocumento.WriteOn IDTipoDocumento
        End If
    
    Else
        If NumeroOrdiniDaEvadere = 1 Then
            Me.cboPagamento.Enabled = True
            Me.cboTipoTrasporto.Enabled = True
            Me.cboVettore.Enabled = True
            Me.txtTargaAutomezzo.Enabled = True
            Me.txtIstruzioniMittente.Enabled = True
            IDTipoDocumento = GET_TIPO_DOCUMENTO_CLIENTE(TheApp.IDUser, Me.CboTipoDocumento.CurrentID)
            If IDTipoDocumento = 0 Then
                GET_DEFAULT_AZIENDA
            Else
                Me.CboTipoDocumento.WriteOn IDTipoDocumento
            End If
        End If
        
        If NumeroOrdiniDaEvadere = 0 Then
            Me.cboPagamento.Enabled = True
            Me.cboTipoTrasporto.Enabled = True
            Me.cboVettore.Enabled = True
            Me.txtTargaAutomezzo.Enabled = True
            Me.txtIstruzioniMittente.Enabled = True
            
            GET_DEFAULT_AZIENDA
        End If
    End If
    
    CboTipoDocumento_Click
End Sub
Private Sub sbSelectSelectedRow(ByVal Selected As Boolean, Griglia As Integer)
On Error GoTo ERR_sbSelectSelectedRow

    If AVVIA_FATTURAZIONE = 1 Then Exit Sub

    If Not rsGrigliaOrdini.EOF And Not rsGrigliaOrdini.BOF Then
        If fnNotNullN(rsGrigliaOrdini!IDUtente) > 0 Then
            If fnNotNullN(rsGrigliaOrdini!IDUtente) <> TheApp.IDUser Then
                MsgBox "Ordine bloccato dall'utente " & fnNotNull(rsGrigliaOrdini!Utente)
                Exit Sub
            End If
        End If
        If Selected = False Then
           
            rsGrigliaOrdini.Fields("IDUtente").Value = 0
            rsGrigliaOrdini.Fields("Utente").Value = ""
        Else
            rsGrigliaOrdini.Fields("IDUtente").Value = TheApp.IDUser
            rsGrigliaOrdini.Fields("Utente").Value = TheApp.User
        End If
        
        rsGrigliaOrdini.Fields("DaRegistrare").Value = Abs(CLng(Selected))
        If Abs(rsGrigliaOrdini.Fields("DaRegistrare").Value) = 1 Then
            Testo = GET_MESSAGGIO_PER_ANAGRAFICA(rsGrigliaOrdini.Fields("IDCliente").Value)
            If Len(Trim(Testo)) > 0 Then
                MsgBox Testo, vbInformation, "Messaggio importante"
            End If
        End If
        
        rsGrigliaOrdini.UpdateBatch
                
        GrigliaOrdini.Refresh
    End If
Exit Sub
ERR_sbSelectSelectedRow:
    rsGrigliaOrdini.Requery
    GrigliaOrdini.Refresh
End Sub

Private Function GET_TIPO_OGGETTO(NomeGestore) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT TipoOggetto.IDTipoOggetto, Gestore.Gestore "
sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(NomeGestore) & ")"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_OGGETTO = 0
Else
    GET_TIPO_OGGETTO = fnNotNullN(rs!IDTipoOggetto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_FUNZIONE(IDTipoOggetto) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione "
sSQL = sSQL & "FROM Funzione  "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE = 0
Else
    GET_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_PARAMETRI_LIQUIDAZIONE_ARTICOLO(IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDUnitaDiMisuraLiq, RV_POQuantitaPerCollo, RV_POMoltiplicatore "
sSQL = sSQL & "FROM Articolo  "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    LINK_UM_LIQ = 0
    QUANTITA_PER_COLLI = 1
    Moltiplicatore = 1
Else
    LINK_UM_LIQ = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
    
    If fnNotNullN(rs!RV_POQuantitaPerCollo) = 0 Then
        QUANTITA_PER_COLLI = 1
    Else
        QUANTITA_PER_COLLI = fnNotNullN(rs!RV_POQuantitaPerCollo)
    End If
    If fnNotNullN(rs!RV_POMoltiplicatore) = 0 Then
        Moltiplicatore = 1
    Else
        Moltiplicatore = fnNotNullN(rs!RV_POMoltiplicatore)
    End If
    
End If

rs.CloseResultset
Set rs = Nothing

End Sub

Private Sub GET_IVA_ARTICOLO(IDArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Iva.Iva, Iva.AliquotaIva, Iva.Codice, Iva.IDIva "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    Link_IVAArticolo = 0
    AliquotaIvaArticolo = 0
Else
    Link_IVAArticolo = fnNotNullN(rs!IDIva)
    AliquotaIvaArticolo = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_ALIQUOTA_IVA_ARTICOLO(IDIva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AliquotaIva FROM Iva "
sSQL = sSQL & " WHERE IDIva=" & IDIva


Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_ALIQUOTA_IVA_ARTICOLO = 0
Else
    GET_ALIQUOTA_IVA_ARTICOLO = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_IDOggetto(NumeroDocumento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto FROM Oggetto "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDSezionale=" & Me.cboSezionale.CurrentID
sSQL = sSQL & " AND Numero=" & NumeroDocumento
sSQL = sSQL & " AND DataEmissione=" & fnNormDate(Me.txtDataDoc.Text)
sSQL = sSQL & " AND IDTipoOggetto=" & Me.CboTipoDocumento.CurrentID


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_IDOggetto = 0
Else
    GET_IDOggetto = fnNotNullN(rs!IDOggetto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub fnGrigliaFattureCreate()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim oItem As MSComctlLib.ListItem

Me.LVFattureCreate.ListItems.Clear

sSQL = "SELECT * FROM RV_POTMPOrdiniFatturati "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & "ORDER BY NumeroDocumento DESC"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    Set oItem = Me.LVFattureCreate.ListItems.Add
    
    'Popola l 'item della listview
    oItem.Text = fnNotNullN(rs!IDOggetto)
    oItem.SubItems(1) = fnNotNullN(rs!IDTipoOggetto)
    oItem.SubItems(2) = fnNotNull(rs!DataDocumento)
    oItem.SubItems(3) = fnNotNull(rs!NumeroDocumento)
    oItem.SubItems(4) = fnNotNull(rs!Cliente)
    oItem.SubItems(5) = fnNotNull(rs!SitoPerAnagrafica)
    oItem.SubItems(6) = fnNotNull(rs!Vettore)

rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
DoEvents
If Me.LVFattureCreate.ListItems.Count > 0 Then

    Set LVFattureCreate.SelectedItem = LVFattureCreate.ListItems(LVFattureCreate.ListItems.Count)
    LVFattureCreate.SelectedItem.EnsureVisible
    
Else
    Me.LabelLink1.Visible = False
    
End If
End Sub

Private Function GET_CLIENTE(IDCliente As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Anagrafica FROM IERepCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CLIENTE = ""
Else
    GET_CLIENTE = fnNotNull(rs!Anagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CLIENTE_DESTINAZIONE(IDSitoPerAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SitoPerAnagrafica FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoPerAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CLIENTE_DESTINAZIONE = ""
Else
    GET_CLIENTE_DESTINAZIONE = fnNotNull(rs!SitoPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_VETTORE(IDVettore As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Vettore FROM Vettore "
sSQL = sSQL & "WHERE IDVettore=" & IDVettore

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_VETTORE = ""
Else
    GET_VETTORE = fnNotNull(rs!Vettore)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_DOCUMENTO_DEFAULF_CLIENTE(IDCliente As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoOggettoDocEvasione FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_OGGETTO_PER_FATT Me.CboTipoDocumento.ItemData(Me.CboTipoDocumento.ListIndex)
Else
    If fnNotNullN(rs!IDTipoOggettoDocEvasione) = 0 Then
        GET_TIPO_OGGETTO_PER_FATT Me.CboTipoDocumento.ItemData(Me.CboTipoDocumento.ListIndex)
    Else
        Select Case fnNotNullN(rs!IDTipoOggettoDocEvasione)
            Case 2
                GET_TIPO_OGGETTO_PER_FATT 1
            Case 114
                GET_TIPO_OGGETTO_PER_FATT 2
            Case 8
                GET_TIPO_OGGETTO_PER_FATT 3
            Case Else
                GET_TIPO_OGGETTO_PER_FATT Me.CboTipoDocumento.ItemData(Me.CboTipoDocumento.ListIndex)
        End Select
    End If
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Public Sub fnChiudiOrdiniElaborati()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE DaRegistrare=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
    sSQL = sSQL & "Doc_ordine_chiuso=1 "
    sSQL = sSQL & " WHERE IDOggetto=" & fnNotNullN(rs!IDOggetto)
    CnDMT.Execute sSQL
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Public Function GET_IDSEZIONALE_TIPOOGGETTO(IDTipoOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select IDSezionale "
    sSQL = sSQL & "FROM DefaultFilialePerTipoOggetto "
    sSQL = sSQL & " WHERE IDFiliale = " & TheApp.Branch
    sSQL = sSQL & " AND IDTipoOggetto = " & IDTipoOggetto
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_IDSEZIONALE_TIPOOGGETTO = fnNotNullN(rs!IDSezionale)
    Else
        GET_IDSEZIONALE_TIPOOGGETTO = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End Function
Private Function GET_DESCRIZIONE_TIPOOGGETTO(IDTipoOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select Oggetto "
    sSQL = sSQL & "FROM TipoOggetto "
    sSQL = sSQL & "WHERE IDTipoOggetto = " & IDTipoOggetto
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_DESCRIZIONE_TIPOOGGETTO = fnNotNull(rs!Oggetto)
    Else
        GET_DESCRIZIONE_TIPOOGGETTO = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function GET_PREZZO_PEDANA(IDArticolo As Long) As Double
Dim sSQL As String
Dim IDListinoDefault As Long
Dim rs As DmtOleDbLib.adoResultset




rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LISTINO_DEFAULT(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListinoDefault "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    rs.CloseResultset
    Set rs = Nothing
    
    sSQL = "SELECT IDListinoDiBase "
    sSQL = sSQL & "FROM ConfigurazioneVendite "
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_LISTINO_DEFAULT = 0
    Else
        GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
    End If
Else
    If fnNotNullN(rs!IDListinoDefault) = 0 Then
        rs.CloseResultset
        Set rs = Nothing
        
        sSQL = "SELECT IDListinoDiBase "
        sSQL = sSQL & "FROM ConfigurazioneVendite "
        sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
        
        Set rs = CnDMT.OpenResultset(sSQL)
        
        If rs.EOF Then
            GET_LISTINO_DEFAULT = 0
        Else
            GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
        End If
    Else
        GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDefault)
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListinoImballiDefault "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = 0
Else
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = fnNotNullN(rs!IDListinoImballiDefault)
    
End If
rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_PREZZO_IMBALLO(IDListino As Long, IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoNettoIVA "
sSQL = sSQL & "FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDListino=" & IDListino
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_IMBALLO = 0
Else
    GET_PREZZO_IMBALLO = fnNotNullN(rs!PrezzoNettoIva)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function sbCalcolaImportoLiquidazione(LINK_UM_COOP As Long, IDListino As Long, Qta_UM As Double, Colli As Double, IDImballo As Long, ImportoUnitario As Double, ImportoUnitarioImballo As Double, QTA_LIQ As Double) As Double
Dim ImportoImballo As Double
Dim ImportoImballoUnitario As Double



If ImportoUnitarioImballo = 0 Then
    ImportoImballo = GET_PREZZO_IMBALLO(IDListino, IDImballo)
Else
    ImportoImballo = ImportoUnitarioImballo
End If

If Qta_UM > 0 Then
    ImportoImballoUnitario = (ImportoImballo * Colli) / Qta_UM

    sbCalcolaImportoLiquidazione = ImportoImballoUnitario

Else
    sbCalcolaImportoLiquidazione = 0
End If

If LINK_UM_LIQ = LINK_UM_COOP Then
    If Moltiplicatore > 0 Then
        sbCalcolaImportoLiquidazione = sbCalcolaImportoLiquidazione / Moltiplicatore
    Else
        sbCalcolaImportoLiquidazione = sbCalcolaImportoLiquidazione
    End If
Else
    If QTA_LIQ > 0 Then
        sbCalcolaImportoLiquidazione = (sbCalcolaImportoLiquidazione * Qta_UM) / QTA_LIQ
    Else
        sbCalcolaImportoLiquidazione = sbCalcolaImportoLiquidazione
    End If
End If

If sbCalcolaImportoLiquidazione > 0 Then
    sbCalcolaImportoLiquidazione = sbCalcolaImportoLiquidazione
End If
End Function
Private Function GET_VARIAZIONE_PREZZO_IMBALLO(LINK_UM_COOP As Long, IDListino As Long, Qta_UM As Double, Colli As Double, IDImballo As Long, ImportoUnitario As Double, ImportoUnitarioImballo As Double, QTA_LIQ As Double) As Double
Dim ImportoImballo As Double
Dim ImportoImballoUnitario As Double

If ImportoUnitarioImballo = 0 Then
    ImportoImballo = GET_PREZZO_IMBALLO(IDListino, IDImballo)
Else
    ImportoImballo = ImportoUnitarioImballo
End If

If Qta_UM > 0 Then
    ImportoImballoUnitario = (ImportoImballo * Colli) / Qta_UM

    GET_VARIAZIONE_PREZZO_IMBALLO = ImportoImballoUnitario
Else
    GET_VARIAZIONE_PREZZO_IMBALLO = 0
End If

End Function

Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
    sSQL = sSQL & "IDUnitaDiMisura = " & Link_UMAcq
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetUMCoop = rs!RV_POIDUnitaDiMisuraCoop
    Else
        fnGetUMCoop = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Function GET_TIPO_SCONTO(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTipoScontoPerCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_SCONTO = 1
Else
    If fnNotNullN(rs!IDRV_POTipoCalcoloScontoDoc) <= 1 Then
        GET_TIPO_SCONTO = 1
    Else
        GET_TIPO_SCONTO = 2
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_SCONTO_ABITUALE(IDAnagrafica As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ScontoAbituale FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_SCONTO_ABITUALE = 0
Else
    If fnNotNullN(rs!ScontoAbituale) = 0 Then
        GET_SCONTO_ABITUALE = 0
    Else
        GET_SCONTO_ABITUALE = fnNotNullN(rs!ScontoAbituale)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI(IDTipoOggetto As Long, IDOggetto As Long, IDDestinazione As Long)
On Error GoTo ERR_AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim QuantitaMovimento As Double
Dim Link_Unita_di_Misura_Conferimeto As Long


sSQL = "SELECT " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".IDValoriOggettoDettaglio, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".Art_pre_uni_net_sco_net_IVA, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".Link_art_articolo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".Art_quantita_pezzi, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".Art_numero_colli, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".Art_tara, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".Art_peso, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PODataConferimento, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDConferimentoRighe, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POCodiceLotto, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POImportoLiq, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDSocio, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POLottoCampagna, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POImportoDaLiq, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POQuantitaLiq, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDAssegnazioneMerce, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDProcessoIVGamma, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POPrezzoMedioInLiq, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDImballo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POImportoUnitarioImballo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POImportoMerceNetta, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDIvaImballo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POVariazionePrezzoImballo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POImportoImballoInArticolo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDAnagraficaFatturazione, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoImportoVenditaLiq, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POVariazionePrezzoManuale, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoDocumentoCoop, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POImportoRigaCommissioni, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PODataLavorazione, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoLavorazione, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoCategoria, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDCalibro, "

sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDPedana, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POCodicePedana, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POPesoPedana, "

sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PORigaRiscontroPeso, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POAnnotazioniAggiuntiveLav, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PONotaRigaOrdRaggr, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PODataOrdineCliente, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PONumeroOrdineCliente, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PODataOrdineInterno, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PONumeroOrdineInterno, "

sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDImballoPrim, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PONumeroConfezioniPerImballo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POTaraConfezioneImballo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POCodiceImballoPrim, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_PODescrizioneImballoPrim, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POCostoConfezioneImballo, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POCostoKitLiq, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POCostoConfezioneImballoLiq, "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDLottoCampagnaLavorazione, "

sSQL = sSQL & "RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceTesta.NumeroDocumento, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDUnitaDiMisura, RV_POCaricoMerceRighe.CodiceLotto, "
sSQL = sSQL & "RV_POCaricoMerceTesta.IDMagazzinoConferimento, RV_POCaricoMerceRighe.IDUnitaDiMisuraDiamante, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDRV_POTipoLavorazione AS IDTipoLavorazioneConf, RV_POCaricoMerceRighe.PrezzoMedio AS PrezzoMedioConf "
sSQL = sSQL & "FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON "
sSQL = sSQL & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDConferimentoRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".IDOggetto=" & IDOggetto
sSQL = sSQL & " AND " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POTipoRiga=1"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection
    
While Not rs.EOF
    Select Case fnNotNullN(rs!IDUnitaDiMisura)
        Case 1
            QuantitaMovimento = fnNotNullN(rs!Art_numero_colli)
        Case 2
            QuantitaMovimento = fnNotNullN(rs!Art_peso)
        Case 3
            QuantitaMovimento = fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)
        Case 4
            QuantitaMovimento = fnNotNullN(rs!Art_tara)
        Case 5
            QuantitaMovimento = fnNotNullN(rs!Art_quantita_pezzi)
    End Select
    
    Aggiorna_Movimento_Documento fnNotNullN(rs!Link_Art_articolo), fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNullN(rs!NumeroDocumento), _
    fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
    fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), QuantitaMovimento, _
    fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)), _
    fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POPrezzoMedioInLiq), fnNotNullN(rs!RV_POIDImballo), fnNotNullN(rs!RV_POImportoUnitarioImballo), _
    fnNotNullN(rs!RV_POIDIvaImballo), fnNotNullN(rs!RV_POVariazionePrezzoImballo), fnNotNullN(rs!RV_POImportoMerceNetta), fnNotNullN(rs!RV_POImportoImballoInArticolo), fnNotNullN(rs!RV_POIDAnagraficaFatturazione), IDDestinazione, fnNotNullN(rs!RV_POIDTipoImportoVenditaLiq), _
    fnNotNullN(rs!RV_POIDTipoDocumentoCoop), fnNotNullN(rs!RV_POVariazionePrezzoManuale), fnNotNull(rs!RV_PODataLavorazione), _
    fnNotNullN(rs!RV_POIDTipoLavorazione), fnNotNullN(rs!RV_POIDTipoCategoria), fnNotNullN(rs!RV_POIDCalibro), fnNotNullN(rs!IDTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), _
    fnNotNullN(rs!RV_POIDPedana), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNull(rs!RV_POCodicePedana), fnNotNullN(rs!RV_POPesoPedana), fnNotNullN(rs!RV_POImportoRigaCommissioni), _
    fnNotNull(rs!RV_POAnnotazioniAggiuntiveLav), fnNotNull(rs!RV_PONotaRigaOrdRaggr), fnNotNull(rs!RV_PODataOrdineCliente), fnNotNull(rs!RV_PONumeroOrdineCliente), fnNotNull(rs!RV_PODataOrdineInterno), fnNotNull(rs!RV_PONumeroOrdineInterno), _
    fnNotNullN(rs!RV_POIDImballoPrim), fnNotNull(rs!RV_POCodiceImballoPrim), fnNotNull(rs!RV_PODescrizioneImballoPrim), fnNotNullN(rs!RV_PONumeroConfezioniPerImballo), fnNotNullN(rs!RV_POTaraConfezioneImballo), _
    fnNotNullN(rs!RV_POCostoConfezioneImballo), fnNotNullN(rs!RV_POCostoConfezioneImballoLiq), fnNotNullN(rs!RV_POCostoKitLiq), fnNotNullN(rs!RV_POIDLottoCampagnaLavorazione)
    
rs.MoveNext
Wend

rs.Close
Set rs = Nothing
Exit Sub
ERR_AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI:
    MsgBox Err.Description, vbCritical, "ERR_AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI"
End Sub
Private Function Aggiorna_Movimento_Documento(IDArticolo As Long, ImportoUnitarioArticolo As Double, IDRiga As Long, IDRigaConferimento, IDAssegnazione As Long, IDProcessoIVGamma As Long, IDSocio As Long, _
DataConferimento As String, NumeroConferimento As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, CodiceLottoVendita As String, _
QuantitaLiquidazione As Double, ImportoInclusoImballo As Double, ImportoLiquidazione As Double, QuantitaMovimentata As Double, Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, _
PrezzoMedioLiq As Double, IDArticoloImballo As Long, ImportoUnitarioImballo, IDIvaImballo As Long, VariazionePrezzoImballo As Double, PrezzoMerceNetta As Double, MerceInclusaImballo As Long, IDAnagraficaFatturazione As Long, _
IDSitoPerAnagrafica As Long, IDTipoImportoLiq As Long, IDTipoDocumentoCoop As Long, VarImpLiqMan As Double, DataLavorazione As String, IDTipoLavorazione As Long, IDTipoCategoria As Long, IDCalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDPedana As Long, IDTipoPedana As Long, CodicePedana As String, PesoPedana As Double, ImportoRigaCommissioni As Double, _
AnnotazioniAggiuntive As String, RaggrOrdine As String, DataOrdineCliente As String, NumeroOrdineCliente As String, DataOrdineInterno As String, NumeroOrdineInterno As String, _
IDImballoPrimario As Long, CodiceImballoPrimario As String, DescrizioneImballoPrimario As String, NConfezioniImballo As Double, TaraConfezione As Double, _
CostoConfezioneImballo As Double, CostoConfezioneImballoLiq As Double, CostoKitLiq As Double, IDLottoProduzioneLavorazione As Long) As Long

Dim Prezzo As Double
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim Moltiplicatore As Double
Dim MerceNettaPerLiquidazione As Double

    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(IDArticolo)
    
    sSQL = "SELECT * FROM Movimento "
    sSQL = sSQL & "WHERE IDTipoOggetto=" & ObjDoc.IDTipoOggetto
    sSQL = sSQL & " AND IDOggetto=" & ObjDoc.IDOggetto
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDRiga
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    If Not rs.EOF Then
        
        rs("RV_POTipoRiga").Value = 1
        rs("RV_POIDCaricoMerceRighe").Value = IDRigaConferimento
        rs("RV_POIDAssegnazioneMerce").Value = IDAssegnazione
        rs("RV_POIDProcessoIVGamma") = IDProcessoIVGamma
        rs("RV_POIDAnagraficaSocio") = IDSocio
        If Len(DataConferimento) > 0 Then
            rs("RV_PODataConferimento") = DataConferimento
        End If
        rs("RV_PONumeroConferimento") = NumeroConferimento
        rs("RV_POCodiceLotto") = CodiceLottoEntrata
        rs("RV_POCodiceLottoCampagna") = CodiceLottoCampagna
        rs("RV_POCodiceLottoVendita") = CodiceLottoVendita
        rs("RV_POQuantitaLiquidazione") = QuantitaLiquidazione
        rs("RV_POImportoInclusoImballo") = ImportoInclusoImballo
        rs("RV_POPrezzoMerceNetta").Value = PrezzoMerceNetta
        
'        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Prezzo = ImportoUnitarioArticolo
'
'        Prezzo = Prezzo / Moltiplicatore
'        MerceNettaPerLiquidazione = fnNotNullN(rs("RV_POPrezzoMerceNetta").Value) / Moltiplicatore
'        If ImportoInclusoImballo > 0 Then
'            Prezzo = Prezzo + ImportoInclusoImballo
'        Else
'            Prezzo = Prezzo - Abs(ImportoInclusoImballo)
'        End If
'
'
'        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto, Moltiplicatore, rs, MerceNettaPerLiquidazione)
'        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
        rs("RV_POImportoRigaCommissioni") = ImportoRigaCommissioni
        rs("RV_POImportoLiquidazione") = ImportoLiquidazione
        rs("RV_POQuantitaMovimentata") = QuantitaMovimentata
        rs("RV_PONumeroColli") = Colli
        rs("RV_POPesoLordo") = PesoLordo
        rs("RV_POPesoNetto") = PesoLordo - Tara
        rs("RV_POTara") = Tara
        rs("RV_POQuantitaPezzi") = Pezzi
        
        rs("RV_POPrezzoMedioInLiq").Value = PrezzoMedioLiq
        rs("RV_POIDImballo").Value = IDArticoloImballo
        rs("RV_POImportoUnitarioImballo").Value = ImportoUnitarioImballo
        rs("RV_POIDIvaImballo").Value = IDIvaImballo
        rs("RV_POVariazionePrezzoImballo").Value = VariazionePrezzoImballo
        rs("RV_POQuantitaLiqPerPrezzoMedio").Value = QuantitaLiquidazione
        rs("RV_POMerceInclusaImballo").Value = MerceInclusaImballo
        rs("RV_POTipoRigaCollegata").Value = 0
        rs("RV_POIDAnagraficaFatturazione").Value = IDAnagraficaFatturazione
        rs("RV_POIDSitoPerAnagrafica").Value = IDSitoPerAnagrafica
        rs("RV_POIDTipoImportoVenditaLiq").Value = IDTipoImportoLiq
        rs("Oggetto").Value = GET_DESCRIZIONE_TIPOOGGETTO(ObjDoc.IDTipoOggetto)
        
        rs("RV_POVariazionePrezzoManuale").Value = VarImpLiqMan
        rs("RV_POIDTipoDocumentoCoop").Value = IDTipoDocumentoCoop
        
        rs("RV_PODataLavorazione").Value = DataLavorazione
        rs("RV_POIDTipoLavorazione").Value = IDTipoLavorazione
        rs("RV_POIDCalibro").Value = IDCalibro
        rs("RV_POIDTipoCategoria").Value = IDTipoCategoria
        rs("RV_POIDTipoLavorazioneConf").Value = IDTipoLavorazioneConf
        rs("RV_POPrezzoMedioConf").Value = PrezzoMedioConf
        
        rs("RV_POIDPedana").Value = IDPedana
        rs("RV_POIDTipoPedana").Value = IDTipoPedana
        rs("RV_POCodicePedana").Value = CodicePedana
        rs("RV_POPesoPedana").Value = PesoPedana

        rs("RV_PORigaRiscontroPeso") = 0
        rs("RV_POAnnotazioniAggiuntiveLav").Value = AnnotazioniAggiuntive
        rs("RV_PONotaRigaOrdRaggr").Value = RaggrOrdine
        
        If Len(DataOrdineCliente) > 0 Then
            rs("RV_PODataOrdineCliente").Value = DataOrdineCliente
        End If
        rs("RV_PONumeroOrdineCliente").Value = NumeroOrdineCliente
        
        If Len(DataOrdineInterno) > 0 Then
            rs("RV_PODataOrdineInterno").Value = DataOrdineInterno
        End If
        
        rs("RV_PONumeroOrdineInterno").Value = NumeroOrdineInterno
        rs("RV_POIDImballoPrim").Value = IDImballoPrimario
        rs("RV_POCodiceImballoPrim").Value = CodiceImballoPrimario
        rs("RV_PODescrizioneImballoPrim").Value = DescrizioneImballoPrimario
        rs("RV_PONumeroConfezioniPerImballo").Value = NConfezioniImballo
        rs("RV_POTaraConfezioneImballo").Value = TaraConfezione
        rs("RV_POCostoConfezioneImballo").Value = CostoConfezioneImballo
        rs("RV_POCostoConfezioneImballoLiq").Value = CostoConfezioneImballoLiq
        rs("RV_POCostoKitLiq").Value = CostoKitLiq
        rs("RV_POImportoLiqDoc") = ImportoLiquidazione
        rs("RV_POIDLottoCampagnaLavorazione").Value = IDLottoProduzioneLavorazione
        rs("RV_PODataCompetenzaLiq").Value = ObjDoc.DataEmissione
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
End Function
Private Function GET_COMMISSIONI_DOCUMENTO(PrezzoLiquidazione As Double, IDOggetto As Long, IDTipoOggetto As Long, Moltiplicatore As Double, rstmp As ADODB.Recordset, ImportoMerceRiga, PrezzoScontatoRiga As Double, PrezzoLiquidazioneScontato As Double) As Double
Dim sSQL As String
Dim rscomm As DmtOleDbLib.adoResultset
Dim ImportoRigaCommissioni As Double

GET_COMMISSIONI_DOCUMENTO = 0
ImportoRigaCommissioni = 0

sSQL = "SELECT RV_POCommissioniPerDoc.IDRV_POCommissioniPerDoc, RV_POCommissioniPerDoc.IDOggetto, RV_POCommissioniPerDoc.IDRV_POTipoCommissione, RV_POCommissioniPerDoc.Percentuale, "
sSQL = sSQL & "RV_POCommissioniPerDoc.Importo, RV_POCommissioniPerDoc.ImportoRiga, RV_POCommissioniPerDoc.Quantita, RV_POCommissioniPerDoc.APercentuale, RV_POCommissioniPerDoc.ImportoTotale,"
sSQL = sSQL & "RV_POCommissioniPerDoc.IDRV_POTipoPedana , RV_POCommissioniPerDoc.PercentualeDaCommissione, RV_POCommissioniPerDoc.IDArticoloImballo, RV_POTipoCommissione.IDRV_POTipoValoreDocumento "
sSQL = sSQL & "FROM RV_POCommissioniPerDoc INNER JOIN "
sSQL = sSQL & "RV_POTipoCommissione ON RV_POCommissioniPerDoc.IDRV_POTipoCommissione = RV_POTipoCommissione.IDRV_POTipoCommissione "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND ((IDRV_POTipoPedana=0) OR (IDRV_POTipoPedana IS NULL))"
sSQL = sSQL & " AND ((APercentuale=0) OR (APercentuale IS NULL))"

Set rscomm = CnDMT.OpenResultset(sSQL)

While Not rscomm.EOF
    If (fnNotNullN(rscomm!IDRV_POTipoValoreDocumento) < 5) Then
        ImportoRigaCommissioni = ImportoRigaCommissioni + ((ImportoMerceRiga / 100) * fnNotNullN(rscomm!Percentuale))
        ImportoRigaCommissioni = ImportoRigaCommissioni + (fnNotNullN(rscomm!Importo))
    
        GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + ((PrezzoLiquidazione / 100) * fnNotNullN(rscomm!Percentuale))
        GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + (fnNotNullN(rscomm!Importo))
    Else
        ImportoRigaCommissioni = ImportoRigaCommissioni + ((PrezzoScontatoRiga / 100) * fnNotNullN(rscomm!Percentuale))
        ImportoRigaCommissioni = ImportoRigaCommissioni + (fnNotNullN(rscomm!Importo))
    
        GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + ((PrezzoLiquidazioneScontato / 100) * fnNotNullN(rscomm!Percentuale))
        GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + (fnNotNullN(rscomm!Importo))
    End If
rscomm.MoveNext
Wend

rscomm.CloseResultset
Set rscomm = Nothing

rstmp!RV_POImportoRigaCommissioni = ImportoRigaCommissioni
GET_COMMISSIONI_DOCUMENTO = PrezzoLiquidazione - GET_COMMISSIONI_DOCUMENTO

End Function

Private Function GET_MOLTIPLICATORE_ARTICOLO(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POMoltiplicatore FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_MOLTIPLICATORE_ARTICOLO = 1
Else
    If fnNotNullN(rs!RV_POMoltiplicatore) = 0 Then
        GET_MOLTIPLICATORE_ARTICOLO = 1
    Else
        GET_MOLTIPLICATORE_ARTICOLO = fnNotNullN(rs!RV_POMoltiplicatore)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_IMPORTO_SPESE_TRASPORTO(IDCliente As Long, IDOggettoOrdine As Long, IDDestinazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCli As ADODB.Recordset


''''''''''''''''''''''''''''''''SPESE DI TRASPORTO DELLE COMMISSIONI DELLE PEDANE''''''''''''''''''''''''''''
sSQL = "SELECT IDArticoloImballo, Articolo, CodiceArticolo, COUNT(IDRV_POPedana) AS QuantitaPedana "
sSQL = sSQL & "FROM dbo.RV_POIEArticoloPedanaOrdine "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine
sSQL = sSQL & "GROUP BY IDArticoloImballo, Articolo, CodiceArticolo "

Set rs = CnDMT.OpenResultset(sSQL)

GET_IMPORTO_SPESE_TRASPORTO = 0

While Not rs.EOF
    sSQL = "SELECT * FROM RV_POConfigurazioneClienteTrasporto "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDAnagrafica=" & IDCliente
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDDestinazione
    sSQL = sSQL & " AND IDArticolo=" & fnNotNullN(rs!IDArticoloImballo)
    sSQL = sSQL & " AND IDRV_POTipoCommissione=0"
    Set rsCli = New ADODB.Recordset
    
    rsCli.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rsCli.EOF
        GET_IMPORTO_SPESE_TRASPORTO = GET_IMPORTO_SPESE_TRASPORTO + (fnNotNullN(rsCli!PrezzoTrasporto) * (fnNotNullN(rs!QuantitaPedana) / fnNotNullN(rsCli!Quantita)))
    rsCli.MoveNext
    Wend
    
    rsCli.Close
    Set rsCli = Nothing
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''SPESE DI TRASPORTO DEGLI IMBALLI'''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDImballoVendita AS IDArticoloImballo, SUM(Colli) AS QuantitaPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine = " & IDOggettoOrdine
sSQL = sSQL & " GROUP BY IDImballoVendita"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "SELECT * FROM RV_POConfigurazioneClienteTrasporto "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDAnagrafica=" & IDCliente
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDDestinazione
    sSQL = sSQL & " AND IDArticolo=" & fnNotNullN(rs!IDArticoloImballo)
    sSQL = sSQL & " AND IDRV_POTipoCommissione=0"
    Set rsCli = New ADODB.Recordset
    
    rsCli.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rsCli.EOF
        GET_IMPORTO_SPESE_TRASPORTO = GET_IMPORTO_SPESE_TRASPORTO + (fnNotNullN(rsCli!PrezzoTrasporto) * (fnNotNullN(rs!QuantitaPedana) / fnNotNullN(rsCli!Quantita)))
    rsCli.MoveNext
    Wend
    
    rsCli.Close
    Set rsCli = Nothing
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Function
Private Sub IMPOSTA_COMMISSIONI_PER_CLIENTE(IDCliente As Long, IDOggettoOrdine As Long, IDOggettoDocumento As Long, IDSitoPerAnagrafica As Long)
On Error GoTo ERR_IMPOSTA_COMMISSIONI_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim rsCli As ADODB.Recordset
Dim SpesaTrasporto As Double
Dim Totale_merce_lavorato As Double
Dim Totale_merce_lavorato_lordo As Double
Dim Totale_documento_netto_iva As Double
Dim Totale_documento_lordo_iva As Double
Dim Totale_Importo As Double
Dim AvviaInserimentoCommissione As Long

Totale_merce_lavorato = GET_TOTALE_MERCE_DOCUMENTO(IDOggettoDocumento)
Totale_merce_lavorato_lordo = GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO()
Totale_documento_netto_iva = fnNotNullN(ObjDoc.Field("Tot_imponibile_corr", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
Totale_documento_lordo_iva = fnNotNullN(ObjDoc.Field("Tot_documento_corr", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))


'''''''''''''''''''''''''''COMMISSIONI PER CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIECommissioniPerCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

Set rsNew = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
        rsNew!IDOggetto = IDOggettoDocumento
        rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
        rsNew!Percentuale = fnNotNullN(rs!Percentuale)
        rsNew!PercentualeDaCommissione = fnNotNullN(rs!Percentuale)
        rsNew!Importo = 0
        rsNew!ImportoTotale = 0
        rsNew!Quantita = 1
        rsNew!APercentuale = 0
        
        Select Case fnNotNullN(rs!IDRV_POTipoValoreDocumento)
            
            Case 1
                Totale_Importo = Totale_merce_lavorato
            Case 2
                Totale_Importo = Totale_merce_lavorato_lordo
            Case 3
                Totale_Importo = Totale_documento_netto_iva
            Case 4
                Totale_Importo = Totale_documento_lordo_iva
            Case Else
                Totale_Importo = Totale_merce_lavorato
            
        End Select
        
        If Totale_Importo = 0 Then
            rsNew!ImportoRiga = 0
        Else
            If fnNotNullN(rs!IDRV_POTipoValoreDocumento) <= 1 Then
                rsNew!ImportoRiga = (Totale_merce_lavorato / 100) * fnNotNullN(rs!Percentuale)
            Else
                rsNew!ImportoRiga = (Totale_Importo / 100) * fnNotNullN(rs!Percentuale)
                If (Totale_merce_lavorato > 0) Then
                    rsNew!Percentuale = (rsNew!ImportoRiga / Totale_merce_lavorato) * 100
                Else
                    rsNew!Percentuale = 0
                End If
            End If
        End If
        
    rsNew.Update
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Totale_merce_lavorato = 0 Then
    rsNew.Close
    Set rsNew = Nothing
Exit Sub
End If

'''''''''''''''''''''''''''''''COMMISSIONI PER TIPO PEDANA''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT Link_Art_articolo AS IDArticoloImballo, SUM(Art_quantita_totale) AS QuantitaPedana "
sSQL = sSQL & "FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
sSQL = sSQL & " WHERE IDOggetto =" & ObjDoc.IDOggetto
sSQL = sSQL & " AND RV_POTipoRiga = 2 "
sSQL = sSQL & "GROUP BY Link_Art_articolo"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "SELECT * FROM RV_POIEConfigurazioneClienteTrasporto "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDAnagrafica=" & IDCliente
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDSitoPerAnagrafica
    sSQL = sSQL & " AND IDArticolo=" & fnNotNullN(rs!IDArticoloImballo)
    sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
    sSQL = sSQL & " AND ((CommissionePerPedana=0) OR (CommissionePerPedana IS NULL))"
    
    Set rsCli = New ADODB.Recordset
    
    rsCli.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rsCli.EOF
        rsNew.AddNew
            rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
            rsNew!IDOggetto = IDOggettoDocumento
            rsNew!IDRV_POTipoCommissione = fnNotNullN(rsCli!IDRV_POTipoCommissione)
            SpesaTrasporto = (fnNotNullN(rsCli!PrezzoTrasporto) * (fnNotNullN(rs!QuantitaPedana) / fnNotNullN(rsCli!Quantita)))
            rsNew!Percentuale = (SpesaTrasporto / Totale_merce_lavorato) * 100
            rsNew!Importo = 0
            rsNew!ImportoRiga = SpesaTrasporto
            rsNew!IDArticoloImballo = fnNotNullN(rs!IDArticoloImballo)
        rsNew.Update
    rsCli.MoveNext
    Wend
    
    rsCli.Close
    Set rsCli = Nothing
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = ""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_IMPOSTA_COMMISSIONI_PER_CLIENTE:
    MsgBox Err.Description, vbCritical, "IMPOSTA_COMMISSIONI_PER_CLIENTE"

End Sub


Private Sub IMPOSTA_COMMISSIONI_PER_TIPO_PEDANA(IDCliente As Long, IDOggettoOrdine As Long, IDOggettoDocumento As Long, IDSitoPerAnagrafica As Long)
On Error GoTo ERR_IMPOSTA_COMMISSIONI_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim rsCli As ADODB.Recordset
Dim rsPed As ADODB.Recordset
Dim rsQuantita As ADODB.Recordset

Dim SpesaTrasporto As Double
Dim Totale_merce_lavorato As Double


If fnNotNullN(ObjDoc.Field("Link_Nom_porto", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
    If (fnNotNullN(ObjDoc.Field("Link_Nom_porto", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) = LINK_PORTO_NO_TRASP) Then Exit Sub
End If

sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
sSQL = sSQL & "WHERE IDOggetto=" & ObjDoc.IDOggetto

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic


'''''''''''''''''''''''''''''''COMMISSIONI PER TIPO PEDANA''''''''''''''''''''''''''''''''''''''''''

'sSQL = "SELECT IDOggetto, RV_POIDTipoPedana, RV_POTipoPedana.IDArticoloImballo, "
'sSQL = sSQL & "(SELECT COUNT(*) AS QuantitaTipoPedana "
'
'sSQL = sSQL & "FROM (SELECT " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDPedana "
'sSQL = sSQL & "FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " INNER JOIN "
'sSQL = sSQL & "RV_POTipoPedana ON " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
'sSQL = sSQL & "WHERE (" & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".IDOggetto = " & ObjDoc.IDOggetto & ") And (" & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POTipoRiga = 1) "
'sSQL = sSQL & "GROUP BY " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDPedana, RV_POTipoPedana.IDArticoloImballo) AS X) AS QuantitaPedana"
'
'sSQL = sSQL & " FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " INNER JOIN "
'sSQL = sSQL & " RV_POTipoPedana ON " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
'sSQL = sSQL & " WHERE IDOggetto = " & ObjDoc.IDOggetto
'sSQL = sSQL & " AND  RV_POTipoRiga = 1 "
'sSQL = sSQL & " GROUP BY IDOggetto, RV_POIDTipoPedana, IDArticoloImballo"

sSQL = "SELECT IDOggetto, RV_POIDTipoPedana, RV_POTipoPedana.IDArticoloImballo,"
sSQL = sSQL & " (SELECT COUNT(*) AS QuantitaTipoPedana "
sSQL = sSQL & " FROM (SELECT RV_POIDPedana "
sSQL = sSQL & " FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " AS Tabella INNER JOIN "
sSQL = sSQL & " RV_POTipoPedana ON Tabella.RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
sSQL = sSQL & " Where (IDOggetto = " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".IDOggetto) And (RV_POTipoRiga = 1) And (RV_POIDTipoPedana = " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana) "
sSQL = sSQL & " GROUP BY Tabella.RV_POIDPedana, RV_POTipoPedana.IDArticoloImballo) AS X) AS QuantitaPedana "
sSQL = sSQL & " FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " INNER JOIN "
sSQL = sSQL & " RV_POTipoPedana ON " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
sSQL = sSQL & " Where IDOggetto = " & ObjDoc.IDOggetto
sSQL = sSQL & " AND  RV_POTipoRiga = 1"
sSQL = sSQL & " GROUP BY IDOggetto, RV_POIDTipoPedana, IDArticoloImballo"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    If (DISATTIVA_SCALATA_COMM_TRASP = 1) Then
        sSQL = "SELECT * FROM RV_POIEConfigurazioneClienteTrasporto "
        sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND IDAnagrafica=" & IDCliente
        sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDSitoPerAnagrafica
        sSQL = sSQL & " AND IDArticolo=" & fnNotNullN(rs!IDArticoloImballo)
        sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
        sSQL = sSQL & " AND CommissionePerPedana=1"
        
        Set rsCli = New ADODB.Recordset
        
        rsCli.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
        
        While Not rsCli.EOF
            If GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED(fnNotNullN(rsCli!IDRV_POTipoCommissione), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNullN(rs!IDArticoloImballo)) = False Then
                rsNew.AddNew
                    rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
                    rsNew!IDOggetto = IDOggettoDocumento
                    rsNew!IDRV_POTipoCommissione = fnNotNullN(rsCli!IDRV_POTipoCommissione)
                    rsNew!Percentuale = 0
                    rsNew!Importo = 0
                    rsNew!ImportoRiga = fnNotNullN(rsCli!PrezzoTrasporto)
                    rsNew!Quantita = fnNotNullN(rs!QuantitaPedana)
                    rsNew!ImportoTotale = 0
                    rsNew!IDRV_POTipoPedana = fnNotNullN(rs!RV_POIDTipoPedana)
                    rsNew!APercentuale = 1
                    rsNew!IDArticoloImballo = fnNotNullN(rs!IDArticoloImballo)
                rsNew.Update
            End If
        rsCli.MoveNext
        Wend
        
        rsCli.Close
        Set rsCli = Nothing

        sSQL = "SELECT * FROM RV_POTipoPedana "
        sSQL = sSQL & " WHERE IDRV_POTipoPedana=" & fnNotNullN(rs!RV_POIDTipoPedana)
        sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
        
        Set rsPed = New ADODB.Recordset
        rsPed.Open sSQL, CnDMT.InternalConnection
        
        While Not rsPed.EOF
            If GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED(fnNotNullN(rsPed!IDRV_POTipoCommissione), fnNotNullN(rsPed!IDRV_POTipoPedana), fnNotNullN(rsPed!IDArticoloImballo)) = False Then
                rsNew.AddNew
                    rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
                    rsNew!IDOggetto = IDOggettoDocumento
                    rsNew!IDRV_POTipoCommissione = fnNotNullN(rsPed!IDRV_POTipoCommissione)
                    rsNew!Percentuale = 0
                    rsNew!Importo = 0
                    rsNew!ImportoRiga = fnNotNullN(rsPed!ImportoCommissione)
                    rsNew!Quantita = fnNotNullN(rs!QuantitaPedana)
                    rsNew!ImportoTotale = 0
                    rsNew!IDRV_POTipoPedana = fnNotNullN(rsPed!IDRV_POTipoPedana)
                    rsNew!APercentuale = 1
                    rsNew!IDArticoloImballo = fnNotNullN(rsPed!IDArticoloImballo)
                rsNew.Update
            End If
        rsPed.MoveNext
        Wend
        
        rsPed.Close
        Set rsPed = Nothing
    Else
        sSQL = "SELECT * FROM RV_POIEConfigurazioneClienteTrasporto "
        sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
        sSQL = sSQL & " AND IDAnagrafica=" & IDCliente
        sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDSitoPerAnagrafica
        sSQL = sSQL & " AND IDArticolo=" & fnNotNullN(rs!IDArticoloImballo)
        sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
        sSQL = sSQL & " AND CommissionePerPedana=1"
        
        Set rsCli = New ADODB.Recordset
        
        rsCli.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
        If Not rsCli.EOF Then
            While Not rsCli.EOF
                If GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED(fnNotNullN(rsCli!IDRV_POTipoCommissione), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNullN(rs!IDArticoloImballo)) = False Then
                    rsNew.AddNew
                        rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
                        rsNew!IDOggetto = IDOggettoDocumento
                        rsNew!IDRV_POTipoCommissione = fnNotNullN(rsCli!IDRV_POTipoCommissione)
                        rsNew!Percentuale = 0
                        rsNew!Importo = 0
                        rsNew!ImportoRiga = fnNotNullN(rsCli!PrezzoTrasporto)
                        rsNew!Quantita = fnNotNullN(rs!QuantitaPedana)
                        rsNew!ImportoTotale = 0
                        rsNew!IDRV_POTipoPedana = fnNotNullN(rs!RV_POIDTipoPedana)
                        rsNew!APercentuale = 1
                        rsNew!IDArticoloImballo = fnNotNullN(rs!IDArticoloImballo)
                    rsNew.Update
                End If
            rsCli.MoveNext
            Wend
            
            rsCli.Close
            Set rsCli = Nothing
            
        Else
            rsCli.Close
            Set rsCli = Nothing
        
            sSQL = "SELECT * FROM RV_POTipoPedana "
            sSQL = sSQL & " WHERE IDRV_POTipoPedana=" & fnNotNullN(rs!RV_POIDTipoPedana)
            sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
            
            Set rsPed = New ADODB.Recordset
            rsPed.Open sSQL, CnDMT.InternalConnection
            
            While Not rsPed.EOF
                If GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED(fnNotNullN(rsPed!IDRV_POTipoCommissione), fnNotNullN(rsPed!IDRV_POTipoPedana), fnNotNullN(rsPed!IDArticoloImballo)) = False Then
                    rsNew.AddNew
                        rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
                        rsNew!IDOggetto = IDOggettoDocumento
                        rsNew!IDRV_POTipoCommissione = fnNotNullN(rsPed!IDRV_POTipoCommissione)
                        rsNew!Percentuale = 0
                        rsNew!Importo = 0
                        rsNew!ImportoRiga = fnNotNullN(rsPed!ImportoCommissione)
                        rsNew!Quantita = fnNotNullN(rs!QuantitaPedana)
                        rsNew!ImportoTotale = 0
                        rsNew!IDRV_POTipoPedana = fnNotNullN(rsPed!IDRV_POTipoPedana)
                        rsNew!APercentuale = 1
                        rsNew!IDArticoloImballo = fnNotNullN(rsPed!IDArticoloImballo)
                    rsNew.Update
                End If
            rsPed.MoveNext
            Wend
            
            rsPed.Close
            Set rsPed = Nothing
            
        End If
    End If
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_IMPOSTA_COMMISSIONI_PER_CLIENTE:
    MsgBox Err.Description, vbCritical, "IMPOSTA_COMMISSIONI_PER_TIPO_PEDANA"

End Sub

Private Function GET_TOTALE_MERCE_DOCUMENTO(IDOggettoDocumento As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoListinoImballo As Double


GET_TOTALE_MERCE_DOCUMENTO = 0

sSQL = "SELECT Art_importo_totale_neutro, RV_POIDImballo, RV_POImportoUnitarioImballo, "
sSQL = sSQL & "RV_POImportoImballoInArticolo, Art_numero_colli, Art_quantita_totale, RV_POImportoMerceNetta "
sSQL = sSQL & "FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
sSQL = sSQL & " WHERE (IDOggetto = " & IDOggettoDocumento & ") "
sSQL = sSQL & " AND (RV_POTipoRiga = 1)"

Set rs = CnDMT.OpenResultset(sSQL)
While Not rs.EOF
    GET_TOTALE_MERCE_DOCUMENTO = GET_TOTALE_MERCE_DOCUMENTO + (fnNotNullN(rs!RV_POImportoMerceNetta) * fnNotNullN(rs!Art_quantita_totale))
    
'    If fnNotNullN(rs!RV_POImportoImballoInArticolo) = False Then
'        GET_TOTALE_MERCE_DOCUMENTO = GET_TOTALE_MERCE_DOCUMENTO + fnNotNullN(rs!Art_importo_totale_neutro)
'    Else
'        ImportoListinoImballo = GET_PREZZO_IMBALLO_PER_COMMMISSIONI(fnNotNullN(rs!RV_POIDImballo), ObjDoc.Field("Link_Doc_Listino", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
'        GET_TOTALE_MERCE_DOCUMENTO = GET_TOTALE_MERCE_DOCUMENTO + fnNotNullN(rs!Art_importo_totale_neutro) - (ImportoListinoImballo * fnNotNullN(rs!Art_numero_colli))
'    End If

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_PREZZO_IMBALLO_PER_COMMMISSIONI(IDArticolo As Long, IDListino As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE ("
sSQL = sSQL & "(IDListino=" & IDListino & ") "
sSQL = sSQL & "AND (IDArticolo=" & IDArticolo & "))"
    
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_PREZZO_IMBALLO_PER_COMMMISSIONI = fnNotNullN(rs!PrezzoNettoIva)
Else
    GET_PREZZO_IMBALLO_PER_COMMMISSIONI = 0
End If
 
rs.CloseResultset
Set rs = Nothing

End Function

Private Sub SCRIVI_RIFERIMENTO_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDVettoreTour As Long
Dim TargaAutomezzoTour As String
Dim IDAgenziaTrasporto As Long


sSQL = "SELECT * "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    
    GET_DATI_TOUR IDOggettoOrdine, IDVettoreTour, TargaAutomezzoTour, IDAgenziaTrasporto
    
    If IDAgenziaTrasporto > 0 Then
        ObjDoc.Field "RV_POIDAgenziaTrasportatore", IDAgenziaTrasporto, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    If IDVettoreTour > 0 Then
        ObjDoc.ReadDataFromCarrier IDVettoreTour, MainCarrier, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    Else
        ObjDoc.ReadDataFromCarrier fnNotNullN(rs!Link_Vet_vettore), MainCarrier, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    If Len(Trim(TargaAutomezzoTour)) > 0 Then
        ObjDoc.Field "RV_POTargaAutomezzo", TargaAutomezzoTour, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    Else
        ObjDoc.Field "RV_POTargaAutomezzo", fnNotNull(rs!RV_POTargaAutomezzo), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    If fnNotNullN(rs!Link_Doc_pagamento) > 0 Then
        ObjDoc.ReadDataFromPayment fnNotNullN(rs!Link_Doc_pagamento), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    If fnNotNullN(rs!Link_Doc_agente) > 0 Then
        ObjDoc.ReadDataFromAgent fnNotNullN(rs!Link_Doc_agente), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    If Len(Trim(fnNotNull(rs!Doc_causale_documento))) > 0 Then
        ObjDoc.Field "Doc_causale_documento", fnNotNull(rs!Doc_causale_documento), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    If fnNotNullN(rs!Link_Nom_raggrup_fatturato) > 0 Then
        ObjDoc.Field "Link_Nom_raggrup_fatturato", fnNotNull(rs!Link_Nom_raggrup_fatturato), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    
    ObjDoc.Field "Link_nom_lettera_intento", fnNotNullN(rs!Link_Nom_Lettera_Intento), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Nom_IVA", fnNotNullN(rs!Link_Nom_IVA), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Nom_contratto_bancario", fnNotNullN(rs!Link_Nom_contratto_bancario), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Doc_contratto_bancario_az", fnNotNullN(rs!Link_Doc_contratto_bancario_az), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Nom_accordi_commerciali", fnNotNullN(rs!Link_Nom_accordi_commerciali), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Doc_Listino", fnNotNullN(rs!Link_Doc_Listino), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Doc_listino_base", fnNotNullN(rs!Link_Doc_listino_base), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    ObjDoc.Field "Link_Val_valuta", fnNotNullN(rs!Link_Val_valuta), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'    ObjDoc.Field "Doc_data_ns_ordine_di_rifer", fnNotNull(rs!Doc_data), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'    ObjDoc.Field "Doc_numero_ns_ordine_di_rifer", fnNotNull(rs!Doc_numero), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    ObjDoc.Field "Doc_data_ns_ordine_di_rifer", fnNotNull(rs!RV_PODataOrdinePadre), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Doc_numero_ns_ordine_di_rifer", fnNotNull(rs!RV_PONumeroOrdinePadre), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POIDOggettoOrdineOri", fnNotNull(rs!IDOggetto), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    ObjDoc.Field "Doc_data_vs_ordine_di_rifer", fnNotNull(rs!Doc_data_presso_nom), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Doc_numero_vs_ordine_di_rifer", fnNotNull(rs!Doc_numero_presso_nom), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    ObjDoc.Field "Doc_annotazioni_variazio", fnNotNull(rs!Doc_annotazioni_variazio), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POAnnotazioniInterna", fnNotNull(rs!RV_POAnnotazioniInterna), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POIDLuogoPresaMerce", fnNotNullN(rs!RV_POIDLuogoPresaMerce), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POIDTrasportatoreSuccessivo", fnNotNullN(rs!RV_POIDTrasportatoreSuccessivo), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    ObjDoc.Field "RV_PODataArrivoMerce", fnNotNullN(rs!RV_PODataArrivoMerce), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    If ObjDoc.IDTipoOggetto = 2 Then
        ObjDoc.Field "Doc_data_consegna_bolla", fnNotNullN(rs!RV_PODataArrivoMerce), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    
    ObjDoc.Field "RV_POOraArrivoMerce", fnNotNull(rs!RV_POOraArrivoMerce), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_PODataArrivoMerceLuogo", fnNotNullN(rs!RV_PODataArrivoMerceLuogo), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POOraArrivoMerceLuogo", fnNotNull(rs!RV_POOraArrivoMerceLuogo), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Nom_porto", fnNotNullN(rs!Link_Nom_porto), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
    If Len(Trim(Me.txtIstruzioniMittente.Text)) > 0 Then
        ObjDoc.Field "RV_POIstruzioniMittente", Me.txtIstruzioniMittente.Text, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    Else
        ObjDoc.Field "RV_POIstruzioniMittente", fnNotNull(rs!RV_POIstruzioniMittente), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    End If
    ObjDoc.Field "Link_Doc_spedizione", fnNotNullN(rs!Link_Doc_spedizione), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "Link_Doc_aspetto_esteriore", fnNotNullN(rs!Link_Doc_aspetto_esteriore), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    ObjDoc.Field "RV_POIDAnagraficaDestinazione", fnNotNullN(rs!RV_POIDAnagraficaDestinazione), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
    
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_DESCRIZIONE_RIGA_CORPO(IDOggettoOrdine As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT DescrizioneCorpoDocEv "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_DESCRIZIONE_RIGA_CORPO = fnNotNull(rs!DescrizioneCorpoDocEv)
Else
    GET_DESCRIZIONE_RIGA_CORPO = ""
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Function GET_TIPO_IMPORTO_ARTICOLO_DA_LIQUIDAZIONE() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDTipoImportoArticolo "
sSQL = sSQL & "FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_IMPORTO_ARTICOLO_DA_LIQUIDAZIONE = 0
Else
    If fnNotNullN(rs!IDTipoImportoArticolo) = 3 Then
        GET_TIPO_IMPORTO_ARTICOLO_DA_LIQUIDAZIONE = 1
    Else
        GET_TIPO_IMPORTO_ARTICOLO_DA_LIQUIDAZIONE = 0
    End If
End If



rs.CloseResultset
Set rs = Nothing
End Function

Private Sub LVFattureCreate_DblClick()
On Error GoTo ERR_LVFattureCreate_DblClick

    Me.LabelLink1.DisableDoEvents = True
    
    Select Case Me.LVFattureCreate.ListItems(Me.LVFattureCreate.SelectedItem.Index).SubItems(1)
        Case 2
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_PODDTL"))
            Me.LabelLink1.IDReturn = CLng(Me.LVFattureCreate.SelectedItem.Text)
            Me.LabelLink1.RunApplication
        Case 114
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POFAL"))
            Me.LabelLink1.IDReturn = CLng(Me.LVFattureCreate.SelectedItem.Text)
            Me.LabelLink1.RunApplication
        Case 8
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POSNFL"))
            Me.LabelLink1.IDReturn = CLng(Me.LVFattureCreate.SelectedItem.Text)
            Me.LabelLink1.RunApplication
    End Select
    
Exit Sub
ERR_LVFattureCreate_DblClick:
    MsgBox Err.Description, vbCritical, "LVFattureCreate_DblClick"
End Sub

Private Sub txtDataDoc_LostFocus()
    cboSezionale_Click
End Sub
Private Function GET_CONTROLLO_CLIENTE_BLOCCATO(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

''''''''''FIDO PER CLIENTE
sSQL = "SELECT BloccoEmissioneDoc "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_CLIENTE_BLOCCATO = 0
    
Else
    GET_CONTROLLO_CLIENTE_BLOCCATO = fnNotNullN(rs!BloccoEmissioneDoc)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_DATI_ARTICOLO_INTRA(IDArticolo As Long, PesoNetto As Double, Pezzi As Double)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MassaNettaInKg, IDNomenclaturaCombinata FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    ObjDoc.Field "Art_intra_qta_tot_massa_netta", PesoNetto, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
Else
    If fnNotNullN(rs!MassaNettaInKg) <= 1 Then
        ObjDoc.Field "Art_intra_qta_tot_massa_netta", PesoNetto, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    Else
        ObjDoc.Field "Art_intra_qta_tot_massa_netta", Pezzi * fnNotNullN(rs!MassaNettaInKg), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    End If
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function LINK_SEZIONALE_CMR() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionaleCMRPred FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_SEZIONALE_CMR = fnNotNullN(rs!IDSezionaleCMRPred)
Else
    LINK_SEZIONALE_CMR = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
'Funzionalita': Determina il numero documento.
Function fnDocumentNumberCMR(sDate As String, IDSezionale As Long, IDPeriodoIva As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT ProgressivoDisponibile "
sSQL = sSQL & "FROM ProgressivoSezionale "
sSQL = sSQL & "WHERE IDPeriodoIva=" & IDPeriodoIva
sSQL = sSQL & " AND IDTipoModulo=1"
sSQL = sSQL & " AND IDSezionale=" & IDSezionale
sSQL = sSQL & " AND VirtualDelete=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    fnDocumentNumberCMR = 1
Else
    fnDocumentNumberCMR = fnNotNullN(rs!ProgressivoDisponibile)
End If

rs.CloseResultset
Set rs = Nothing


End Function

Private Sub txtDataDocCMR_LostFocus()
    cboSezionaleCMR_Click
End Sub
Private Sub GET_DATI_CMR(IDSezionale As Long, IDPeriodoIva As Long, NumeroCMR As Long, DataCMR As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroNuovoCMR As Long


If GET_CONTROLLO_NUMERO_DOCUMENTO(IDSezionale, DataCMR, NumeroCMR, 0) = True Then
    NumeroNuovoCMR = fnDocumentNumberCMR(DataCMR, IDSezionale, IDPeriodoIva)
Else
    NumeroNuovoCMR = NumeroCMR
End If

ObjDoc.Field "RV_PODocumentoCRM", 1, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
ObjDoc.Field "RV_POIDSezionaleCMR", IDSezionale, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
ObjDoc.Field "RV_PODataDocumentoCMR", DataCMR, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
ObjDoc.Field "RV_PONumeroDocumentoCMR", NumeroNuovoCMR, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
 
fnAvanzamentoProgressivoCMR IDSezionale, NumeroNuovoCMR, DataCMR

End Sub
Private Function GET_CONTROLLO_NUMERO_DOCUMENTO(IDSezionaleCMR As Long, DataDocumentoCMR As String, NumeroDocumentoCMR As Long, IDOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim DataInizioPeriodo As String
Dim DataFinePeriodo As String

GET_CONTROLLO_NUMERO_DOCUMENTO = False

DataInizioPeriodo = "01/01/" & Year(DataDocumentoCMR)
DataFinePeriodo = "31/12/" & Year(DataDocumentoCMR)


''''CONTROLLO SUI DOCUMENTI DI TRASPORTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT ValoriOggettoPerTipo0002.IDOggetto, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Oggetto.Oggetto "
sSQL = sSQL & "FROM Oggetto INNER JOIN "
sSQL = sSQL & "ValoriOggettoPerTipo0002 ON Oggetto.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto AND "
sSQL = sSQL & "Oggetto.IDTipoOggetto = ValoriOggettoPerTipo0002.IDTipoOggetto "
sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND ValoriOggettoPerTipo0002.RV_PODataDocumentoCMR>=" & fnNormDate(DataInizioPeriodo)
sSQL = sSQL & " AND ValoriOggettoPerTipo0002.RV_PODataDocumentoCMR<=" & fnNormDate(DataFinePeriodo)
sSQL = sSQL & " AND ValoriOggettoPerTipo0002.RV_PONumeroDocumentoCMR=" & fnNormNumber(NumeroDocumentoCMR)
sSQL = sSQL & " AND ValoriOggettoPerTipo0002.RV_POIDSezionaleCMR=" & IDSezionaleCMR
If IDOggetto > 0 Then
    sSQL = sSQL & " AND Oggetto.IDOggetto<>" & IDOggetto
End If

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_NUMERO_DOCUMENTO = False
Else
    GET_CONTROLLO_NUMERO_DOCUMENTO = True
End If


rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If GET_CONTROLLO_NUMERO_DOCUMENTO = False Then
    
    ''''CONTROLLO SULLE FATTURE ACCOMPAGNATORIE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT ValoriOggettoPerTipo0072.IDOggetto, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, Oggetto.Oggetto "
    sSQL = sSQL & "FROM Oggetto INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0072 ON Oggetto.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto AND "
    sSQL = sSQL & "Oggetto.IDTipoOggetto = ValoriOggettoPerTipo0072.IDTipoOggetto "
    sSQL = sSQL & "WHERE Oggetto.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND ValoriOggettoPerTipo0072.RV_PODataDocumentoCMR>=" & fnNormDate(DataInizioPeriodo)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0072.RV_PODataDocumentoCMR<=" & fnNormDate(DataFinePeriodo)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0072.RV_PONumeroDocumentoCMR=" & fnNormNumber(NumeroDocumentoCMR)
    sSQL = sSQL & " AND ValoriOggettoPerTipo0072.RV_POIDSezionaleCMR=" & IDSezionaleCMR
    If IDOggetto > 0 Then
        sSQL = sSQL & " AND Oggetto.IDOggetto<>" & IDOggetto
    End If
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLO_NUMERO_DOCUMENTO = False
    Else
        GET_CONTROLLO_NUMERO_DOCUMENTO = True
    End If

    rs.CloseResultset
    Set rs = Nothing

End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Function
Private Function fnAvanzamentoProgressivoCMR(IDSezionaleCMR As Long, NumeroDocumentoCMR As Long, DataDocumentoCMR As String) As Boolean
Dim link_PeriodoIva As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

If IDSezionaleCMR = 0 Then Exit Function

link_PeriodoIva = fnGetPeriodoIVA(DataDocumentoCMR)

sSQL = "SELECT * "
sSQL = sSQL & "FROM ProgressivoSezionale "
sSQL = sSQL & "WHERE IDPeriodoIva=" & link_PeriodoIva
sSQL = sSQL & " AND IDSezionale=" & IDSezionaleCMR
sSQL = sSQL & " AND IDTipoModulo=1"

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    rs!IDProgressivoSezionale = fnGetNewKey("ProgressivoSezionale", "IDProgressivoSezionale")
    rs!IDPeriodoIva = link_PeriodoIva
    rs!IDTipoModulo = 1
    rs!IDSezionale = IDSezionaleCMR
    rs!ProgressivoDisponibile = NumeroDocumentoCMR + 1
    rs!DataUltimaVariazione = Date
    rs!IDUtenteUltimaVariazione = TheApp.IDUser
    rs!VirtualDelete = 0
Else
    If NumeroDocumentoCMR >= fnNotNullN(rs!ProgressivoDisponibile) Then
        rs!ProgressivoDisponibile = NumeroDocumentoCMR + 1
    End If
End If

rs.Update

rs.Close
Set rs = Nothing
End Function

Private Function GET_DATI_TOUR(IDOggettoOrdine As Long, IDVettoreTour As Long, TargaAutomezzoTour As String, IDAgenziaTrasporto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POTour.IDVettore, RV_POTourRighe.IDOggettoOrdine, RV_POTour.TargaMezzoTrasporto, IDAnagraficaFornitore "
sSQL = sSQL & "FROM RV_POTour INNER JOIN "
sSQL = sSQL & "RV_POTourRighe ON RV_POTour.IDRV_POTour = RV_POTourRighe.IDRV_POTour "
sSQL = sSQL & "WHERE RV_POTourRighe.IDOggettoOrdine=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    IDVettoreTour = 0
    TargaAutomezzoTour = ""
    IDAgenziaTrasporto = 0
Else
    
    IDVettoreTour = fnNotNullN(rs!IDVettore)
    TargaAutomezzoTour = fnNotNull(rs!TargaMezzoTrasporto)
    IDAgenziaTrasporto = fnNotNullN(rs!IDAnagraficaFornitore)
    
End If

rs.CloseResultset
Set rs = Nothing


If Me.cboVettore.CurrentID > 0 Then
    IDVettoreTour = Me.cboVettore.CurrentID
End If

If Len(Trim(Me.txtTargaAutomezzo.Text)) > 0 Then
    TargaAutomezzoTour = Me.txtTargaAutomezzo.Text
End If

End Function
Private Function GET_CONTROLLO_STATO_ORDINE(IDOggettoOrdine As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_STATO_ORDINE = ""


'''''''''''''''''''CONTROLLLO DEL BLOCCO ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT Link_Vet_Vettore, Link_nom_ult_sito, Doc_ordine_chiuso, Doc_data_prevista_evasione, "
sSQL = sSQL & "Doc_data_presso_nom, Doc_numero_presso_nom, Doc_annotazioni_variazio, RV_POAnnotazioniInterna, "
sSQL = sSQL & "RV_PODescrizioneCorpoDocEv, RV_POIDLuogoPresaMerce, RV_POIDTrasportatoreSuccessivo, RV_POIDUtenteBlocco "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_STATO_ORDINE = ""
Else
    If fnNotNullN(rs!RV_POIDUtenteBlocco) > 0 Then
        GET_CONTROLLO_STATO_ORDINE = "Ordine bloccato dall'utente " & GET_UTENTE(fnNotNullN(rs!RV_POIDUtenteBlocco))
    End If
End If

rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_CONTROLLO_STATO_ORDINE)) > 0 Then Exit Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE STANDARD''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND IDTipoOggetto=" & 15
sSQL = sSQL & "  AND IDFunzione=" & 128

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_CONTROLLO_STATO_ORDINE = ""
Else
    
    GET_CONTROLLO_STATO_ORDINE = "Ordine aperto dall'utente " & GET_UTENTE(fnNotNullN(rs!IDUtente))
    
End If

rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_CONTROLLO_STATO_ORDINE)) > 0 Then Exit Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''CONTROLLO DELL'ORDINE BLOCCATO DA GESTORE ORDINE GREEN TOP'''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_POOrdineL")
sSQL = sSQL & "  AND IDFunzione=" & GET_FUNZIONE_PER_TIPO_OGGETTO(fnGetTipoOggetto("RV_POOrdineL"))

Set rs = CnDMT.OpenResultset(sSQL)
If rs.EOF Then
    GET_CONTROLLO_STATO_ORDINE = ""
Else
    
    GET_CONTROLLO_STATO_ORDINE = "Ordine aperto dall'utente " & GET_UTENTE(fnNotNullN(rs!IDUtente))
    
End If

rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_CONTROLLO_STATO_ORDINE)) > 0 Then Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Function
Private Function GET_UTENTE(IDUtente As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Utente FROM Utente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_UTENTE = ""
Else
    GET_UTENTE = fnNotNull(rs!Utente)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_FUNZIONE_PER_TIPO_OGGETTO(IDTipoOggetto As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDFunzione "
sSQL = sSQL & "FROM Funzione  "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE_PER_TIPO_OGGETTO = 0
Else
    GET_FUNZIONE_PER_TIPO_OGGETTO = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = CnDMT.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If

    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function GET_RIFERIMENTO_ORDINE(IDOggettoOrdine As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDOggetto, Doc_numero, Doc_data, Nom_ragione_sociale_o_cognome "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_RIFERIMENTO_ORDINE = ""
Else
    GET_RIFERIMENTO_ORDINE = "Rif. ordine n° " & fnNotNullN(rs!Doc_numero) & " del " & fnNotNull(rs!Doc_data) & vbCrLf
    GET_RIFERIMENTO_ORDINE = GET_RIFERIMENTO_ORDINE & " intestato al cliente " & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & vbCrLf & vbCrLf
End If

rs.CloseResultset
Set rs = Nothing
End Function
Public Function TrovaCartella(IDLCartella As Long) As String

    TrovaCartella = String$(MAX_PATH, 0)
    
    Call SHGetSpecialFolderPath(ByVal 0&, TrovaCartella, IDLCartella, ByVal 0&)
    
    TrovaCartella = Left$(TrovaCartella, InStr(1, TrovaCartella, Chr$(0)) - 1)
    
    If Len(TrovaCartella) > 0 And Right$(TrovaCartella, 1) <> "\" Then TrovaCartella = TrovaCartella & "\"
    
End Function
Private Function fnSetCausaleDocumento() As String
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    sSQL = "SELECT CausaleTrasporto FROM CausaleTrasportoPerFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & ObjDoc.IDFunzione
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        fnSetCausaleDocumento = ""
    Else
        fnSetCausaleDocumento = fnNotNull(rs!CausaleTrasporto)
    End If
    
    
    rs.CloseResultset
    Set rs = Nothing
    
End Function
Private Function GET_LISTINO_PER_DESTINAZIONE(IDAnagrafica As Long, IDSitoPerAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDListino FROM RV_POConfigurazioneClienteListino "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDSitoPerAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_PER_DESTINAZIONE = 0
Else
    GET_LISTINO_PER_DESTINAZIONE = fnNotNullN(rs!IDListino)
End If



rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_MODALITA_PAGAMENTO(DataDocumento As String, IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim GiornoMese As Long

GiornoMese = Day(DataDocumento)

sSQL = "SELECT IDPagamento FROM RV_POConfigurazioneClientePagamenti "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND DalGiornoMese<=" & GiornoMese
sSQL = sSQL & " AND AlGiornoMese>=" & GiornoMese
Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_MODALITA_PAGAMENTO = 0
Else
    GET_MODALITA_PAGAMENTO = fnNotNullN(rs!IDPagamento)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_RIGHE_DOCUMENTO(NomeTabella As String)
On Error GoTo ERR_AGGIORNA_RIGHE_DOCUMENTO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsCount As DmtOleDbLib.adoResultset
Dim Unita_Progresso As Double
Dim NumeroRecord As Long
Dim Moltiplicatore As Double
Dim Prezzo As Double
Dim MerceNettaPerLiquidazione As Double
Dim IDUMCoop_ArtVenduto As Long
Dim PrezzoScontato As Double

sSQL = "SELECT IDOggetto, RV_POPrezzoMedioInLiq, RV_POVariazionePrezzoImballo, RV_POImportoMerceNetta, RV_POImportoLiq, "
sSQL = sSQL & "RV_POPrezzoUnitarioOrigine, RV_POIDIvaImballo, RV_POIDTipoVariazione, Art_pre_uni_net_sco_net_IVA ,"
sSQL = sSQL & "RV_POImportoDaLiq, RV_POLinkRiga, RV_POImportoImballoInArticolo, "
sSQL = sSQL & "Art_numero_colli, Art_quantita_Totale, Link_art_articolo, "
sSQL = sSQL & "RV_POVariazionePrezzoManuale, RV_POImportoRigaCommissioni,  "
sSQL = sSQL & "RV_POIDConferimentoRighe, RV_POIDAssegnazioneMerce, RV_POIDProcessoIVGamma, RV_PODataLavorazione, "
sSQL = sSQL & "RV_POAnnotazioniAggiuntiveLav, RV_PONotaRigaOrdRaggr, RV_PODataOrdineCliente, RV_PONumeroOrdineCliente, "
sSQL = sSQL & "RV_PODataOrdineInterno, RV_PONumeroOrdineInterno,  "
sSQL = sSQL & "RV_POIDImballoPrim, RV_PONumeroConfezioniPerImballo, RV_POCostoConfezioneImballo, RV_POCostoKitLiq, "
sSQL = sSQL & "RV_POCostoConfezioneImballoLiq, RV_POQuantitaLiq,  "
sSQL = sSQL & "Art_numero_colli,Art_Peso, Art_Tara, Art_quantita_pezzi, RV_POImportoImballoSel, RV_POImportoLiqDoc, Art_prezzo_unitario_neutro, "
sSQL = sSQL & "Art_sco_in_percentuale_1,Art_sco_in_percentuale_2 "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE RV_POTipoRiga=1 "
sSQL = sSQL & " AND IDOggetto=" & ObjDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto

Set rs = New ADODB.Recordset

rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    
    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_Art_articolo))
    rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
    IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_Art_articolo))
    
    Select Case IDUMCoop_ArtVenduto
        Case 1
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_numero_colli) * Moltiplicatore
        Case 2
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_peso) * Moltiplicatore
        Case 3
            rs!RV_POQuantitaLiq = (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)) * Moltiplicatore
        Case 4
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_tara) * Moltiplicatore
        Case 5
            rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_pezzi) * Moltiplicatore
    End Select

    rs!RV_POImportoDaLiq = 0
    If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 1 Then
        rs!RV_POImportoDaLiq = -(fnNotNullN(rs!RV_POImportoImballoSel) * fnNotNullN(rs!Art_numero_colli)) / fnNotNullN(rs!RV_POQuantitaLiq)
    End If
    
    
    If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 0 Then
        rs!RV_POVariazionePrezzoImballo = 0
        rs!RV_POImportoMerceNetta = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
    Else
        rs!RV_POVariazionePrezzoImballo = ((Abs(fnNotNullN(rs!RV_POImportoDaLiq)) * fnNotNullN(rs!RV_POQuantitaLiq)) / fnNotNullN(rs!Art_quantita_totale))
        rs!RV_POImportoMerceNetta = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA) - rs!RV_POVariazionePrezzoImballo
    End If

    'Prezzo = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
    
    Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
    Prezzo = (Prezzo * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
    'Prezzo = Prezzo / Moltiplicatore
    'MerceNettaPerLiquidazione = fnNotNullN(rs!RV_POImportoMerceNetta) / Moltiplicatore
    MerceNettaPerLiquidazione = Prezzo
    PrezzoScontato = Prezzo
    
    If fnNotNullN(rs!RV_POImportoImballoInArticolo) > 0 Then
        Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoDaLiq)
    Else
        Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoDaLiq))
    End If
    
    'CALCOLO DEL PREZZO DI LIQUIDAZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, ObjDoc.IDOggetto, ObjDoc.IDTipoOggetto, Moltiplicatore, rs, MerceNettaPerLiquidazione, fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), PrezzoScontato)
    
    If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
        Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
    Else
        Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
    End If
    
    rs!RV_POCostoConfezioneImballoLiq = 0
    rs!RV_POCostoKitLiq = 0
    
    'COSTO CONFEZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If fnNotNullN(rs!RV_PONumeroConfezioniPerImballo) > 0 Then
        If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
            rs!RV_POCostoConfezioneImballoLiq = (((fnNotNullN(rs!Art_numero_colli) * fnNotNullN(rs!RV_PONumeroConfezioniPerImballo)) * fnNotNullN(rs!RV_POCostoConfezioneImballo)) / fnNotNullN(rs!RV_POQuantitaLiq))
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'COSTO KIT''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If fnNotNullN(rs!RV_POIDAssegnazioneMerce) > 0 Then
        If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
            rs!RV_POCostoKitLiq = GET_TOTALE_COSTO_KIT(fnNotNullN(rs!RV_POIDAssegnazioneMerce)) / fnNotNullN(rs!RV_POQuantitaLiq)
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If AGG_COSTO_CONFEZ_PRZ_LIQ > 0 Then
        Prezzo = Prezzo - rs!RV_POCostoConfezioneImballoLiq
    End If
    
    If AGG_COSTO_KIT_PRZ_LIQ > 0 Then
        Prezzo = Prezzo - rs!RV_POCostoKitLiq
    End If
    
    rs!RV_POImportoLiq = Prezzo
    rs!RV_POImportoLiqDoc = Prezzo
    rs.Update
    
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
Exit Sub
ERR_AGGIORNA_RIGHE_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "AGGIORNA_RIGHE_DOCUMENTO"
End Sub
Private Sub GET_PARAMETRI_IVA_IMBALLO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IvaBloccata, IvaImballoARendere "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDUtente=" & 0

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    FLAG_IVA_UGUALE = 0
    FLAG_IVA_IMBALLO_A_RENDERE = 0
Else
    FLAG_IVA_UGUALE = Abs(fnNotNullN(rs!IvaBloccata))
    FLAG_IVA_IMBALLO_A_RENDERE = Abs(fnNotNullN(rs!IvaImballoARendere))
    
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_TIPO_IMBALLO_A_RENDERE(IDArticoloImballo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo.IDIvaVendita, Iva.AliquotaIva, Articolo.Tara, Articolo.Articolo, Articolo.NonRiportoIntrastat, Articolo.MassaNettaInKg, "
sSQL = sSQL & "Articolo.IDNomenclaturaCombinata , Articolo.RV_POIDNaturaTransazione, RV_POTipoImballo.Rendere "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoImballo ON Articolo.RV_POIDTipoImballo = RV_POTipoImballo.IDRV_POTipoImballo LEFT OUTER JOIN "
sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticoloImballo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_IMBALLO_A_RENDERE = 0
Else
    GET_TIPO_IMBALLO_A_RENDERE = Abs(fnNotNullN(rs!Rendere))
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CESSIONE_INTRA_CLIENTE(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDClassificazioneIva FROM Cliente "
sSQL = sSQL & "WHERE Cliente.IDAzienda =" & TheApp.IDFirm
sSQL = sSQL & " AND Cliente.IDAnagrafica=" & IDAnagraficaCliente

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_CESSIONE_INTRA_CLIENTE = 0
Else
    If fnNotNullN(rs!IDClassificazioneIva) = 1 Then
        GET_CESSIONE_INTRA_CLIENTE = 1
    Else
        GET_CESSIONE_INTRA_CLIENTE = 0
    End If
End If

rs.CloseResultset
Set rs = Nothing

FLAG_INTRASTAT_DOC = GET_CESSIONE_INTRA_CLIENTE
End Function
Private Function GET_CAMPI_INSTRASTAT_AZIENDA(IDAnagrafica As Long, IDAzienda As Long, IDFiliale As Long, DataDocumento As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

ObjDoc.Field "Link_Doc_intra_dati_richiesti", GET_LINK_DATI_RICHIESTI_INTRA(IDAzienda, DataDocumento), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'ObjDoc.Field "Link_Doc_intra_sezione", 1, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
ObjDoc.Field "Link_Doc_intra_modo_trasporto", GET_LINK_MODO_DI_TRASPORTO(IDAnagrafica, IDAzienda, IDFiliale), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)

''''''PROVINCIA E NAZIONE DI PAGAMENTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'sSQL = "SELECT IDComune, IDNazione FROM RepAzienda "
'sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
'Set rs = CnDMT.OpenResultset(sSQL)
'
'If rs.EOF Then
'    ObjDoc.Field "Link_Doc_intra_modo_trasporto", 0, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'    ObjDoc.Field "Link_Doc_intra_naz_pagamento", 0, NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'Else
'    ObjDoc.Field "Link_Doc_intra_provinc_merce", GET_LINK_PROVINCIA(fnNotNullN(rs!IDComune)), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'    ObjDoc.Field "Link_Doc_intra_naz_pagamento", fnNotNullN(rs!IDNazione), NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)
'End If
'
'rs.CloseResultset
'Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function
Private Function GET_LINK_PROVINCIA(IDComune As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDProvincia FROM Comune "
sSQL = sSQL & "WHERE IDComune=" & IDComune

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_PROVINCIA = 0
Else
    GET_LINK_PROVINCIA = fnNotNullN(rs!IDProvincia)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_NATURA_TRANSAZIONE(IDArticolo As Long, IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_LINK_NATURA_TRANSAZIONE = 0

'''''''''''''''''''''DA ARTICOLI''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT RV_POIDNaturaTransazione FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & "  AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_NATURA_TRANSAZIONE = 0
Else
    GET_LINK_NATURA_TRANSAZIONE = fnNotNullN(rs!RV_POIDNaturaTransazione)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If GET_LINK_NATURA_TRANSAZIONE > 0 Then Exit Function

'''''''''''''''''''''PARAMETRI FILIALE''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDNaturaTransazione FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & "  AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_NATURA_TRANSAZIONE = 0
Else
    GET_LINK_NATURA_TRANSAZIONE = fnNotNullN(rs!IDNaturaTransazione)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Function
Private Function GET_LINK_MODO_DI_TRASPORTO(IDAnagrafica As Long, IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_LINK_MODO_DI_TRASPORTO = 0

'''''''''''''''''''''DA ANAGRAFICA CLIENTE''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDModoTrasportoIntra FROM RV_POConfigurazioneCliente "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & "  AND IDAnagrafica=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_MODO_DI_TRASPORTO = 0
Else
    GET_LINK_MODO_DI_TRASPORTO = fnNotNullN(rs!IDModoTrasportoIntra)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If GET_LINK_MODO_DI_TRASPORTO > 0 Then Exit Function

'''''''''''''''''''''PARAMETRI FILIALE''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDModoTrasportoIntra FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & "  AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_MODO_DI_TRASPORTO = 0
Else
    GET_LINK_MODO_DI_TRASPORTO = fnNotNullN(rs!IDModoTrasportoIntra)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Function

Private Function GET_QUANTITA_MASSA_NETTA_INSTRASTAT(IDArticolo As Long, IDAzienda As Long, QuantitaVenduta As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MassaNettaInKg FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_QUANTITA_MASSA_NETTA_INSTRASTAT = 0
Else
    GET_QUANTITA_MASSA_NETTA_INSTRASTAT = QuantitaVenduta * fnNotNullN(rs!MassaNettaInKg)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_NOMENCLATURA_COMBINATA(IDArticolo As Long, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDNomenclaturaCombinata FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_NOMENCLATURA_COMBINATA = 0
Else
    GET_LINK_NOMENCLATURA_COMBINATA = fnNotNullN(rs!IDNomenclaturaCombinata)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_INTRAST_RIGA_ARTICOLO(IDArticolo As Long, IDAzienda As Long, IDFiliale As Long, Quantita As Double)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT NonRiportoIntrastat, MassaNettaInKg, IDNomenclaturaCombinata, RV_POIDNaturaTransazione "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    ObjDoc.Field "Art_intra_non_riporto", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Link_Art_intra_nomenclatura", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Link_Art_intra_natura_trans", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    ObjDoc.Field "Art_intra_qta_tot_massa_netta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)

Else
    If Abs(fnNotNullN(rs!NonRiportoIntrastat)) = 1 Then
        ObjDoc.Field "Art_intra_non_riporto", fnNotNullN(rs!NonRiportoIntrastat), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Link_Art_intra_nomenclatura", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Link_Art_intra_natura_trans", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_intra_qta_tot_massa_netta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    Else
        ObjDoc.Field "Art_intra_non_riporto", fnNotNullN(rs!NonRiportoIntrastat), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Link_Art_intra_nomenclatura", fnNotNullN(rs!IDNomenclaturaCombinata), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Link_Art_intra_natura_trans", GET_LINK_NATURA_TRANSAZIONE(IDArticolo, IDAzienda, IDFiliale), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        ObjDoc.Field "Art_intra_qta_tot_massa_netta", GET_QUANTITA_MASSA_NETTA_INSTRASTAT(IDArticolo, IDAzienda, Quantita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    End If

End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Function GET_LINK_AGENTE_CLIENTE(IDAnagrafica As Long, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagraficaAgente FROM ProvvClienteAgente "
sSQL = sSQL & " WHERE IDAziendaCliente=" & IDAzienda
sSQL = sSQL & " AND IDTipoAnagraficaCliente=2"
sSQL = sSQL & " AND IDAnagraficaCliente=" & IDAnagrafica

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_AGENTE_CLIENTE = 0
Else
    GET_LINK_AGENTE_CLIENTE = fnNotNullN(rs!IDAnagraficaAgente)
End If





rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_LINK_REGOLA_PROVV_AGE(IDAnagraficaAgente As Long, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRegolaProvv FROM RegolaProvvAgente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaAgente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND Predefinita=" & fnNormBoolean(1)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REGOLA_PROVV_AGE = 0
Else
    GET_LINK_REGOLA_PROVV_AGE = fnNotNullN(rs!IDRegolaProvv)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CODICE_AGE(IDAnagraficaAgente As Long, IDAzienda As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Codice FROM IERepAgente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaAgente
sSQL = sSQL & " AND IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CODICE_AGE = ""
Else
    GET_LINK_CODICE_AGE = fnNotNull(rs!Codice)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_REGOLA_PROVV(IDRegolaProvv As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RegolaProvv FROM RegolaProvv "
sSQL = sSQL & "WHERE IDRegolaProvv=" & IDRegolaProvv
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_REGOLA_PROVV = ""
Else
    GET_DESCRIZIONE_REGOLA_PROVV = fnNotNull(rs!regolaprovv)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_LINK_DATI_RICHIESTI_INTRA(IDAzienda As Long, DataDocumento As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim LINK_PERIODO_IVA As Long
Dim sDate As String

sDate = DataDocumento

If DataDocumento = "" Then
    sDate = Date
End If
If Val(DataDocumento) = 0 Then
    sDate = Date
End If

sSQL = "SELECT IDPeriodoIVA, IDPeriodicitaCessione FROM PeriodoIva "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND Anno=" & DatePart("yyyy", sDate)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_DATI_RICHIESTI_INTRA = 2
Else
    Select Case fnNotNullN(rs!IDPeriodicitaCessione)
        Case 1
            GET_LINK_DATI_RICHIESTI_INTRA = 1
        Case 2
            GET_LINK_DATI_RICHIESTI_INTRA = 2
        Case 3
            GET_LINK_DATI_RICHIESTI_INTRA = 2
        Case 4
            GET_LINK_DATI_RICHIESTI_INTRA = 2
        Case 6
            GET_LINK_DATI_RICHIESTI_INTRA = 2
        Case 5
            GET_LINK_DATI_RICHIESTI_INTRA = 2
        Case 6
            GET_LINK_DATI_RICHIESTI_INTRA = 2
        Case 7
            GET_LINK_DATI_RICHIESTI_INTRA = 2
        Case Else
            GET_LINK_DATI_RICHIESTI_INTRA = 2

    End Select
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub VScroll1_Change()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
   Me.Pic1.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
   Me.Pic1.Left = -HScroll1.Value
End Sub
Private Function GET_NUMERO_ORDINI_DA_EVADERE(IDUtente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(NumeroRiga) AS NumeroOrdiniSel "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini"
sSQL = sSQL & " WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND DaRegistrare=1"


Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_ORDINI_DA_EVADERE = 0
Else
    GET_NUMERO_ORDINI_DA_EVADERE = fnNotNullN(rs!NumeroOrdiniSel)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_PREFISSO_SEZ(IDSezionale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Prefisso FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDSezionale=" & IDSezionale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREFISSO_SEZ = ""
Else
    GET_PREFISSO_SEZ = Trim(fnNotNull(rs!Prefisso))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GetNumeroDocumentoModificato(NumeroDocumento As String) As String
Const Totale As Integer = 6
Dim I As Integer
Dim Count As Integer

GetNumeroDocumentoModificato = ""
For I = 1 To (Totale - Len(NumeroDocumento))
    GetNumeroDocumentoModificato = GetNumeroDocumentoModificato & "0"
Next

GetNumeroDocumentoModificato = GetNumeroDocumentoModificato & NumeroDocumento

End Function
Private Sub fnAggiornaDescrizioneDocumento(LetteraSezionale As String, NomeTabellaDettaglio As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_PODescrizioneDocumento="

Select Case ObjDoc.IDTipoOggetto
    Case 2
        If NUMERO_ZERI_DOC_RIF = 0 Then
            sSQL = sSQL & fnNormString("Rif. D.d.t. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(ObjDoc.Numero) & " del " & ObjDoc.DataEmissione)
        Else
            sSQL = sSQL & fnNormString("Rif. D.d.t. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & ObjDoc.Numero & " del " & ObjDoc.DataEmissione)
        End If
    Case 114
        If NUMERO_ZERI_DOC_RIF = 0 Then
            sSQL = sSQL & fnNormString("Rif. f.a. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(ObjDoc.Numero) & " del " & ObjDoc.DataEmissione)
        Else
            sSQL = sSQL & fnNormString("Rif. f.a. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & ObjDoc.Numero & " del " & ObjDoc.DataEmissione)
        End If
    Case 8
        If NUMERO_ZERI_DOC_RIF = 0 Then
            sSQL = sSQL & fnNormString("Rif. s.n.f. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(ObjDoc.Numero) & " del " & ObjDoc.DataEmissione)
        Else
            sSQL = sSQL & fnNormString("Rif. s.n.f. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & ObjDoc.Numero & " del " & ObjDoc.DataEmissione)
        End If
End Select

sSQL = sSQL & "WHERE IDOggetto=" & ObjDoc.IDOggetto

CnDMT.Execute sSQL
End Sub
Private Sub fnAggiornaDescrizioneDocumentoOrd(LetteraSezionale As String, NomeTabellaDettaglio As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_PODescrizioneDocumentoOrdinamento="

Select Case ObjDoc.IDTipoOggetto
    Case 2
        sSQL = sSQL & fnNormString("Rif. D.d.t. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(ObjDoc.Numero) & " del " & ObjDoc.DataEmissione)
    Case 114
        sSQL = sSQL & fnNormString("Rif. f.a. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(ObjDoc.Numero) & " del " & ObjDoc.DataEmissione)
    Case 8
        sSQL = sSQL & fnNormString("Rif. s.n.f. n: " & IIf(Len(Trim(LetteraSezionale)) > 0, LetteraSezionale & "/", "") & GetNumeroDocumentoModificato(ObjDoc.Numero) & " del " & ObjDoc.DataEmissione)
End Select

sSQL = sSQL & "WHERE IDOggetto=" & ObjDoc.IDOggetto

CnDMT.Execute sSQL

End Sub
Private Sub fnAggiornaDescrizioneOrdineCliente(NomeTabellaDettaglio As String, DataOrdine As String, NumeroOrdine As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_POOggettoOrdineCliente="
sSQL = sSQL & fnNormString("Rif. Vs Ordine: " & GetNumeroDocumentoModificato(NumeroOrdine) & " del " & DataOrdine)
sSQL = sSQL & "WHERE IDOggetto=" & ObjDoc.IDOggetto

CnDMT.Execute sSQL


End Sub
Private Sub fnAggiornaDescrizioneOrdineInterno(NomeTabellaDettaglio As String, DataOrdine As String, NumeroOrdine As String)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_POOggettoOrdineInterno="
sSQL = sSQL & fnNormString("Rif. Ns Ordine: " & GetNumeroDocumentoModificato(NumeroOrdine) & " del " & DataOrdine)
sSQL = sSQL & "WHERE IDOggetto=" & ObjDoc.IDOggetto

CnDMT.Execute sSQL

End Sub
Private Sub fnAggiornaDescrizioneProtocollo(NomeTabellaDettaglio As String, Descrizione As String, NumeroProtocollo As Long)
Dim sSQL As String

sSQL = "UPDATE " & NomeTabellaDettaglio & " SET "
sSQL = sSQL & "RV_PODescrizioneProtocolloICE=" & fnNormString(Descrizione) & ", "
sSQL = sSQL & "RV_PONumeroProtocolloICE=" & NumeroProtocollo
sSQL = sSQL & " WHERE IDOggetto=" & ObjDoc.IDOggetto

CnDMT.Execute sSQL

End Sub
Private Function GET_PREZZO_MEDIO_CLIENTE(IDAnagrafica As Long, IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim EsisteDettaglio As Boolean

EsisteDettaglio = False

sSQL = "SELECT IDArticolo, RV_PONonPartecipaPrezzoMedio "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!RV_PONonPartecipaPrezzoMedio) = 1 Then
        GET_PREZZO_MEDIO_CLIENTE = 0
    Else
        GET_PREZZO_MEDIO_CLIENTE = 1
    End If
    
    EsisteDettaglio = True
End If

rs.CloseResultset
Set rs = Nothing


If EsisteDettaglio = True Then Exit Function

sSQL = "SELECT NonCalcolarePrezzoMedio "
sSQL = sSQL & " FROM RV_POConfigurazioneClienteArtVend"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!NonCalcolarePrezzoMedio) = 1 Then
        GET_PREZZO_MEDIO_CLIENTE = 0
    Else
        GET_PREZZO_MEDIO_CLIENTE = 1
    End If
    
    EsisteDettaglio = True
End If

rs.CloseResultset
Set rs = Nothing

If EsisteDettaglio = True Then Exit Function

sSQL = "SELECT NonCalcolarePrezzoMedio "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_MEDIO_CLIENTE = 1
Else
    If fnNotNullN(rs!NonCalcolarePrezzoMedio) = 1 Then
        GET_PREZZO_MEDIO_CLIENTE = 0
    Else
        GET_PREZZO_MEDIO_CLIENTE = 1
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_FORZATURA_PREZZO_LIQ_CLIENTE(IDAnagrafica As Long, IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim EsisteDettaglio As Boolean

EsisteDettaglio = False


sSQL = "SELECT IDRV_POTipoImportoVenditaLiq "
sSQL = sSQL & " FROM RV_POConfigurazioneClienteArtVend"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = fnNotNullN(rs!IDRV_POTipoImportoVenditaLiq)
    
    EsisteDettaglio = True
End If

rs.CloseResultset
Set rs = Nothing

If EsisteDettaglio = True Then Exit Function


sSQL = "SELECT IDRV_POTipoImportoVenditaLiq "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = 0
Else
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = fnNotNullN(rs!IDRV_POTipoImportoVenditaLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_MESSAGGIO_PER_ANAGRAFICA(IDAnagrafica As Long) As String
On Error GoTo ERR_GET_MESSAGGIO_PER_ANAGRAFICA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM MessaggioPerAnagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoMessaggio=1"
sSQL = sSQL & " ORDER BY DataUltimaVariazione DESC, IDMessaggioPerAnagrafica DESC"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_MESSAGGIO_PER_ANAGRAFICA = ""
Else
    GET_MESSAGGIO_PER_ANAGRAFICA = fnNotNull(rs!MessaggioPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_MESSAGGIO_PER_ANAGRAFICA:
    MsgBox Err.Description, vbCritical, "GET_MESSAGGIO_PER_ANAGRAFICA"
End Function
Private Function GET_DATI_RIGA_CONFERIMENTO(IDRigaConferimento As Long, NomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_DATI_RIGA_CONFERIMENTO = 0
Else
    GET_DATI_RIGA_CONFERIMENTO = fnNotNullN(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ParametroAggiornaPrezzoMedioDaConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AggiornaPrezzoMedioDaConf, IvaArticoloDaDocumentoCollegato, LetteraIntentoDaDocumentoCollegato"
sSQL = sSQL & " FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    AGGIORNA_PREZZO_MEDIO = fnNotNullN(rs!AggiornaPrezzoMedioDaConf)
    RIP_IVA_DA_DOC_COLL = fnNotNullN(rs!IvaArticoloDaDocumentoCollegato)
    RIP_LET_INT_DA_DOC_COLL = fnNotNullN(rs!LetteraIntentoDaDocumentoCollegato)
Else
    AGGIORNA_PREZZO_MEDIO = 0
    RIP_IVA_DA_DOC_COLL = 0
    RIP_LET_INT_DA_DOC_COLL = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroAggiornaTipoLavorazioneDaConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AggiornaTipoLavDaConf FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    AGGIORNA_TIPO_LAVORAZIONE = fnNotNullN(rs!AggiornaTipoLavDaConf)
Else
    AGGIORNA_TIPO_LAVORAZIONE = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_LINK_TIPO_PEDANA(IDPedana As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPedana FROM RV_POPedana "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_TIPO_PEDANA = 0
Else
    GET_LINK_TIPO_PEDANA = fnNotNullN(rs!IDRV_POTipoPedana)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PESO_PEDANA(IDPedana As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PesoPedana FROM RV_POPedana "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_PESO_PEDANA = 0
Else
    GET_PESO_PEDANA = fnNotNullN(rs!PesoPedana)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub AGGIORNA_RIGHE_COMMISSIONI(IDOggetto As Long, IDTipoOggetto As Long, NomeTabella As String, DataDocumento As String)
On Error GoTo ERR_AGGIORNA_RIGHE_COMMISSIONI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim rscomm As DmtOleDbLib.adoResultset

'ELIMINAZIONE DELLE RIGHE PER COMMISSIONI
sSQL = "DELETE FROM RV_POCommissioniPerDocRighe "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
CnDMT.OpenResultset sSQL


sSQL = "SELECT RV_POCommissioniPerDoc.IDRV_POCommissioniPerDoc, RV_POCommissioniPerDoc.IDOggetto, RV_POCommissioniPerDoc.IDRV_POTipoCommissione, RV_POCommissioniPerDoc.Percentuale, "
sSQL = sSQL & "RV_POCommissioniPerDoc.Importo, RV_POCommissioniPerDoc.ImportoRiga, RV_POCommissioniPerDoc.Quantita, RV_POCommissioniPerDoc.APercentuale, RV_POCommissioniPerDoc.ImportoTotale,"
sSQL = sSQL & "RV_POCommissioniPerDoc.IDRV_POTipoPedana , RV_POCommissioniPerDoc.PercentualeDaCommissione, RV_POCommissioniPerDoc.IDArticoloImballo, RV_POTipoCommissione.IDRV_POTipoValoreDocumento "
sSQL = sSQL & "FROM RV_POCommissioniPerDoc INNER JOIN "
sSQL = sSQL & "RV_POTipoCommissione ON RV_POCommissioniPerDoc.IDRV_POTipoCommissione = RV_POTipoCommissione.IDRV_POTipoCommissione "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND ((IDRV_POTipoPedana=0) OR (IDRV_POTipoPedana IS NULL))"
sSQL = sSQL & " AND ((APercentuale=0) OR (APercentuale IS NULL))"
Set rscomm = CnDMT.OpenResultset(sSQL)

If rscomm.EOF Then
    rscomm.CloseResultset
    Set rscomm = Nothing
    Exit Sub
End If

'SETTAGGIO DEL RECORDSET PER LE RIGHE COMMISSIONI'''''''''''''''''''''''''''''
Set rsNew = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POCommissioniPerDocRighe "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT IDValoriOggettoDettaglio, RV_PODataLavorazione, RV_PODataConferimento, "
sSQL = sSQL & "Art_pre_uni_net_sco_net_IVA, Link_art_articolo, art_quantita_totale, RV_POImportoMerceNetta, RV_POQuantitaLiq "
sSQL = sSQL & "FROM " & NomeTabella
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    CALCOLO_RIGHE_COMMISSIONI IDOggetto, IDTipoOggetto, DataDocumento, rsNew, rs, rscomm
rs.MoveNext
Wend

rscomm.CloseResultset
Set rscomm = Nothing

rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_AGGIORNA_RIGHE_COMMISSIONI:
    MsgBox Err.Description, vbCritical, "AGGIORNA_RIGHE_COMMISSIONI"
End Sub
Private Function CALCOLO_RIGHE_COMMISSIONI(IDOggetto As Long, IDTipoOggetto As Long, DataVendita As String, rsNew As ADODB.Recordset, rs As DmtOleDbLib.adoResultset, rscomm As DmtOleDbLib.adoResultset)
Dim ImportoMerceNetta As Double

rscomm.MoveFirst

While Not rscomm.EOF
    If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
        If fnNotNullN(rscomm!IDRV_POTipoValoreDocumento) < 5 Then
            ImportoMerceNetta = (fnNotNullN(rs!RV_POImportoMerceNetta) * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
        Else
            ImportoMerceNetta = (fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA) * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
        End If
        rsNew.AddNew
            rsNew!IDRV_POCommissioniPerDocRighe = fnGetNewKey("RV_POCommissioniPerDocRighe", "IDRV_POCommissioniPerDocRighe")
            rsNew!IDRV_POCommissioniPerDoc = fnNotNullN(rscomm!IDRV_POCommissioniPerDoc)
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = IDTipoOggetto
            rsNew!IDValoriOggettoDettaglio = rs!IDValoriOggettoDettaglio
            rsNew!IDRV_POTipoCommissione = fnNotNullN(rscomm!IDRV_POTipoCommissione)
            rsNew!Percentuale = fnNotNullN(rscomm!Percentuale)
            rsNew!DataVendita = DataVendita
            rsNew!DataConferimento = rs!RV_PODataConferimento
            rsNew!DataLavorazione = rs!RV_PODataLavorazione
            'rsNew!Importo = (fnNotNullN(rs!RV_POImportoMerceNetta) / 100) * fnNotNullN(rsNew!Percentuale)
            rsNew!Importo = (ImportoMerceNetta / 100) * fnNotNullN(rsNew!Percentuale)
            rsNew!IDArticolo = fnNotNullN(rs!Link_Art_articolo)
            'rsNew!Quantita = fnNotNullN(rs!Art_quantita_totale)
            rsNew!Quantita = fnNotNullN(rs!RV_POQuantitaLiq)
        rsNew.Update
    End If
rscomm.MoveNext
Wend

End Function
Private Function GET_TIPO_DOCUMENTO_CLIENTE(IDUtente As Long, IDTipoDocumentoAzienda As Long) As Long
On Error GoTo ERR_GET_TIPO_DOCUMENTO_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Numero As Long

Dim rsCliente As ADODB.Recordset

Set rsCliente = New ADODB.Recordset
rsCliente.CursorLocation = adUseClient

rsCliente.Fields.Append "IDCliente", adInteger, , adFldIsNullable

rsCliente.Open , , adOpenKeyset, adLockBatchOptimistic


GET_TIPO_DOCUMENTO_CLIENTE = 0

sSQL = "SELECT IDCliente AS NumeroOrdiniSel "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini"
sSQL = sSQL & " WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND DaRegistrare=1"


Set rs = CnDMT.OpenResultset(sSQL)
Numero = 0
While Not rs.EOF
    
    rsCliente.Filter = "IDCliente=" & fnNotNullN(rs!NumeroOrdiniSel)
    
    If rsCliente.EOF Then
        rsCliente.AddNew
            rsCliente!IDCliente = fnNotNullN(rs!NumeroOrdiniSel)
            Numero = Numero + 1
        rsCliente.Update
    
    End If
    
    rsCliente.Filter = vbNullString
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If Numero = 0 Then Exit Function

If Numero = 1 Then
    rsCliente.MoveFirst
    
    sSQL = "SELECT IDTipoOggettoDocEvasione FROM Cliente "
    sSQL = sSQL & "WHERE IDAnagrafica=" & fnNotNullN(rsCliente!IDCliente)
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_TIPO_DOCUMENTO_CLIENTE = 0
    Else
        GET_TIPO_DOCUMENTO_CLIENTE = fnNotNullN(rs!IDTipoOggettoDocEvasione)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    If GET_TIPO_DOCUMENTO_CLIENTE = 0 Then
        GET_TIPO_DOCUMENTO_CLIENTE = IDTipoDocumentoAzienda
    End If
End If

rsCliente.Close
Set rsCliente = Nothing



Exit Function
ERR_GET_TIPO_DOCUMENTO_CLIENTE:
    MsgBox Err.Description, vbCritical, "Selezione documento predefinito"
    GET_TIPO_DOCUMENTO_CLIENTE = 0
End Function
Private Sub GET_PESO_PEDANA_IN_VENDITA()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PesoArticoloPedanaInFattura FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    PESO_PEDANA_IN_VENDITA = fnNotNullN(rs!PesoArticoloPedanaInFattura)
Else
    PESO_PEDANA_IN_VENDITA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_PESO_ARTICOLO(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PesoNetto FROM Articolo"
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_PESO_ARTICOLO = fnNotNullN(rs!PesoNetto)
Else
    GET_PESO_ARTICOLO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_IMBALLI_NOLEGGIO(rsImbTmp As ADODB.Recordset, IDImballoVendita As Long, Quantita As Double, IDAgente As Long, IDCliente As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POConfigurazioneClienteImbCauz "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticoloImballo=" & IDImballoVendita

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF

    rsImbTmp.Filter = "IDArticoloImballo=" & fnNotNullN(rs!IDArticoloImballoNolCauz)
    rsImbTmp.Filter = rsImbTmp.Filter & " AND IDListino=" & fnNotNullN(rs!IDListinoNolCauz)
    
    If rsImbTmp.EOF Then
        rsImbTmp.AddNew
        rsImbTmp!IDArticoloImballo = fnNotNullN(rs!IDArticoloImballoNolCauz)
    End If
        rsImbTmp!Quantita = fnNotNullN(rsImbTmp!Quantita) + Quantita
        rsImbTmp!IDAgente = IDAgente
        rsImbTmp!IDListino = fnNotNullN(rs!IDListinoNolCauz)
    
    rsImbTmp.Update
        
    rsImbTmp.Filter = vbNullString

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA(tabella As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsPed As DmtOleDbLib.adoResultset
Dim rsRiga As ADODB.Recordset
Dim Moltiplicatore As Double
Dim Prezzo As Double
Dim ImportoCommPerPedana As Double
Dim rsNew As ADODB.Recordset
Dim ImportoCommArticoloPerPedana As Double


'SETTAGGIO DEL RECORDSET PER LE RIGHE COMMISSIONI'''''''''''''''''''''''''''''
Set rsNew = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POCommissioniPerDocRighe "
sSQL = sSQL & " WHERE IDOggetto=" & ObjDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
sSQL = sSQL & " WHERE IDOggetto=" & ObjDoc.IDOggetto
sSQL = sSQL & " AND IDRV_POTipoPedana>0"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "SELECT RV_POIDPedana, SUM(Art_peso - Art_tara) AS TotalePesoNetto, SUM(Art_volume*Art_numero_colli) AS TotaleVolume "
    sSQL = sSQL & " FROM " & tabella
    sSQL = sSQL & " WHERE IDOggetto = " & ObjDoc.IDOggetto
    sSQL = sSQL & " AND RV_POIDTipoPedana = " & fnNotNullN(rs!IDRV_POTipoPedana)
    sSQL = sSQL & " AND RV_POTipoRiga = 1 "
    sSQL = sSQL & " GROUP BY RV_POIDPedana"

    Set rsPed = CnDMT.OpenResultset(sSQL)
    
    While Not rsPed.EOF
        If TIPO_CALCOLO_TRASPORTO = 1 Then
            If fnNotNullN(rsPed!TotalePesoNetto) > 0 Then
                
                sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, RV_POPrezzoMedioInLiq, RV_POVariazionePrezzoImballo, RV_POImportoMerceNetta, RV_POImportoLiq, "
                sSQL = sSQL & "RV_POPrezzoUnitarioOrigine, RV_POIDIvaImballo, RV_POIDTipoVariazione, Art_pre_uni_net_sco_net_IVA ,"
                sSQL = sSQL & "RV_POImportoDaLiq, RV_POLinkRiga, RV_POImportoImballoInArticolo, link_art_articolo, "
                sSQL = sSQL & "Art_numero_colli, Art_quantita_Totale, Link_art_articolo, (Art_Peso - Art_Tara) AS PesoNettoRiga, Art_Volume,  (Art_volume*Art_numero_colli) VolumeRigaVendita, "
                sSQL = sSQL & "RV_POVariazionePrezzoManuale, RV_POImportoRigaCommissioni, RV_PODataLavorazione, RV_PODataConferimento "
                sSQL = sSQL & " FROM " & tabella
                sSQL = sSQL & " WHERE RV_POTipoRiga=1 "
                sSQL = sSQL & " AND IDOggetto=" & ObjDoc.IDOggetto
                sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto
                sSQL = sSQL & " AND RV_POIDPedana=" & fnNotNullN(rsPed!RV_POIDPedana)
                
                Set rsRiga = New ADODB.Recordset
                rsRiga.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
                
                While Not rsRiga.EOF
                
                    If fnNotNullN(rsRiga!PesoNettoRiga) > 0 Then
                        
                        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rsRiga!Link_Art_articolo))
                        
                        ImportoCommArticoloPerPedana = (fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotalePesoNetto)) * fnNotNullN(rsRiga!PesoNettoRiga)

                        ImportoCommArticoloPerPedana = ImportoCommArticoloPerPedana / fnNotNullN(rsRiga!Art_quantita_totale) / Moltiplicatore
                        
                        rsRiga!RV_POImportoLiq = rsRiga!RV_POImportoLiq - Abs(ImportoCommArticoloPerPedana)
                        rsRiga!RV_POImportoRigaCommissioni = rsRiga!RV_POImportoRigaCommissioni + Abs(ImportoCommArticoloPerPedana)
                        
                        rsRiga.Update
                        
                        rsNew.AddNew
                            rsNew!IDRV_POCommissioniPerDocRighe = fnGetNewKey("RV_POCommissioniPerDocRighe", "IDRV_POCommissioniPerDocRighe")
                            rsNew!IDRV_POCommissioniPerDoc = fnNotNullN(rs!IDRV_POCommissioniPerDoc)
                            rsNew!IDOggetto = ObjDoc.IDOggetto
                            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
                            rsNew!IDValoriOggettoDettaglio = rsRiga!IDValoriOggettoDettaglio
                            rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
                            rsNew!Percentuale = 0
                            rsNew!DataVendita = ObjDoc.DataEmissione
                            rsNew!DataConferimento = rsRiga!RV_PODataConferimento
                            rsNew!DataLavorazione = rsRiga!RV_PODataLavorazione
                            rsNew!Importo = ImportoCommArticoloPerPedana '(fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotalePesoNetto)) * fnNotNullN(rsRiga!PesoNettoRiga)
                            rsNew!IDArticolo = fnNotNullN(rsRiga!Link_Art_articolo)
                            rsNew!Quantita = fnNotNullN(rsRiga!Art_quantita_totale)
                            
                        rsNew.Update
                        
                    End If
                    
                rsRiga.MoveNext
                Wend
                
                rsRiga.Close
                Set rsRiga = Nothing
            End If
        Else
            
            If fnNotNullN(rsPed!TotaleVolume) > 0 Then
                
                sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, RV_POPrezzoMedioInLiq, RV_POVariazionePrezzoImballo, RV_POImportoMerceNetta, RV_POImportoLiq, "
                sSQL = sSQL & "RV_POPrezzoUnitarioOrigine, RV_POIDIvaImballo, RV_POIDTipoVariazione, Art_pre_uni_net_sco_net_IVA ,"
                sSQL = sSQL & "RV_POImportoDaLiq, RV_POLinkRiga, RV_POImportoImballoInArticolo, link_art_articolo, "
                sSQL = sSQL & "Art_numero_colli, Art_quantita_Totale, Link_art_articolo, (Art_Peso - Art_Tara) AS PesoNettoRiga, Art_volume, (Art_volume*Art_numero_colli) VolumeRigaVendita, "
                sSQL = sSQL & "RV_POVariazionePrezzoManuale, RV_POImportoRigaCommissioni, RV_PODataLavorazione, RV_PODataConferimento, RV_POImportoLiqDoc "
                sSQL = sSQL & " FROM " & tabella
                sSQL = sSQL & " WHERE RV_POTipoRiga=1 "
                sSQL = sSQL & " AND IDOggetto=" & ObjDoc.IDOggetto
                sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto
                sSQL = sSQL & " AND RV_POIDPedana=" & fnNotNullN(rsPed!RV_POIDPedana)
                
                Set rsRiga = New ADODB.Recordset
                rsRiga.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
                
                While Not rsRiga.EOF
                
                    If fnNotNullN(rsRiga!VolumeRigaVendita) <> 0 Then
                        
                        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rsRiga!Link_Art_articolo))
                        
                        
                        ImportoCommArticoloPerPedana = (fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotaleVolume)) * fnNotNullN(rsRiga!VolumeRigaVendita)
                        
                        
                        ImportoCommArticoloPerPedana = ImportoCommArticoloPerPedana / fnNotNullN(rsRiga!Art_quantita_totale) / Moltiplicatore
                        
                        rsRiga!RV_POImportoLiq = rsRiga!RV_POImportoLiq - Abs(ImportoCommArticoloPerPedana)
                        rsRiga!RV_POImportoRigaCommissioni = rsRiga!RV_POImportoRigaCommissioni + Abs(ImportoCommArticoloPerPedana)
                        rsRiga!RV_POImportoLiqDoc = rsRiga!RV_POImportoLiq
                        rsRiga.Update
                        
                        rsNew.AddNew
                            rsNew!IDRV_POCommissioniPerDocRighe = fnGetNewKey("RV_POCommissioniPerDocRighe", "IDRV_POCommissioniPerDocRighe")
                            rsNew!IDRV_POCommissioniPerDoc = fnNotNullN(rs!IDRV_POCommissioniPerDoc)
                            rsNew!IDOggetto = ObjDoc.IDOggetto
                            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
                            rsNew!IDValoriOggettoDettaglio = rsRiga!IDValoriOggettoDettaglio
                            rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
                            rsNew!Percentuale = 0
                            rsNew!DataVendita = ObjDoc.DataEmissione
                            rsNew!DataConferimento = rsRiga!RV_PODataConferimento
                            rsNew!DataLavorazione = rsRiga!RV_PODataLavorazione
                            rsNew!Importo = ImportoCommArticoloPerPedana '(fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotalePesoNetto)) * fnNotNullN(rsRiga!PesoNettoRiga)
                            rsNew!IDArticolo = fnNotNullN(rsRiga!Link_Art_articolo)
                            rsNew!Quantita = fnNotNullN(rsRiga!Art_quantita_totale)
                            rsNew!VolumeUnitariaImballo = fnNotNullN(rsRiga!Art_volume)
                            rsNew!VolumeTotaleMerce = fnNotNullN(rsRiga!VolumeRigaVendita)
                            rsNew!VolumeTotalePedana = fnNotNullN(rsPed!TotaleVolume)
                        rsNew.Update
                        
                    End If
                    
                rsRiga.MoveNext
                Wend
                
                rsRiga.Close
                Set rsRiga = Nothing
                                
            End If
        End If
        
    rsPed.MoveNext
    Wend
    
    rsPed.CloseResultset
    Set rsPed = Nothing

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
End Sub
Private Function GET_LINK_TIPO_PEDANA_DA_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_TIPO_PEDANA_DA_ARTICOLO = 0

sSQL = "SELECT IDRV_POTipoPedana FROM RV_POTipoPedana "
sSQL = sSQL & " WHERE IDArticoloImballo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_TIPO_PEDANA_DA_ARTICOLO = fnNotNullN(rs!IDRV_POTipoPedana)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_RIGA_PIANALE(IDOggettoOrdine As Long, I As Integer, Link_Riga As Long, ProgressivoArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImballoARendere As Long
Dim LINK_REGOLA_PROVV As Long

    sSQL = sSQL & "SELECT ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, ValoriOggettoDettaglio0010.RV_PO01_IDArticoloPianale AS IDArticolo, SUM(ValoriOggettoDettaglio0010.RV_PO01_QuantitaPianale) "
    sSQL = sSQL & "AS Quantita, ValoriOggettoDettaglio0010.RV_POTipoRiga, Articolo.CodiceArticolo, Articolo.Articolo "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0010.RV_PO01_IDArticoloPianale = Articolo.IDArticolo "
    sSQL = sSQL & "GROUP BY ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, ValoriOggettoDettaglio0010.RV_PO01_IDArticoloPianale, ValoriOggettoDettaglio0010.RV_POTipoRiga, Articolo.CodiceArticolo, "
    sSQL = sSQL & "Articolo.Articolo "
    sSQL = sSQL & " HAVING IDOggetto=" & IDOggettoOrdine
    sSQL = sSQL & " AND IDTipoOggetto=15 "
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        If (fnNotNullN(rs!IDArticolo)) > 0 Then
            
            GET_IVA_ARTICOLO fnNotNullN(rs!IDArticolo)
            ImballoARendere = GET_TIPO_IMBALLO_A_RENDERE(fnNotNullN(rs!IDArticolo))
            

            ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                        
            ObjDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_codice", fnNotNull(rs!CodiceArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_descrizione", fnNotNull(rs!Articolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Quantita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_prezzo_unitario_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_lordo_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_pre_uni_net_sco_net_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_pre_uni_net_sco_lor_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_Importo_totale_neutro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_neutro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        
            ObjDoc.Field "Art_Importo_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
        
            If (ImballoARendere = 1) And (FLAG_IVA_IMBALLO_A_RENDERE = 1) Then
                ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            Else
                If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                    ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA_ARTICOLO(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                End If
            End If
                
                
            ObjDoc.Field "Art_tara", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_importo_totale_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ''''AGENTE
            If ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) > 0 Then
                LINK_REGOLA_PROVV = GET_LINK_REGOLA_PROVV_AGE(ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), TheApp.IDFirm)
                
                If LINK_REGOLA_PROVV > 0 Then
                    ObjDoc.Field "Link_Art_agente", ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_nome", ObjDoc.Field("Doc_age_nome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_codice", ObjDoc.Field("Doc_age_codice", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_ragione_sociale", ObjDoc.Field("Doc_age_ragione_sociale", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    ObjDoc.Field "Link_Art_age_regola_provv", LINK_REGOLA_PROVV, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_regola_provv", GET_DESCRIZIONE_REGOLA_PROVV(LINK_REGOLA_PROVV), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    If fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))) > 0 Then
                        ObjDoc.Field "Link_Art_age_tipo_ordine", GET_LINK_TIPO_ORDINE(LINK_REGOLA_PROVV, fnNotNullN(rs!Sconto1), fnNotNullN(rs!Sconto2)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
            End If
            
            If FLAG_INTRASTAT_DOC = 1 Then
                GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDArticolo), TheApp.IDFirm, TheApp.Branch, ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
            End If
                
            ObjDoc.Field "Link_Art_unita_di_misura", 8, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POTipoRiga", 2, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PORigaCompleta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDImballo", fnNotNullN(rs!IDArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ProgressivoArticolo = ProgressivoArticolo + 1
            ObjDoc.Field "ID_Art_dettaglio_prog", ProgressivoArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_POIDCalibro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoCategoria", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoLavorazione", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PODataConferimento", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDConferimentoRighe", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDSocio", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PONomeSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POLottoCampagna", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceLotto", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            Link_Riga = Link_Riga + 1
            I = I + 1
        End If
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_RIGA_PEDANA(IDOggettoOrdine As Long, I As Integer, Link_Riga As Long, ProgressivoArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImballoARendere As Long
Dim LINK_REGOLA_PROVV As Long
    sSQL = sSQL & "SELECT ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, ValoriOggettoDettaglio0010.RV_POIDArticoloPedana AS IDArticolo, SUM(ValoriOggettoDettaglio0010.RV_POQuantitaPedana) "
    sSQL = sSQL & "AS Quantita, ValoriOggettoDettaglio0010.RV_POTipoRiga, Articolo.CodiceArticolo, Articolo.Articolo "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0010.RV_POIDArticoloPedana = Articolo.IDArticolo "
    sSQL = sSQL & "GROUP BY ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, ValoriOggettoDettaglio0010.RV_POIDArticoloPedana, ValoriOggettoDettaglio0010.RV_POTipoRiga, Articolo.CodiceArticolo, "
    sSQL = sSQL & "Articolo.Articolo "
    sSQL = sSQL & " HAVING IDOggetto=" & IDOggettoOrdine
    sSQL = sSQL & " AND IDTipoOggetto=15 "
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        If (fnNotNullN(rs!IDArticolo)) > 0 Then
            
            GET_IVA_ARTICOLO fnNotNullN(rs!IDArticolo)
            ImballoARendere = GET_TIPO_IMBALLO_A_RENDERE(fnNotNullN(rs!IDArticolo))
            

            ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                        
            ObjDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_codice", fnNotNull(rs!CodiceArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_descrizione", fnNotNull(rs!Articolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Quantita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_prezzo_unitario_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_lordo_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_pre_uni_net_sco_net_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_pre_uni_net_sco_lor_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_Importo_totale_neutro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_neutro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        
            ObjDoc.Field "Art_Importo_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
        
            If (ImballoARendere = 1) And (FLAG_IVA_IMBALLO_A_RENDERE = 1) Then
                ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            Else
                If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                    ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA_ARTICOLO(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                End If
            End If
                
                
            ObjDoc.Field "Art_tara", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_importo_totale_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ''''AGENTE
            If ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) > 0 Then
                LINK_REGOLA_PROVV = GET_LINK_REGOLA_PROVV_AGE(ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), TheApp.IDFirm)
                
                If LINK_REGOLA_PROVV > 0 Then
                    ObjDoc.Field "Link_Art_agente", ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_nome", ObjDoc.Field("Doc_age_nome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_codice", ObjDoc.Field("Doc_age_codice", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_ragione_sociale", ObjDoc.Field("Doc_age_ragione_sociale", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    ObjDoc.Field "Link_Art_age_regola_provv", LINK_REGOLA_PROVV, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_regola_provv", GET_DESCRIZIONE_REGOLA_PROVV(LINK_REGOLA_PROVV), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    If fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))) > 0 Then
                        ObjDoc.Field "Link_Art_age_tipo_ordine", GET_LINK_TIPO_ORDINE(LINK_REGOLA_PROVV, fnNotNullN(rs!Sconto1), fnNotNullN(rs!Sconto2)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
            End If
            
            If FLAG_INTRASTAT_DOC = 1 Then
                GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDArticolo), TheApp.IDFirm, TheApp.Branch, ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
            End If
                
            ObjDoc.Field "Link_Art_unita_di_misura", 8, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POTipoRiga", 2, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PORigaCompleta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDImballo", fnNotNullN(rs!IDArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ProgressivoArticolo = ProgressivoArticolo + 1
            ObjDoc.Field "ID_Art_dettaglio_prog", ProgressivoArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_POIDCalibro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoCategoria", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoLavorazione", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PODataConferimento", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDConferimentoRighe", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDSocio", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PONomeSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POLottoCampagna", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceLotto", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            Link_Riga = Link_Riga + 1
            I = I + 1
        End If
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub GET_RIGA_PROLUNGA(IDOggettoOrdine As Long, I As Integer, Link_Riga As Long, ProgressivoArticolo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImballoARendere As Long
Dim LINK_REGOLA_PROVV As Long

    sSQL = sSQL & "SELECT ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, ValoriOggettoDettaglio0010.RV_PO01_IDArticoloProlunga AS IDArticolo, SUM(ValoriOggettoDettaglio0010.RV_PO01_QuantitaProlunga) "
    sSQL = sSQL & "AS Quantita, ValoriOggettoDettaglio0010.RV_POTipoRiga, Articolo.CodiceArticolo, Articolo.Articolo "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 INNER JOIN "
    sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0010.RV_PO01_IDArticoloProlunga = Articolo.IDArticolo "
    sSQL = sSQL & "GROUP BY ValoriOggettoDettaglio0010.IDOggetto, ValoriOggettoDettaglio0010.IDTipoOggetto, ValoriOggettoDettaglio0010.RV_PO01_IDArticoloProlunga, ValoriOggettoDettaglio0010.RV_POTipoRiga, Articolo.CodiceArticolo, "
    sSQL = sSQL & "Articolo.Articolo "
    sSQL = sSQL & " HAVING IDOggetto=" & IDOggettoOrdine
    sSQL = sSQL & " AND IDTipoOggetto=15 "
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    While Not rs.EOF
        If (fnNotNullN(rs!IDArticolo)) > 0 Then
            
            GET_IVA_ARTICOLO fnNotNullN(rs!IDArticolo)
            ImballoARendere = GET_TIPO_IMBALLO_A_RENDERE(fnNotNullN(rs!IDArticolo))
            
            ObjDoc.Tables(NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)).SetActiveRetail I
                        
            ObjDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_codice", fnNotNull(rs!CodiceArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_descrizione", fnNotNull(rs!Articolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_quantita_totale", fnNotNullN(rs!Quantita), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_prezzo_unitario_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_lordo_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_pre_uni_net_sco_net_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_pre_uni_net_sco_lor_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Art_Importo_totale_neutro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_prezzo_unitario_neutro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            ObjDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
        
            ObjDoc.Field "Art_Importo_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
        
            If (ImballoARendere = 1) And (FLAG_IVA_IMBALLO_A_RENDERE = 1) Then
                ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            Else
                If fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
                    ObjDoc.Field "Link_art_IVA", fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", GET_ALIQUOTA_IVA_ARTICOLO(fnNotNullN(ObjDoc.Field("Link_Nom_IVA", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                Else
                    ObjDoc.Field "Link_art_IVA", Link_IVAArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_aliquota_IVA", AliquotaIvaArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                End If
            End If
                
                
            ObjDoc.Field "Art_tara", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "Art_importo_totale_netto_IVA", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                
            '''''''''''''''''''''''''''''''''''''''AGENTE
            If ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) > 0 Then
                LINK_REGOLA_PROVV = GET_LINK_REGOLA_PROVV_AGE(ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), TheApp.IDFirm)
                
                If LINK_REGOLA_PROVV > 0 Then
                    ObjDoc.Field "Link_Art_agente", ObjDoc.Field("Link_Doc_Agente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_nome", ObjDoc.Field("Doc_age_nome", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_codice", ObjDoc.Field("Doc_age_codice", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_ragione_sociale", ObjDoc.Field("Doc_age_ragione_sociale", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    
                    ObjDoc.Field "Link_Art_age_regola_provv", LINK_REGOLA_PROVV, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    ObjDoc.Field "Art_age_regola_provv", GET_DESCRIZIONE_REGOLA_PROVV(LINK_REGOLA_PROVV), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    If fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))) > 0 Then
                        ObjDoc.Field "Link_Art_age_tipo_ordine", GET_LINK_TIPO_ORDINE(LINK_REGOLA_PROVV, fnNotNullN(rs!Sconto1), fnNotNullN(rs!Sconto2)), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
                    End If
                End If
            End If
            
            If FLAG_INTRASTAT_DOC = 1 Then
                GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDArticolo), TheApp.IDFirm, TheApp.Branch, ObjDoc.Field("Art_quantita_totale", , NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo))
            End If
                
            ObjDoc.Field "Link_Art_unita_di_misura", 8, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POLinkRiga", Link_Riga, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POTipoRiga", 2, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PORigaCompleta", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDImballo", fnNotNullN(rs!IDArticolo), NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ProgressivoArticolo = ProgressivoArticolo + 1
            ObjDoc.Field "ID_Art_dettaglio_prog", ProgressivoArticolo, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            ObjDoc.Field "RV_POIDCalibro", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoCategoria", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDTipoLavorazione", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PODataConferimento", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDConferimentoRighe", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POIDSocio", 0, NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_PONomeSocio", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POLottoCampagna", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            ObjDoc.Field "RV_POCodiceLotto", "", NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
            
            Link_Riga = Link_Riga + 1
            I = I + 1
        End If
    rs.MoveNext
    Wend
    
rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_LINK_TIPO_ORDINE(IDRegolaProvv As Long, Sconto1 As Double, Sconto2 As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



sSQL = "SELECT * FROM RV_POIEConfigRegolaProvv "
sSQL = sSQL & "WHERE IDRegolaProvv=" & IDRegolaProvv
sSQL = sSQL & " AND Sconto1=" & fnNormNumber(Sconto1)
sSQL = sSQL & " AND Sconto2=" & fnNormNumber(Sconto2)

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_TIPO_ORDINE = 0
Else
    GET_LINK_TIPO_ORDINE = fnNotNullN(rs!Valore1)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoListinoImballo As Double


GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO = 0

sSQL = "SELECT Art_importo_net_sconto_net_IVA, RV_POIDImballo, RV_POImportoUnitarioImballo,RV_POImportoMerceNetta,Art_quantita_totale, "
sSQL = sSQL & "RV_POImportoImballoInArticolo, Art_numero_colli "
sSQL = sSQL & "FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
sSQL = sSQL & " WHERE (IDOggetto = " & ObjDoc.IDOggetto & ") "
sSQL = sSQL & " AND (RV_POTipoRiga = 1)"


Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO = GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO + (fnNotNullN(rs!Art_importo_net_sconto_net_IVA))
    If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 0 Then
        GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO = GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO + (fnNotNullN(rs!RV_POImportoUnitarioImballo) * fnNotNullN(rs!Art_numero_colli))
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_SIGLA_UM(IDUM As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoListinoImballo As Double

GET_SIGLA_UM = ""


sSQL = "SELECT * FROM UnitaDiMisura "
sSQL = sSQL & " WHERE IDUnitaDiMisura = " & IDUM

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SIGLA_UM = fnNotNull(rs!DescrizioneFattura)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub GET_LINK_PORTO_PER_COMM_TRASP()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDPortoNoCalcoloTrasportoComm FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_PORTO_NO_TRASP = fnNotNullN(rs!IDPortoNoCalcoloTrasportoComm)
Else
    LINK_PORTO_NO_TRASP = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_TIPO_CALCOLO_COMM_TRASP()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AbilitaCalcoloPesoTrasportoTipoPedana, "
sSQL = sSQL & "AttivaCommissioniDaOrdine, RicalcolaCommTipoPedDaOrdInEvasione,"
sSQL = sSQL & "VisElencoRigheOrdineSeNonTroviAssociazione, NonRipAgenteDaEvOrdNoPres "
sSQL = sSQL & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    TIPO_CALCOLO_TRASPORTO = fnNotNullN(rs!AbilitaCalcoloPesoTrasportoTipoPedana)
    ATTIVA_COMMISSIONI_DA_ORDINE = fnNotNullN(rs!AttivaCommissioniDaOrdine)
    RIC_COMM_TIPO_PED_DA_ORD = fnNotNullN(rs!RicalcolaCommTipoPedDaOrdInEvasione)
    VIS_ELECO_RIGHE_ORD = fnNotNullN(rs!VisElencoRigheOrdineSeNonTroviAssociazione)
    NO_RIP_AGENTE_IN_DOC_EVASIONE = fnNotNullN(rs!NonRipAgenteDaEvOrdNoPres)
Else
    TIPO_CALCOLO_TRASPORTO = 0
    ATTIVA_COMMISSIONI_DA_ORDINE = 0
    RIC_COMM_TIPO_PED_DA_ORD = 0
    VIS_ELECO_RIGHE_ORD = 0
    NO_RIP_AGENTE_IN_DOC_EVASIONE = 0
End If

rs.CloseResultset
Set rs = Nothing

End Sub
Private Sub AGGIORNA_ORDINE_PADRE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim link_oggetto_padre As Long
Dim Completato As Long

link_oggetto_padre = 0

sSQL = "SELECT IDOggetto, RV_POIDOrdinePadre "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    link_oggetto_padre = fnNotNullN(rs!RV_POIDOrdinePadre)
End If
rs.CloseResultset
Set rs = Nothing

If link_oggetto_padre = 0 Then Exit Sub

Completato = CONTROLLO_ORDINE_COMPLETATO(link_oggetto_padre)

If Completato = 1 Then
    sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
    sSQL = sSQL & "RV_POOrdineCompletato=" & Completato
    sSQL = sSQL & "WHERE IDOggetto=" & link_oggetto_padre
    
    CnDMT.Execute sSQL
End If

End Sub
Private Function CONTROLLO_ORDINE_COMPLETATO(IDOggettoOrdinePadre As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
CONTROLLO_ORDINE_COMPLETATO = 0

sSQL = "SELECT IDOggetto, RV_POIDOrdinePadre "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE RV_POIDOrdinePadre=" & IDOggettoOrdinePadre
sSQL = sSQL & " AND Doc_ordine_chiuso=0 "

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    CONTROLLO_ORDINE_COMPLETATO = 1
Else
    CONTROLLO_ORDINE_COMPLETATO = 0
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_STAMPA_DOC_ATT()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT StampaDocEvasioneOrdAttivo, StampaDocEvasioneOrdNonAttivo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    STAMPA_DOCUMENTO_ATTIVO = fnNotNullN(rs!StampaDocEvasioneOrdAttivo)
    STAMPA_DOCUMENTO_NON_ATTIVO = fnNotNullN(rs!StampaDocEvasioneOrdNonAttivo)
Else
    STAMPA_DOCUMENTO_ATTIVO = 0
    STAMPA_DOCUMENTO_NON_ATTIVO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_TOTALE_COSTO_KIT(IDLavorazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(CostoTotaleRiga) as TotaleCosto "
sSQL = sSQL & "FROM RV_POAssegnazioneMerceImbPrim "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_COSTO_KIT = 0
Else
    GET_TOTALE_COSTO_KIT = fnNotNullN(rs!TotaleCosto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_PARAMETRI_LIQ(IDFiliale As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    AGG_COSTO_KIT_PRZ_LIQ = 0
    AGG_COSTO_CONFEZ_PRZ_LIQ = 0
    COMMISSIONI_DA_PRZ_VENDITA = 0
Else
    AGG_COSTO_KIT_PRZ_LIQ = fnNotNullN(rs!AggiungiCostoKitInPrezzoLiq)
    AGG_COSTO_CONFEZ_PRZ_LIQ = fnNotNullN(rs!AggiungiCostoConfezInPrezzoLiq)
    COMMISSIONI_DA_PRZ_VENDITA = fnNotNullN(rs!CommissioniDaPrezzoDiVendita)
End If


rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroICE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT UsaProtocolloICEPeriodo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    USA_PROT_ICE_PERIODO = fnNotNullN(rs!UsaProtocolloICEPeriodo)
Else
    USA_PROT_ICE_PERIODO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_IMPORTO_ART_TRATT_CLI(IDAnagrafica As Long, IDArticolo As Long, QtaLiquidazione As Double) As Double
On Error GoTo ERR_GET_IMPORTO_ART_TRATT_CLI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_IMPORTO_ART_TRATT_CLI = 0

sSQL = "SELECT * FROM RV_POConfigurazioneClienteArtTratt "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = CnDMT.OpenResultset(sSQL)
    GET_IMPORTO_ART_TRATT_CLI = -(fnNotNullN(rs!ValoreTrattenuta))
If Not rs.EOF Then
    
End If


rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_IMPORTO_ART_TRATT_CLI:
    

End Function
Private Sub ConsolidaDettaglioFatturaElettronica()
On Error GoTo ERR_ConsolidaDettaglioFatturaElettronica
    Dim oRs As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFilter As String
    
    'DATI
    'leggo i dati legati al documento
    With ObjDoc.ElectronicInvoiceAdditionalData.AdditionalData
        'conservo il filtro per rimetterlo dopo
        If Len(.Filter) > 0 Then sFilter = .Filter
                            
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            While Not .EOF
                If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                    .Fields("Temporaneo").Value = False
                End If
                .MoveNext
            Wend
        End If
        .Filter = sFilter
    End With
    'DATI
    'leggo i dati legati al documento
    With ObjDoc.ElectronicInvoiceAdditionalData.AdditionalCodes
        'conservo il filtro per rimetterlo dopo
        If Len(.Filter) > 0 Then sFilter = .Filter
                            
        If Not (.EOF And .BOF) Then
            .MoveFirst
            
            While Not .EOF
                If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                    .Fields("Temporaneo").Value = False
                End If
                .MoveNext
            Wend
        End If
        .Filter = sFilter
    End With
    
    ObjDoc.ElectronicInvoiceAdditionalData.Changed = True

Exit Sub
ERR_ConsolidaDettaglioFatturaElettronica:
    MsgBox Err.Description, vbCritical, "ConsolidaDettaglioFatturaElettronica"
End Sub
Private Sub sbLoadElectronicInvoiceData4Article(ByVal lID_Art_dettaglio_prog As Long, ByVal lIDArticle As Long)
On Error GoTo ERR_sbLoadElectronicInvoiceData4Article
    Dim oRs As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFilter As String

    'DATI
    'leggo i dati legati all'articolo con IDArticolo richiesto
    Set oRs = ObjDoc.ElectronicInvoiceAdditionalData.GetDataFromArticle(lIDArticle)
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then
            If oRs.RecordCount > 0 Then
                With ObjDoc.ElectronicInvoiceAdditionalData.AdditionalData
                    'conservo il filtro per rimetterlo dopo
                    If Len(.Filter) > 0 Then sFilter = .Filter
                    
                    'metto gli elementi precedenti in naftalina ;)
                    .Filter = "ID_Art_dettaglio_prog = " & lID_Art_dettaglio_prog
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        
                        While Not .EOF
                            If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                                .Fields("ID_Art_dettaglio_prog").Value = -1 * fnNotNullL(.Fields("ID_Art_dettaglio_prog").Value)
                            End If
                            .MoveNext
                        Wend
                    End If
                    .Filter = sFilter
                    
                    'riverso i dati aggiuntivi dell'articolo sul dettaglio
                    oRs.MoveFirst
                    
                    While Not oRs.EOF
                        .AddNew
                        
                        For Each oField In oRs.Fields
                            If oField.Name <> "IDDatoFatturaPAPerArticolo" Then
                                .Fields(oField.Name).Value = oField.Value
                            End If
                        Next
                        Set oField = Nothing
                        
                        .Fields("IDOggetto").Value = ObjDoc.IDOggetto
                        .Fields("IDTipoOggetto").Value = ObjDoc.IDTipoOggetto
                        .Fields("ID_Art_dettaglio_prog").Value = lID_Art_dettaglio_prog
                        .Fields("Eliminato").Value = False
                        'Impostare "Temporaneo" a False se i codici vengono immediatamente legati al dettaglio,
                        'a True se questi codici rimangono sospesi in attesa di un ulteriore conferma
                        '(Salva di dettaglio, ad es., i cui Temporaneo verrà finalmente posto a False)
                        .Fields("Temporaneo").Value = False
                        
                        oRs.MoveNext
                    Wend
                    ObjDoc.ElectronicInvoiceAdditionalData.Changed = True
                End With
            End If
            oRs.Close
        End If
    End If
    Set oRs = Nothing
    
    'CODICI
    'leggo i codici legati all'articolo con IDArticolo richiesto
    Set oRs = ObjDoc.ElectronicInvoiceAdditionalData.GetCodesFromArticle(lIDArticle)
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then
            If oRs.RecordCount > 0 Then
                With ObjDoc.ElectronicInvoiceAdditionalData.AdditionalCodes
                    'conservo il filtro per rimetterlo dopo
                    If Len(.Filter) > 0 Then sFilter = .Filter
                    'metto gli elementi precedenti in naftalina ;)
                    .Filter = "ID_Art_dettaglio_prog = " & lID_Art_dettaglio_prog
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                        
                        While Not .EOF
                            If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                                .Fields("ID_Art_dettaglio_prog").Value = -1 * fnNotNullL(.Fields("ID_Art_dettaglio_prog").Value)
                            End If
                            .MoveNext
                        Wend
                    End If
                    .Filter = sFilter
                    
                    'riverso i codici aggiuntivi dell'articolo sul dettaglio
                    oRs.MoveFirst
                    
                    While Not oRs.EOF
                        .AddNew
                        
                        For Each oField In oRs.Fields
                            If oField.Name <> "IDCodiceFatturaPAPerArticolo" Then
                                .Fields(oField.Name).Value = oField.Value
                            End If
                        Next
                        Set oField = Nothing
                        
                        .Fields("IDOggetto").Value = ObjDoc.IDOggetto
                        .Fields("IDTipoOggetto").Value = ObjDoc.IDTipoOggetto
                        .Fields("ID_Art_dettaglio_prog").Value = lID_Art_dettaglio_prog
                        .Fields("Eliminato").Value = False
                        'Impostare "Temporaneo" a False se i codici vengono immediatamente legati al dettaglio,
                        'a True se questi codici rimangono sospesi in attesa di un ulteriore conferma
                        '(Salva di dettaglio, ad es., i cui Temporaneo verrà finalmente posto a False)
                        .Fields("Temporaneo").Value = False
                        
                        oRs.MoveNext
                    Wend
                    ObjDoc.ElectronicInvoiceAdditionalData.Changed = True
                End With
            End If
            oRs.Close
        End If
    End If
    Set oRs = Nothing
Exit Sub
ERR_sbLoadElectronicInvoiceData4Article:
    MsgBox Err.Description, vbCritical, "sbLoadElectronicInvoiceData4Article"
End Sub
Private Function GET_RIF_PA_ARTICOLO(IDArticolo As Long, IDCliente As Long, IDDestinazione As Long) As String
On Error GoTo ERR_GET_RIF_PA_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_RIF_PA_ARTICOLO = ""

'Articolo - Cliente
sSQL = "SELECT RiferimentoPACliente FROM ClientePerArticolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAnagrafica=" & IDCliente

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPACliente)
End If
rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_RIF_PA_ARTICOLO)) > 0 Then Exit Function

'Destinazione
sSQL = "SELECT RiferimentoPAArticolo FROM SitoPerAnagrafica "
sSQL = sSQL & " WHERE IDSitoPerAnagrafica=" & IDDestinazione

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPAArticolo)
End If
rs.CloseResultset
Set rs = Nothing


If Len(Trim(GET_RIF_PA_ARTICOLO)) > 0 Then Exit Function
'Cliente
sSQL = "SELECT RiferimentoPAArticolo FROM Cliente "
sSQL = sSQL & " WHERE IDAnagrafica=" & IDCliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPAArticolo)
End If
rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_RIF_PA_ARTICOLO)) > 0 Then Exit Function
'Articolo
sSQL = "SELECT RiferimentoPA FROM Articolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPA)
End If
rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_RIF_PA_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_RIF_PA_ARTICOLO"
End Function
Private Function GET_LINK_UM_ART(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_UM_ART = 0

sSQL = "SELECT IDUnitaDiMisuraVendita FROM Articolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_ART = fnNotNullN(rs!IDUnitaDiMisuraVendita)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = 0

sSQL = "SELECT RV_POIDUnitaDiMisuraLiq "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo
        
Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM Azienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_AZIENDA = 0
Else
    GET_LINK_ANAGRAFICA_AZIENDA = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub AGGIORNA_ORDINE_SELEZIONATO(IDOggetto As Long)
On Error GoTo ERR_AGGIORNA_ORDINE_SELEZIONATO
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto

Set rs = New ADODB.Recordset
rs.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rs.EOF Then
    rs!Doc_causale_documento = Me.txtCausaleDocOrdSel.Text
    rs!Link_Doc_spedizione = Me.cboTipoTrasportoOrdSel.CurrentID
    rs!Link_Vet_vettore = Me.cboVettoreOrdSel.CurrentID
    rs!RV_POIstruzioniMittente = Me.txtIstruzioniMittenteOrdSel.Text
    rs!RV_POTargaAutomezzo = Me.txtTargaAutomezzoOrdSel.Text
    rs!RV_POIDTrasportatoreSuccessivo = Me.cboVettoreSuccOrdSel.CurrentID
    rs!RV_POIDLuogoPresaMerce = Me.cboLuogoPresaMerceOrdSel.CurrentID
    If Me.txtDataArrivoLuogoOrdSel.Value > 0 Then
        rs!RV_PODataArrivoMerceLuogo = Me.txtDataArrivoLuogoOrdSel.Value
    Else
        rs!RV_PODataArrivoMerceLuogo = Null
    End If
    If Me.txtOraArrivoLuogoOrdSel.Value > 0 Then
        rs!RV_POOraArrivoMerceLuogo = Me.txtOraArrivoLuogoOrdSel.Text
    Else
        rs!RV_POOraArrivoMerceLuogo = Null
    End If
    rs!Link_Doc_aspetto_esteriore = Me.cboAspettoEsterioreOrdSel.CurrentID
    
    rs.Update
    
    MsgBox "Ordine aggiornato con successo!", vbInformation, "Aggiornamento dati"
Else
    MsgBox "ORDINE NON TROVATO!", vbCritical, "Aggiornamento dati"
End If

rs.Close
Set rs = Nothing

Exit Sub
ERR_AGGIORNA_ORDINE_SELEZIONATO:
    MsgBox Err.Description, vbCritical, "AGGIORNA_ORDINE_SELEZIONATO"
End Sub
Private Sub IMPOSTA_COMMISSIONI_PER_CLIENTE_DA_ORD(IDCliente As Long, IDOggettoOrdine As Long, IDOggettoDocumento As Long, IDSitoPerAnagrafica As Long)
On Error GoTo ERR_IMPOSTA_COMMISSIONI_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim rsCli As ADODB.Recordset
Dim SpesaTrasporto As Double
Dim Totale_merce_lavorato As Double
Dim Totale_merce_lavorato_lordo As Double
Dim Totale_documento_netto_iva As Double
Dim Totale_documento_lordo_iva As Double
Dim Totale_Importo As Double
Dim IDOrdinePadre As Double

sSQL = "SELECT RV_POIDOrdinePadre FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    IDOrdinePadre = fnNotNullN(rs!RV_POIDOrdinePadre)
End If

rs.CloseResultset
Set rs = Nothing

Totale_merce_lavorato = GET_TOTALE_MERCE_DOCUMENTO(IDOggettoDocumento)
Totale_merce_lavorato_lordo = GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO()
Totale_documento_netto_iva = fnNotNullN(ObjDoc.Field("Tot_imponibile_corr", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))
Totale_documento_lordo_iva = fnNotNullN(ObjDoc.Field("Tot_documento_corr", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))


sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
sSQL = sSQL & " WHERE IDRV_POCommissioniPerDoc=0"
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

''''''''COMMISSIONI A PERCENTUALE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIECommissioniPerDoc "
sSQL = sSQL & " WHERE IDOggetto=" & IDOrdinePadre
sSQL = sSQL & " AND CommissionePerPedana=0 "
sSQL = sSQL & " AND ((IDArticoloImballo=0) OR (IDArticoloImballo IS NULL)) "

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
        rsNew!IDOggetto = IDOggettoDocumento
        rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
        rsNew!Percentuale = fnNotNullN(rs!Percentuale)
        rsNew!PercentualeDaCommissione = fnNotNullN(rs!PercentualeDaCommissione)
        rsNew!Importo = 0
        rsNew!ImportoTotale = 0
        rsNew!Quantita = 1
        rsNew!APercentuale = 0
        
        Select Case fnNotNullN(rs!IDRV_POTipoValoreDocumento)
            
            Case 1
                Totale_Importo = Totale_merce_lavorato
            Case 2
                Totale_Importo = Totale_merce_lavorato_lordo
            Case 3
                Totale_Importo = Totale_documento_netto_iva
            Case 4
                Totale_Importo = Totale_documento_lordo_iva
            Case Else
                Totale_Importo = Totale_merce_lavorato
            
        End Select
        
        If Totale_Importo = 0 Then
            rsNew!ImportoRiga = 0
        Else
            If fnNotNullN(rs!IDRV_POTipoValoreDocumento) <= 1 Then
                If (rs!Percentuale > 0) Then
                    'rsNew!Percentuale = (rs!Importoriga / Totale_merce_lavorato) * 100
                    rsNew!ImportoRiga = (Totale_merce_lavorato / 100) * fnNotNullN(rs!Percentuale)
                Else
                    If fnNotNullN(rs!ImportoRiga) > 0 Then
                        rsNew!Percentuale = (rs!ImportoRiga / Totale_merce_lavorato) * 100
                        rsNew!ImportoRiga = fnNotNullN(rs!ImportoRiga)
                    End If
                End If
            Else
                If (rs!PercentualeDaCommissione > 0) Then
                    rsNew!ImportoRiga = (Totale_Importo / 100) * fnNotNullN(rs!PercentualeDaCommissione)
                    If (Totale_merce_lavorato > 0) Then
                        rsNew!Percentuale = (rsNew!ImportoRiga / Totale_merce_lavorato) * 100
                    Else
                        rsNew!Percentuale = 0
                    End If
                Else
                    rsNew!ImportoRiga = fnNotNullN(rs!ImportoRiga)
                    rsNew!Percentuale = (rsNew!ImportoRiga / Totale_Importo) * 100
                End If
            End If
        End If
    rsNew.Update
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Totale_merce_lavorato = 0 Then
    rsNew.Close
    Set rsNew = Nothing
Exit Sub
End If

'''''''''''''''''''''''''''''''COMMISSIONI PER IMBALLO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIECommissioniPerDoc "
sSQL = sSQL & " WHERE IDOggetto=" & IDOrdinePadre
sSQL = sSQL & " AND CommissionePerPedana=0 "
sSQL = sSQL & " AND IDArticoloImballo>0"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF

    sSQL = "SELECT Link_Art_articolo AS IDArticoloImballo, SUM(Art_quantita_totale) AS QuantitaPedana "
    sSQL = sSQL & "FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo)
    sSQL = sSQL & " WHERE IDOggetto =" & ObjDoc.IDOggetto
    sSQL = sSQL & " AND RV_POTipoRiga = 2 "
    sSQL = sSQL & " AND Link_Art_articolo=" & fnNotNullN(rs!IDArticoloImballo)
    sSQL = sSQL & "GROUP BY Link_Art_articolo"
    
    Set rsCli = New ADODB.Recordset
    
    rsCli.Open sSQL, CnDMT.InternalConnection
    
    While Not rsCli.EOF
        rsNew.AddNew
            rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
            rsNew!IDOggetto = IDOggettoDocumento
            rsNew!IDRV_POTipoCommissione = fnNotNullN(rsCli!IDRV_POTipoCommissione)
            SpesaTrasporto = fnNotNullN(rs!ImportoRiga) * fnNotNullN(rsCli!QuantitaPedana)
            rsNew!Percentuale = (SpesaTrasporto / Totale_merce_lavorato) * 100
            rsNew!Importo = 0
            rsNew!ImportoRiga = SpesaTrasporto
            rsNew!IDArticoloImballo = fnNotNullN(rsCli!IDArticoloImballo)
        rsNew.Update
    rsCli.MoveNext
    Wend
    
    rsCli.Close
    Set rsCli = Nothing
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''COMMISSIONI PER TIPO PEDANA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIECommissioniPerDoc "
sSQL = sSQL & " WHERE IDOggetto=" & IDOrdinePadre
sSQL = sSQL & " AND CommissionePerPedana=1 "
sSQL = sSQL & " AND IDRV_POTipoPedana>0"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF

    sSQL = "SELECT IDOggetto, RV_POIDTipoPedana, RV_POTipoPedana.IDArticoloImballo,"
    sSQL = sSQL & " (SELECT COUNT(*) AS QuantitaTipoPedana "
    sSQL = sSQL & " FROM (SELECT RV_POIDPedana "
    sSQL = sSQL & " FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " AS Tabella INNER JOIN "
    sSQL = sSQL & " RV_POTipoPedana ON Tabella.RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
    sSQL = sSQL & " Where (IDOggetto = " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".IDOggetto) And (RV_POTipoRiga = 1) And (RV_POIDTipoPedana = " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana) "
    sSQL = sSQL & " GROUP BY Tabella.RV_POIDPedana, RV_POTipoPedana.IDArticoloImballo) AS X) AS QuantitaPedana "
    sSQL = sSQL & " FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " INNER JOIN "
    sSQL = sSQL & " RV_POTipoPedana ON " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
    sSQL = sSQL & " WHERE IDOggetto = " & ObjDoc.IDOggetto
    sSQL = sSQL & " AND  RV_POTipoRiga = 1"
    sSQL = sSQL & " AND RV_POIDTipoPedana=" & fnNotNullN(rs!IDRV_POTipoPedana)
    sSQL = sSQL & " AND RV_POTipoPedana.IDArticoloImballo=" & fnNotNullN(rs!IDArticoloImballo)
    sSQL = sSQL & " GROUP BY IDOggetto, RV_POIDTipoPedana, IDArticoloImballo"
    
    Set rsCli = New ADODB.Recordset
    
    rsCli.Open sSQL, CnDMT.InternalConnection
    
    While Not rsCli.EOF
        rsNew.AddNew
            rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
            rsNew!IDOggetto = IDOggettoDocumento
            rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
            rsNew!Percentuale = 0
            rsNew!Importo = 0
            rsNew!ImportoRiga = fnNotNullN(rs!ImportoRiga)
            rsNew!Quantita = fnNotNullN(rsCli!QuantitaPedana)
            rsNew!ImportoTotale = 0
            rsNew!IDRV_POTipoPedana = fnNotNullN(rsCli!RV_POIDTipoPedana)
            rsNew!APercentuale = 1
            rsNew!IDArticoloImballo = fnNotNullN(rsCli!IDArticoloImballo)
        rsNew.Update
    rsCli.MoveNext
    Wend
    
    rsCli.Close
    Set rsCli = Nothing
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''COMMISSIONI PER TIPO PEDANA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIECommissioniPerDoc "
sSQL = sSQL & " WHERE IDOggetto=" & IDOrdinePadre
sSQL = sSQL & " AND CommissionePerPedana=1 "
sSQL = sSQL & " AND IDRV_POTipoPedana=0"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF

'    sSQL = "SELECT IDOggetto, RV_POIDTipoPedana, RV_POTipoPedana.IDArticoloImballo,"
'    sSQL = sSQL & " (SELECT COUNT(*) AS QuantitaTipoPedana "
'    sSQL = sSQL & " FROM (SELECT RV_POIDPedana "
'    sSQL = sSQL & " FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " AS Tabella INNER JOIN "
'    sSQL = sSQL & " RV_POTipoPedana ON Tabella.RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
'    sSQL = sSQL & " Where (IDOggetto = " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".IDOggetto) And (RV_POTipoRiga = 1) And (RV_POIDTipoPedana = " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana) "
'    sSQL = sSQL & " GROUP BY Tabella.RV_POIDPedana, RV_POTipoPedana.IDArticoloImballo) AS X) AS QuantitaPedana "
'    sSQL = sSQL & " FROM " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & " INNER JOIN "
'    sSQL = sSQL & " RV_POTipoPedana ON " & NomeTabellaDettaglio & fnGetHex(ObjDoc.IDCorpo) & ".RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
'    sSQL = sSQL & " WHERE IDOggetto = " & ObjDoc.IDOggetto
'    sSQL = sSQL & " AND  RV_POTipoRiga = 1"
'    sSQL = sSQL & " AND RV_POIDTipoPedana=" & fnNotNullN(rs!IDRV_POTipoPedana)
'    sSQL = sSQL & " AND RV_POTipoPedana.IDArticoloImballo=" & fnNotNullN(rs!IDArticoloImballo)
'    sSQL = sSQL & " GROUP BY IDOggetto, RV_POIDTipoPedana, IDArticoloImballo"
'
'    Set rsCli = New ADODB.Recordset
'
'    rsCli.Open sSQL, CnDMT.InternalConnection
'
'    While Not rsCli.EOF
        rsNew.AddNew
            rsNew!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
            rsNew!IDOggetto = IDOggettoDocumento
            rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
            rsNew!Percentuale = 0
            rsNew!Importo = 0
            rsNew!ImportoRiga = fnNotNullN(rs!ImportoRiga)
            rsNew!Quantita = 1
            rsNew!ImportoTotale = 0
            rsNew!IDRV_POTipoPedana = fnNotNullN(rs!IDRV_POTipoPedana)
            rsNew!APercentuale = 1
            rsNew!IDArticoloImballo = fnNotNullN(rs!IDArticoloImballo)
        rsNew.Update
'    rsCli.MoveNext
'    Wend
'
'    rsCli.Close
'    Set rsCli = Nothing
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_IMPOSTA_COMMISSIONI_PER_CLIENTE:
    MsgBox Err.Description, vbCritical, "IMPOSTA_COMMISSIONI_PER_CLIENTE_DA_ORD"

End Sub
Public Function GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED(IDTipoCommissione As Long, IDTipoPedana As Long, IDArticoloImballo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione
sSQL = sSQL & " AND IDRV_POTipoPedana=" & IDTipoPedana
sSQL = sSQL & " AND IDOggetto=" & ObjDoc.IDOggetto
sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED = False
Else
    GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED = True
End If
    
rs.CloseResultset
Set rs = Nothing

End Function
Public Function GET_TIPO_RICALCOLO_COMM(IDTipoCommissione As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTipoCommissione "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione


Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_TIPO_RICALCOLO_COMM = fnNotNullN(rs!IDRV_POTipoRicalcoloComm)
End If
    
rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_RIGHE_DOCUMENTO_PER_PEDANA_TOTALE(tabella As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsPed As DmtOleDbLib.adoResultset
Dim rsRiga As ADODB.Recordset
Dim Moltiplicatore As Double
Dim Prezzo As Double
Dim ImportoCommPerPedana As Double
Dim rsNew As ADODB.Recordset
Dim ImportoCommArticoloPerPedana As Double


'SETTAGGIO DEL RECORDSET PER LE RIGHE COMMISSIONI'''''''''''''''''''''''''''''
Set rsNew = New ADODB.Recordset

sSQL = "SELECT * FROM RV_POCommissioniPerDocRighe "
sSQL = sSQL & " WHERE IDOggetto=" & ObjDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
sSQL = sSQL & " WHERE IDOggetto=" & ObjDoc.IDOggetto
sSQL = sSQL & " AND IDRV_POTipoPedana=0 "
sSQL = sSQL & " AND APercentuale=1 "

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "SELECT IDOggetto, SUM(Art_peso - Art_tara) AS TotalePesoNetto, SUM(Art_volume*Art_numero_colli) AS TotaleVolume "
    sSQL = sSQL & " FROM " & tabella
    sSQL = sSQL & " WHERE IDOggetto = " & ObjDoc.IDOggetto
    'sSQL = sSQL & " AND RV_POIDTipoPedana = " & fnNotNullN(rs!IDRV_POTipoPedana)
    sSQL = sSQL & " AND RV_POTipoRiga = 1 "
    sSQL = sSQL & " GROUP BY IDOggetto"

    Set rsPed = CnDMT.OpenResultset(sSQL)
    
    While Not rsPed.EOF
        If TIPO_CALCOLO_TRASPORTO = 1 Then
            If fnNotNullN(rsPed!TotalePesoNetto) > 0 Then
                
                sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, RV_POPrezzoMedioInLiq, RV_POVariazionePrezzoImballo, RV_POImportoMerceNetta, RV_POImportoLiq, "
                sSQL = sSQL & "RV_POPrezzoUnitarioOrigine, RV_POIDIvaImballo, RV_POIDTipoVariazione, Art_pre_uni_net_sco_net_IVA ,"
                sSQL = sSQL & "RV_POImportoDaLiq, RV_POLinkRiga, RV_POImportoImballoInArticolo, link_art_articolo, "
                sSQL = sSQL & "Art_numero_colli, Art_quantita_Totale, Link_art_articolo, (Art_Peso - Art_Tara) AS PesoNettoRiga, Art_Volume,  (Art_volume*Art_numero_colli) VolumeRigaVendita, "
                sSQL = sSQL & "RV_POVariazionePrezzoManuale, RV_POImportoRigaCommissioni, RV_PODataLavorazione, RV_PODataConferimento, RV_POImportoLiqDoc, RV_POQuantitaLiq "
                sSQL = sSQL & " FROM " & tabella
                sSQL = sSQL & " WHERE RV_POTipoRiga=1 "
                sSQL = sSQL & " AND IDOggetto=" & ObjDoc.IDOggetto
                sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto
                'sSQL = sSQL & " AND RV_POIDPedana=" & fnNotNullN(rsPed!RV_POIDPedana)
                
                Set rsRiga = New ADODB.Recordset
                rsRiga.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
                
                While Not rsRiga.EOF
                
                    If fnNotNullN(rsRiga!PesoNettoRiga) <> 0 Then
                        
                        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rsRiga!Link_Art_articolo))
                        
                        ImportoCommArticoloPerPedana = (fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotalePesoNetto)) * fnNotNullN(rsRiga!PesoNettoRiga)

                        ImportoCommArticoloPerPedana = ImportoCommArticoloPerPedana / fnNotNullN(rsRiga!RV_POQuantitaLiq) / Moltiplicatore
                        
                        rsRiga!RV_POImportoLiq = rsRiga!RV_POImportoLiq - Abs(ImportoCommArticoloPerPedana)
                        rsRiga!RV_POImportoRigaCommissioni = rsRiga!RV_POImportoRigaCommissioni + Abs(ImportoCommArticoloPerPedana)
                        rsRiga!RV_POImportoLiqDoc = rsRiga!RV_POImportoLiq
                        rsRiga.Update
                        
                        rsNew.AddNew
                            rsNew!IDRV_POCommissioniPerDocRighe = fnGetNewKey("RV_POCommissioniPerDocRighe", "IDRV_POCommissioniPerDocRighe")
                            rsNew!IDRV_POCommissioniPerDoc = fnNotNullN(rs!IDRV_POCommissioniPerDoc)
                            rsNew!IDOggetto = ObjDoc.IDOggetto
                            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
                            rsNew!IDValoriOggettoDettaglio = rsRiga!IDValoriOggettoDettaglio
                            rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
                            rsNew!Percentuale = 0
                            rsNew!DataVendita = ObjDoc.DataEmissione
                            rsNew!DataConferimento = rsRiga!RV_PODataConferimento
                            rsNew!DataLavorazione = rsRiga!RV_PODataLavorazione
                            rsNew!Importo = ImportoCommArticoloPerPedana '(fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotalePesoNetto)) * fnNotNullN(rsRiga!PesoNettoRiga)
                            rsNew!IDArticolo = fnNotNullN(rsRiga!Link_Art_articolo)
                            rsNew!Quantita = fnNotNullN(rsRiga!Art_quantita_totale)
                            
                        rsNew.Update
                        
                    End If
                    
                rsRiga.MoveNext
                Wend
                
                rsRiga.Close
                Set rsRiga = Nothing
            End If
        Else
            
            If fnNotNullN(rsPed!TotaleVolume) > 0 Then
                
                sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, RV_POPrezzoMedioInLiq, RV_POVariazionePrezzoImballo, RV_POImportoMerceNetta, RV_POImportoLiq, "
                sSQL = sSQL & "RV_POPrezzoUnitarioOrigine, RV_POIDIvaImballo, RV_POIDTipoVariazione, Art_pre_uni_net_sco_net_IVA ,"
                sSQL = sSQL & "RV_POImportoDaLiq, RV_POLinkRiga, RV_POImportoImballoInArticolo, link_art_articolo, "
                sSQL = sSQL & "Art_numero_colli, Art_quantita_Totale, Link_art_articolo, (Art_Peso - Art_Tara) AS PesoNettoRiga, Art_volume, (Art_volume*Art_numero_colli) VolumeRigaVendita, "
                sSQL = sSQL & "RV_POVariazionePrezzoManuale, RV_POImportoRigaCommissioni, RV_PODataLavorazione, RV_PODataConferimento, RV_POImportoLiqDoc, RV_POQuantitaLiq "
                sSQL = sSQL & " FROM " & tabella
                sSQL = sSQL & " WHERE RV_POTipoRiga=1 "
                sSQL = sSQL & " AND IDOggetto=" & ObjDoc.IDOggetto
                sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto
                'sSQL = sSQL & " AND RV_POIDPedana=" & fnNotNullN(rsPed!RV_POIDPedana)
                
                Set rsRiga = New ADODB.Recordset
                rsRiga.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
                
                While Not rsRiga.EOF
                
                    If fnNotNullN(rsRiga!VolumeRigaVendita) <> 0 Then
                        
                        Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rsRiga!Link_Art_articolo))
                        
                        ImportoCommArticoloPerPedana = (fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotaleVolume)) * fnNotNullN(rsRiga!VolumeRigaVendita)
                        
                        ImportoCommArticoloPerPedana = ImportoCommArticoloPerPedana / fnNotNullN(rsRiga!RV_POQuantitaLiq) / Moltiplicatore
                        
                        rsRiga!RV_POImportoLiq = rsRiga!RV_POImportoLiq - Abs(ImportoCommArticoloPerPedana)
                        rsRiga!RV_POImportoRigaCommissioni = rsRiga!RV_POImportoRigaCommissioni + Abs(ImportoCommArticoloPerPedana)
                        rsRiga!RV_POImportoLiqDoc = rsRiga!RV_POImportoLiq
                        
                        rsRiga.Update
                        
                        rsNew.AddNew
                            rsNew!IDRV_POCommissioniPerDocRighe = fnGetNewKey("RV_POCommissioniPerDocRighe", "IDRV_POCommissioniPerDocRighe")
                            rsNew!IDRV_POCommissioniPerDoc = fnNotNullN(rs!IDRV_POCommissioniPerDoc)
                            rsNew!IDOggetto = ObjDoc.IDOggetto
                            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
                            rsNew!IDValoriOggettoDettaglio = rsRiga!IDValoriOggettoDettaglio
                            rsNew!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
                            rsNew!Percentuale = 0
                            rsNew!DataVendita = ObjDoc.DataEmissione
                            rsNew!DataConferimento = rsRiga!RV_PODataConferimento
                            rsNew!DataLavorazione = rsRiga!RV_PODataLavorazione
                            rsNew!Importo = ImportoCommArticoloPerPedana '(fnNotNullN(rs!ImportoRiga) / fnNotNullN(rsPed!TotalePesoNetto)) * fnNotNullN(rsRiga!PesoNettoRiga)
                            rsNew!IDArticolo = fnNotNullN(rsRiga!Link_Art_articolo)
                            rsNew!Quantita = fnNotNullN(rsRiga!Art_quantita_totale)
                            rsNew!VolumeUnitariaImballo = fnNotNullN(rsRiga!Art_volume)
                            rsNew!VolumeTotaleMerce = fnNotNullN(rsRiga!VolumeRigaVendita)
                            rsNew!VolumeTotalePedana = fnNotNullN(rsPed!TotaleVolume)
                        rsNew.Update
                        
                    End If
                    
                rsRiga.MoveNext
                Wend
                
                rsRiga.Close
                Set rsRiga = Nothing
                                
            End If
        End If
        
    rsPed.MoveNext
    Wend
    
    rsPed.CloseResultset
    Set rsPed = Nothing

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsNew.Close
Set rsNew = Nothing
End Sub

Private Function GET_NOTA_DOCUMENTO(IDTipoOggetto As Long, IDTipoNota As Long) As String
On Error GoTo ERR_GET_NOTA_DOCUMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NOTA_DOCUMENTO = ""

sSQL = "SELECT *FROM RV_POIENotePerDocumentoTipoOggetto "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND IDRV_POTipoNotePerDocumento=" & IDTipoNota
sSQL = sSQL & " AND Predefinito=1"

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_NOTA_DOCUMENTO = fnNotNull(rs!Annotazione)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_NOTA_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "GET_NOTA_DOCUMENTO"
End Function
Private Sub cmdNota2_Click()

    If Me.CboTipoDocumento.CurrentID = 0 Then Exit Sub
    

    LINK_TIPO_NOTA_SEL = 2
    frmElencoNote.Show vbModal
    
    If CONFERMA_RIGA_NOTA = 1 Then
        Me.txtAnnotazione02.Text = RETURN_RIGA_NOTA
    End If
    
    
End Sub

Private Sub cmdNota3_Click()
    If Me.CboTipoDocumento.CurrentID = 0 Then Exit Sub
    LINK_TIPO_NOTA_SEL = 3
    frmElencoNote.Show vbModal

    If CONFERMA_RIGA_NOTA = 1 Then
        Me.txtAnnotazione03.Text = RETURN_RIGA_NOTA
    End If

End Sub

Private Sub cmdNote1_Click()
    If Me.CboTipoDocumento.CurrentID = 0 Then Exit Sub
    LINK_TIPO_NOTA_SEL = 1
    frmElencoNote.Show vbModal
    
    If CONFERMA_RIGA_NOTA = 1 Then
        Me.txtAnnotazione01.Text = RETURN_RIGA_NOTA
    End If
    
End Sub
Private Sub SCRIVI_ORD_CLI_RIF_XML(DataOrdine As String, NumeroOrdine As String)
On Error GoTo ERR_SCRIVI_ORD_CLI_RIF_XML
Dim sSQL As String
Dim rsNew As ADODB.Recordset

sSQL = "SELECT * FROM DatoFatturaPATestataDoc "
sSQL = sSQL & "WHERE IDDatoFatturaPATestataDoc=0"

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
rsNew.AddNew
    rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
    rsNew!IDBloccoXML = 1
    rsNew!IDOggetto = ObjDoc.IDOggetto
    rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
    '//rsNew!RiferimentoNumeroLinea = 1
    rsNew!IDDocumento = NumeroOrdine
    If Len(DataOrdine) > 0 Then
        rsNew!Data = DataOrdine
    End If
    rsNew!NumItem = NumeroOrdine
rsNew.Update

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_SCRIVI_ORD_CLI_RIF_XML:
    MsgBox Err.Description, vbCritical, "SCRIVI_ORD_CLI_RIF_XML"

End Sub
Private Function GET_SEZ_PER_CMR(IDSezionale As Long) As Long
On Error GoTo ERR_GET_SEZ_PER_CMR
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_SEZ_PER_CMR = 0

sSQL = "SELECT IDSezionale FROM RV_POSezionalePerCMR "
sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale
Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SEZ_PER_CMR = 1
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_SEZ_PER_CMR:
    MsgBox Err.Description, vbCritical, "GET_SEZ_PER_CMR"
End Function
Private Sub SCRIVI_CAUSALI_DOC(IDOggetto As Long)
On Error GoTo ERR_SCRIVI_CAUSALI_DOC
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rsDoc As DmtOleDbLib.adoResultset
Dim Tipo As Long
Dim Ordinamento As Long
Dim TestoVettoreSuccessivo As String
Dim TestoAgenziaTraporto As String

If ObjDoc.IDTipoOggetto = 8 Then Exit Sub

Tipo = 0


If (ObjDoc.Field("RV_PODocumentoCRM", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)) = 1) Then
    Tipo = 1
Else
    If fnNotNullN(ObjDoc.Field("RV_POIDAnagraficaDestinazione", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
        Tipo = 2
    End If
End If

sSQL = "SELECT * FROM DatoFatturaPATestataDoc "
sSQL = sSQL & "WHERE IDDatoFatturaPATestataDoc=0"

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic

Ordinamento = ADD_NOTE_TIPO_OGGETTO(ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, Tipo, rsNew)

If Rip_InXMLRifLetteraIntento = 1 Then
    ADD_NOTA_LETTERA_INTENTO ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, rsNew, fnNotNullN(ObjDoc.Field("Link_nom_lettera_intento", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))), Ordinamento
End If

If Rip_InXMLRifNoteIva = 1 Then
    ADD_NOTE_IVA ObjDoc.IDTipoOggetto, ObjDoc.IDOggetto, rsNew, Ordinamento
End If

'ANNOTAZIONE 1 DOCUMENTO
If Rip_InXMLRifNota01Doc = 1 Then
    If Len(Trim(fnNotNull(ObjDoc.Field("RV_POAnnotazioni1", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(ObjDoc.Field("RV_POAnnotazioni1", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), 1, 255)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ANNOTAZIONE 2 DOCUMENTO
If Rip_InXMLRifNota02Doc = 1 Then
    If Len(Trim(fnNotNull(ObjDoc.Field("RV_POAnnotazioni2", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(ObjDoc.Field("RV_POAnnotazioni2", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), 1, 255)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If
'ANNOTAZIONE 3 DOCUMENTO
If Rip_InXMLRifNota03Doc = 1 Then
    If Len(Trim(fnNotNull(ObjDoc.Field("RV_POAnnotazioni3", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(ObjDoc.Field("RV_POAnnotazioni3", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), 1, 255)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ANNOTAZIONE STANDARD DEL DOCUMENTO
If Rip_InXMLRifNotaDoc = 1 Then
    If Len(Trim(fnNotNull(ObjDoc.Field("Doc_annotazioni_variazio", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(ObjDoc.Field("Doc_annotazioni_variazio", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), 1, 255)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ISTRUZIONI DEL MITTENTE
If Rip_InXMLRifIstrMitt = 1 Then
    If Len(Trim(fnNotNull(ObjDoc.Field("RV_POIstruzioniMittente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(ObjDoc.Field("RV_POIstruzioniMittente", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), 1, 255)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If
'VETTORE SUCCESSIVO
If Rip_InXMLRifVettSucc = 1 Then
    If fnNotNullN(ObjDoc.Field("RV_POIDTrasportatoreSuccessivo", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
        TestoVettoreSuccessivo = GET_VETTORE_SUCCESSIVO(fnNotNullN(ObjDoc.Field("RV_POIDTrasportatoreSuccessivo", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))
        If Len(TestoVettoreSuccessivo) > 0 Then
            rsNew.AddNew
                rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                rsNew!IDBloccoXML = 8
                rsNew!IDOggetto = IDOggetto
                rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
                rsNew!Annotazioni = Mid(TestoVettoreSuccessivo, 1, 255)
                rsNew!Ordinamento = Ordinamento
                Ordinamento = Ordinamento + 1
            rsNew.Update
        End If
    End If
End If
'AGENZIA DI TRASPORTO
If Rip_InXMLRifAgenziaTrasp = 1 Then
    If fnNotNullN(ObjDoc.Field("RV_POIDAgenziaTrasportatore", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))) > 0 Then
        TestoAgenziaTraporto = GET_AGENZIA_TRASPORTO(fnNotNullN(ObjDoc.Field("RV_POIDAgenziaTrasportatore", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))
        If Len(TestoAgenziaTraporto) > 0 Then
            rsNew.AddNew
                rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                rsNew!IDBloccoXML = 8
                rsNew!IDOggetto = IDOggetto
                rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
                rsNew!Annotazioni = Mid(TestoAgenziaTraporto, 1, 255)
                rsNew!Ordinamento = Ordinamento
                Ordinamento = Ordinamento + 1
            rsNew.Update
        End If
    End If
End If
'TARGA AUTOMEZZO
If Rip_InXMLRifTargaAutoMezzo = 1 Then
    If Len(Trim(fnNotNull(ObjDoc.Field("RV_POTargaAutomezzo", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto))))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = ObjDoc.IDTipoOggetto
            rsNew!Annotazioni = "Targa automezzo: " & Mid(Trim(fnNotNull(ObjDoc.Field("RV_POTargaAutomezzo", , NOMETABELLAPIANA & fnGetHex(ObjDoc.IDTipoOggetto)))), 1, 255)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

rsNew.Close
Set rsNew = Nothing
Exit Sub
ERR_SCRIVI_CAUSALI_DOC:
    MsgBox Err.Description, vbCritical, "SCRIVI_CAUSALI_DOC"
End Sub
Private Function ADD_NOTE_TIPO_OGGETTO(IDTipoOggetto As Long, IDOggetto As Long, Tipo As Long, rsAdd As ADODB.Recordset) As Long
On Error GoTo ERR_ADD_NOTE_TIPO_OGGETTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Ordinamento As Long

sSQL = "SELECT * FROM RV_PONoteDocumentiCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto

Set rs = CnDMT.OpenResultset(sSQL)

Ordinamento = 1

If Not rs.EOF Then
    Select Case Tipo
        Case 0
            If Len(Trim(fnNotNull(rs!Annotazioni1))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione01) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni1)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni3))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione03) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni3)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni4))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione04) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni4)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni5))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione05) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
        Case 1
            If Len(Trim(fnNotNull(rs!Annotazioni6))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione06) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni6)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni7))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione07) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni7)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni8))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione08) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni8)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni9))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione09) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni9)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni10))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione10) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni10)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni5))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione05) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
        Case 2
            If Len(Trim(fnNotNull(rs!Annotazioni11))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione11) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni11)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni12))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione12) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni12)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni13))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione13) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni13)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni14))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione14) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni14)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni15))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione15) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni15)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
            If Len(Trim(fnNotNull(rs!Annotazioni5))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione05) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 255)
                        rsAdd!Ordinamento = Ordinamento
                        Ordinamento = Ordinamento + 1
                    rsAdd.Update
                End If
            End If
    End Select
End If

rs.CloseResultset
Set rs = Nothing

ADD_NOTE_TIPO_OGGETTO = Ordinamento
Exit Function
ERR_ADD_NOTE_TIPO_OGGETTO:
    MsgBox Err.Description, vbCritical, "ADD_NOTE_TIPO_OGGETTO"
End Function
Private Function ADD_NOTA_LETTERA_INTENTO(IDTipoOggetto As Long, IDOggetto As Long, rsAdd As ADODB.Recordset, IDLetteraIntento As Long, Ordinamento As Long) As Long
On Error GoTo ERR_ADD_NOTA_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Testo As String
Dim rsNota As DmtOleDbLib.adoResultset
Dim Nota2 As String

If IDLetteraIntento = 0 Then Exit Function

Nota2 = ""
Testo = ""

sSQL = "SELECT Annotazioni2 FROM RV_PONoteDocumentiCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDTipoOggetto=" & ObjDoc.IDTipoOggetto

Set rsNota = CnDMT.OpenResultset(sSQL)

If Not rsNota.EOF Then
    Nota2 = fnNotNull(rsNota!Annotazioni2)
End If

rsNota.CloseResultset
Set rsNota = Nothing

sSQL = "SELECT IDLetteraIntento, IDTipoLetteraIntento, IDAzienda, Data, Numero, Anno, NumeroCliFor, AnnoCliFor, "
sSQL = sSQL & "IDAnagrafica_CF, IDTipoAnagrafica_CF, IDAzienda_CF, ProgressivoDichiarazione, ProtocolloDichiarazione, DataEmissione "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Testo = Nota2 & " " & fnNotNull(rs!ProtocolloDichiarazione) & "/" & fnNotNull(rs!ProgressivoDichiarazione)
    Testo = Testo & " del " & fnNotNull(rs!Data)
    Testo = Testo & " - nr. c/o cliente " & fnNotNull(rs!NumeroCliFor)
    Testo = Testo & " del " & fnNotNull(rs!DataEmissione)
End If

rs.CloseResultset
Set rs = Nothing

If Len(Testo) > 0 Then
    rsAdd.AddNew
        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
        rsAdd!IDBloccoXML = 8
        rsAdd!IDOggetto = IDOggetto
        rsAdd!IDTipoOggetto = IDTipoOggetto
        rsAdd!Annotazioni = Mid(Trim(Testo), 1, 255)
        rsAdd!Ordinamento = Ordinamento
        Ordinamento = Ordinamento + 1
    rsAdd.Update
End If
Exit Function
ERR_ADD_NOTA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "ERR_ADD_NOTA_LETTERA_INTENTO"
End Function
Private Sub ADD_NOTE_IVA(IDTipoOggetto As Long, IDOggetto As Long, rsAdd As ADODB.Recordset, Ordinamento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NomeTabella As String

NomeTabella = sTabellaIVA

sSQL = "SELECT " & sTabellaIVA & ".IDValoriOggettoDettaglio, " & sTabellaIVA & ".IDOggetto, " & sTabellaIVA & ".IDTipoOggetto, " & sTabellaIVA & ".Link_Cst_IVA, Iva.Annotazioni"
sSQL = sSQL & " FROM " & sTabellaIVA & " INNER JOIN "
sSQL = sSQL & " Iva ON " & sTabellaIVA & ".Link_Cst_IVA = Iva.IDIva "
sSQL = sSQL & " WHERE (" & sTabellaIVA & ".IDOggetto = " & IDOggetto & ")"

Set rs = CnDMT.OpenResultset(sSQL)

While Not rs.EOF
    If Len(Trim(rs!Annotazioni)) > 0 Then
        rsAdd.AddNew
            rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsAdd!IDBloccoXML = 8
            rsAdd!IDOggetto = IDOggetto
            rsAdd!IDTipoOggetto = IDTipoOggetto
            rsAdd!Annotazioni = Mid(Trim(rs!Annotazioni), 1, 255)
            rsAdd!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsAdd.Update
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_ADD_NOTA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "ADD_NOTA_LETTERA_INTENTO"

End Sub
Private Function GET_VETTORE_SUCCESSIVO(IDVettore As Long) As String
On Error GoTo ERR_GET_VETTORE_SUCCESSIVO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_VETTORE_SUCCESSIVO = ""

sSQL = "SELECT * FROM Vettore "
sSQL = sSQL & "WHERE IDVettore=" & IDVettore

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_VETTORE_SUCCESSIVO = "Vettore successivo: " & fnNotNull(rs!Vettore)
    If Len(Trim(fnNotNull(rs!PartitaIva))) > 0 Then
        GET_VETTORE_SUCCESSIVO = GET_VETTORE_SUCCESSIVO + " - Partita I.V.A.: " & fnNotNull(rs!PartitaIva)
    End If
    If Len(Trim(fnNotNull(rs!NumeroAlbo))) > 0 Then
        GET_VETTORE_SUCCESSIVO = GET_VETTORE_SUCCESSIVO + " - Iscrizione albo: " & fnNotNull(rs!NumeroAlbo)
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_VETTORE_SUCCESSIVO:
    MsgBox Err.Description, vbCritical, "GET_VETTORE_SUCCESSIVO"
End Function

Private Function GET_AGENZIA_TRASPORTO(IDAgenziaTrasporto) As String
On Error GoTo ERR_GET_AGENZIA_TRASPORTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_AGENZIA_TRASPORTO = ""

sSQL = "SELECT IDAnagrafica, Anagrafica, Nome, PartitaIva "
sSQL = sSQL & "FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAgenziaTrasporto

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_AGENZIA_TRASPORTO = "Agenzia di trasporto: " & fnNotNull(rs!Anagrafica) + fnNotNull(rs!Nome)
    If Len(Trim(fnNotNull(rs!PartitaIva))) > 0 Then
        GET_AGENZIA_TRASPORTO = GET_AGENZIA_TRASPORTO + " - Partita I.V.A.: " & fnNotNull(rs!PartitaIva)
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_AGENZIA_TRASPORTO:
    MsgBox Err.Description, vbCritical, "GET_AGENZIA_TRASPORTO"
End Function
Private Function GET_SEZ_PER_CLIENTE(IDTipoOggetto As Long, IDCliente As Long) As Long
On Error GoTo ERR_GET_SEZ_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_SEZ_PER_CLIENTE = 0

sSQL = "SELECT IDRV_POConfigurazioneCliente, IDSezionalePerDDT, IDSezionalePerFA, IDSezionalePerSNF "
sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
'sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAnagrafica=" & IDCliente

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Select Case IDTipoOggetto
        Case 2
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerDDT)
        Case 114
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerFA)
        Case 8
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerSNF)
    End Select
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_SEZ_PER_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_SEZ_PER_CLIENTE"

End Function
Private Function GET_CONTROLLO_NUM_ORD_DA_EVADERE() As Long
On Error GoTo ERR_GET_CONTROLLO_NUM_ORD_DA_EVADERE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_NUM_ORD_DA_EVADERE = 0

sSQL = "SELECT COUNT(NumeroRiga) AS Numero "
sSQL = sSQL & " FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND DaRegistrare=" & fnNormBoolean(1)

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_NUM_ORD_DA_EVADERE = fnNotNullN(rs!Numero)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_NUM_ORD_DA_EVADERE:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_NUM_ORD_DA_EVADERE"
End Function
Private Function GET_IDCLIENTE_DA_EVADERE() As Long
On Error GoTo ERR_GET_IDCLIENTE_DA_EVADERE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_IDCLIENTE_DA_EVADERE = 0

sSQL = "SELECT IDCliente "
sSQL = sSQL & "FROM RV_POTMPEvasioneOrdini "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND DaRegistrare=" & fnNormBoolean(1)

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_IDCLIENTE_DA_EVADERE = fnNotNullN(rs!IDCliente)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_IDCLIENTE_DA_EVADERE:
    MsgBox Err.Description, vbCritical, "GET_IDCLIENTE_DA_EVADERE"
End Function
Private Sub RECUPERA_CONFIG_CAUS_XML()
On Error GoTo ERR_RECUPERA_CONFIG_CAUS_XML
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Rip_InXMLRifLetteraIntento = 0
Rip_InXMLRifNoteIva = 0
Rip_InXMLRifNota01Doc = 0
Rip_InXMLRifNota02Doc = 0
Rip_InXMLRifNota03Doc = 0
Rip_InXMLRifNotaDoc = 0
Rip_InXMLRifIstrMitt = 0
Rip_InXMLRifVettSucc = 0
Rip_InXMLRifAgenziaTrasp = 0
Rip_InXMLRifTargaAutoMezzo = 0
NonRiportaInXMLRifVsNumOrd = 0

sSQL = "SELECT RiportaInXMLRifLetteraIntento, RiportaInXMLRifNoteIva, RiportaInXMLRifNota01Doc, "
sSQL = sSQL & "RiportaInXMLRifNota02Doc, RiportaInXMLRifNota03Doc, RiportaInXMLRifNotaDoc, "
sSQL = sSQL & "RiportaInXMLRifIstrMitt, RiportaInXMLRifVettSucc, RiportaInXMLRifAgenziaTrasp, "
sSQL = sSQL & "RiportaInXMLRifTargaAutoMezzo, NonRiportaInXMLRifVsNumOrd "
sSQL = sSQL & " FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Rip_InXMLRifLetteraIntento = fnNotNullN(rs!RiportaInXMLRifLetteraIntento)
    Rip_InXMLRifNoteIva = fnNotNullN(rs!RiportaInXMLRifNoteIva)
    Rip_InXMLRifNota01Doc = fnNotNullN(rs!RiportaInXMLRifNota01Doc)
    Rip_InXMLRifNota02Doc = fnNotNullN(rs!RiportaInXMLRifNota02Doc)
    Rip_InXMLRifNota03Doc = fnNotNullN(rs!RiportaInXMLRifNota03Doc)
    Rip_InXMLRifNotaDoc = fnNotNullN(rs!RiportaInXMLRifNotaDoc)
    Rip_InXMLRifIstrMitt = fnNotNullN(rs!RiportaInXMLRifIstrMitt)
    Rip_InXMLRifVettSucc = fnNotNullN(rs!RiportaInXMLRifVettSucc)
    Rip_InXMLRifAgenziaTrasp = fnNotNullN(rs!RiportaInXMLRifAgenziaTrasp)
    Rip_InXMLRifTargaAutoMezzo = fnNotNullN(rs!RiportaInXMLRifTargaAutoMezzo)
    NonRiportaInXMLRifVsNumOrd = fnNotNullN(rs!NonRiportaInXMLRifVsNumOrd)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_RECUPERA_CONFIG_CAUS_XML:
    MsgBox Err.Description, vbCritical, "RECUPERA_CONFIG_CAUS_XML"
End Sub
Private Sub GET_IVA_ART_DA_ORDINE(ID As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, IDTipoOggetto, RV_POLinkRiga, "
sSQL = sSQL & "RV_POTipoRiga, Art_aliquota_IVA, Link_Art_IVA, RV_POIDImballo"
sSQL = sSQL & " FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & ID

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
    Link_IVAArticolo = fnNotNullN(rs!Link_Art_IVA)
    AliquotaIvaArticolo = fnNotNullN(rs!Art_aliquota_IVA)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_IVA_IMB_DA_ORDINE(ID As Long, IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim linkRiga As Long

linkRiga = 0

sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, IDTipoOggetto, RV_POLinkRiga, "
sSQL = sSQL & "RV_POTipoRiga, Art_aliquota_IVA, Link_Art_IVA, RV_POIDImballo"
sSQL = sSQL & " FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & ID

Set rs = CnDMT.OpenResultset(sSQL)

If Not rs.EOF Then
  linkRiga = fnNotNullN(rs!RV_POLinkRiga)
End If

rs.CloseResultset
Set rs = Nothing

If (linkRiga > 0) Then
    sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, IDTipoOggetto, RV_POLinkRiga, "
    sSQL = sSQL & "RV_POTipoRiga, Art_aliquota_IVA, Link_Art_IVA, RV_POIDImballo"
    sSQL = sSQL & " FROM ValoriOggettoDettaglio0010 "
    sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
    sSQL = sSQL & " AND RV_POLinkRiga=" & linkRiga
    sSQL = sSQL & " AND RV_POTipoRiga=2"
    
    Set rs = CnDMT.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        Link_IVAArticolo = fnNotNullN(rs!Link_Art_IVA)
        AliquotaIvaArticolo = fnNotNullN(rs!Art_aliquota_IVA)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End If

End Sub
