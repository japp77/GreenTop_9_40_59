VERSION 5.00
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{F95AA20B-3F80-11D3-A741-00105A2E9BAF}#2.1#0"; "DmtSearchAccount2.ocx"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FrmMain 
   Caption         =   "Controllo vendite"
   ClientHeight    =   13785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20085
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13785
   ScaleWidth      =   20085
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   13080
      TabIndex        =   14
      Top             =   3000
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   255
      Left            =   14280
      TabIndex        =   13
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   13575
      Left            =   0
      ScaleHeight     =   13545
      ScaleWidth      =   19185
      TabIndex        =   11
      Top             =   0
      Width           =   19215
      Begin VB.Frame FraStampa 
         Caption         =   "Stampa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   29
         Top             =   5640
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton cmdStampa 
            Enabled         =   0   'False
            Height          =   375
            Left            =   5400
            Picture         =   "FrmMain.frx":4781A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "STAMPA REPORT"
            Top             =   160
            Width           =   375
         End
         Begin DmtActBox.DmtActBoxCtl ActivityBox 
            Height          =   1095
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1931
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
      End
      Begin VB.Frame Frame1 
         Height          =   5655
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   19095
         Begin DMTEDITNUMLib.dmtNumber txtAnnoProcesso 
            Height          =   315
            Left            =   14400
            TabIndex        =   130
            Top             =   4440
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
         Begin VB.CommandButton cmdEspandiGriglia 
            Height          =   495
            Left            =   5160
            Picture         =   "FrmMain.frx":47DA4
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Stampe"
            Top             =   5100
            Width           =   615
         End
         Begin VB.TextBox txtNoteAgg 
            Height          =   315
            Left            =   9480
            ScrollBars      =   2  'Vertical
            TabIndex        =   120
            Top             =   4440
            Width           =   4815
         End
         Begin VB.TextBox txtRaggrRigaOrdine 
            Height          =   315
            Left            =   11040
            TabIndex        =   118
            Top             =   3840
            Width           =   2295
         End
         Begin DmtSearchAccount2.DmtSearchACS2 ACSCliente 
            Height          =   600
            Left            =   120
            TabIndex        =   95
            Top             =   120
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   1058
            WidthCode       =   700
            WidthDescription=   3600
            WidthSecondDescription=   1300
            Object.Visible         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            CaptionDescription=   "Cliente"
            CaptionCode     =   "Codice"
            OnlyAccounts    =   -1  'True
         End
         Begin VB.TextBox txtCodicePedana 
            Height          =   315
            Left            =   120
            TabIndex        =   91
            Top             =   3360
            Width           =   3735
         End
         Begin VB.TextBox txtDescrizioneArticoloConf 
            Height          =   315
            Left            =   11280
            TabIndex        =   72
            Top             =   1460
            Width           =   3855
         End
         Begin VB.TextBox txtDescrizioneArticolo 
            Height          =   315
            Left            =   1920
            TabIndex        =   69
            Top             =   970
            Width           =   3855
         End
         Begin VB.TextBox txtDescrizioneImballo 
            Height          =   315
            Left            =   1920
            TabIndex        =   68
            Top             =   1570
            Width           =   3855
         End
         Begin VB.CommandButton cmdConferimento 
            Height          =   495
            Left            =   2760
            Picture         =   "FrmMain.frx":4832E
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Vai al documento del conferimento"
            Top             =   5100
            Width           =   615
         End
         Begin VB.CommandButton cmdLavorazione 
            Height          =   495
            Left            =   2160
            Picture         =   "FrmMain.frx":488B8
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Vai alla lavorazione del conferimento"
            Top             =   5100
            Width           =   615
         End
         Begin VB.CommandButton cmdVendita 
            Height          =   495
            Left            =   1560
            Picture         =   "FrmMain.frx":48E42
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   5100
            Width           =   615
         End
         Begin VB.CommandButton cmdPulisciFiltri 
            Height          =   495
            Left            =   720
            Picture         =   "FrmMain.frx":493CC
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Pulisci filtri"
            Top             =   5100
            Width           =   615
         End
         Begin VB.CommandButton cmdEseguiRicerca 
            Height          =   495
            Left            =   120
            Picture         =   "FrmMain.frx":49956
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   5100
            Width           =   615
         End
         Begin DMTEDITNUMLib.dmtNumber txtDaNumeroDocumento 
            Height          =   255
            Left            =   6360
            TabIndex        =   55
            Top             =   1000
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboTipoDocumentoCoop 
            Height          =   315
            Left            =   9480
            TabIndex        =   37
            Top             =   2640
            Width           =   2655
            _ExtentX        =   4683
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
         Begin VB.TextBox txtLottoVendita 
            Height          =   315
            Left            =   15360
            TabIndex        =   28
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox txtLottoConferimento 
            Height          =   315
            Left            =   15360
            TabIndex        =   26
            Top             =   900
            Width           =   3375
         End
         Begin VB.TextBox txtLottoDiCampagna 
            Height          =   315
            Left            =   15360
            TabIndex        =   24
            Top             =   360
            Width           =   3375
         End
         Begin VB.ComboBox cboTipoDocumento 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2760
            Width           =   2895
         End
         Begin DMTDATETIMELib.dmtDate txtDaDataVendita 
            Height          =   285
            Left            =   6360
            TabIndex        =   3
            Top             =   390
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DmtCodDescCtl.DmtCodDesc cdCliente 
            Height          =   615
            Left            =   120
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":49EE0
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":49F2E
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":49F80
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
         Begin DmtCodDescCtl.DmtCodDesc CDArticoloVenduto 
            Height          =   615
            Left            =   120
            TabIndex        =   1
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":49FDA
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4A032
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4A09A
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
         Begin DmtCodDescCtl.DmtCodDesc CDImballoVendita 
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   1320
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":4A0F4
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4A14B
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4A1C0
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
         Begin DMTDATETIMELib.dmtDate txtADataVendita 
            Height          =   285
            Left            =   8040
            TabIndex        =   4
            Top             =   390
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDaDataConferimento 
            Height          =   285
            Left            =   6360
            TabIndex        =   5
            Top             =   1575
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtADataConferimento 
            Height          =   285
            Left            =   8040
            TabIndex        =   6
            Top             =   1575
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DmtCodDescCtl.DmtCodDesc CDSocio 
            Height          =   615
            Left            =   9480
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   640
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":4A21A
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4A268
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4A2C2
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
         Begin DmtCodDescCtl.DmtCodDesc CDArticoloConferito 
            Height          =   615
            Left            =   9480
            TabIndex        =   8
            Top             =   1200
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":4A31C
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4A374
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4A3DE
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
         Begin DMTDATETIMELib.dmtDate txtDaDataConsegnaMerce 
            Height          =   285
            Left            =   6360
            TabIndex        =   32
            Top             =   2205
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtADataConsegnaMerce 
            Height          =   285
            Left            =   8040
            TabIndex        =   33
            Top             =   2205
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboCategoriaFiscale 
            Height          =   315
            Left            =   15360
            TabIndex        =   40
            Top             =   2040
            Width           =   1575
            _ExtentX        =   2778
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
         Begin DmtCodDescCtl.DmtCodDesc CDSocioFatt 
            Height          =   615
            Left            =   9480
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1800
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":4A438
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4A486
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4A4EB
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
         Begin DMTDataCmb.DMTCombo cboCategoriaMerceologica 
            Height          =   315
            Left            =   17040
            TabIndex        =   42
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
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
         Begin DMTDataCmb.DMTCombo cboDestinazione 
            Height          =   315
            Left            =   3120
            TabIndex        =   44
            Top             =   2760
            Width           =   2655
            _ExtentX        =   4683
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
         Begin DMTDataCmb.DMTCombo cboTipoImportoLiq 
            Height          =   315
            Left            =   9480
            TabIndex        =   46
            Top             =   360
            Width           =   3375
            _ExtentX        =   5953
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
         Begin DMTDataCmb.DMTCombo cboPrezzoMedioInLiq 
            Height          =   315
            Left            =   14160
            TabIndex        =   49
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
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
         Begin DMTDataCmb.DMTCombo cboNessunaForzatura 
            Height          =   315
            Left            =   12960
            TabIndex        =   50
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
         Begin DMTEDITNUMLib.dmtNumber txtANumeroDocumento 
            Height          =   255
            Left            =   8040
            TabIndex        =   56
            Top             =   1000
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboConferimentoChiuso 
            Height          =   315
            Left            =   12240
            TabIndex        =   59
            Top             =   2640
            Width           =   1215
            _ExtentX        =   2143
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
         Begin DMTLblLinkCtl.LabelLink LabelLink1 
            Height          =   255
            Left            =   3960
            TabIndex        =   65
            Top             =   0
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Caption         =   "Vendita"
            Name            =   "LabelLink"
         End
         Begin DMTLblLinkCtl.LabelLink LabelLink2 
            Height          =   255
            Left            =   2640
            TabIndex        =   66
            Top             =   0
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Caption         =   "Conferimento"
            Name            =   "LabelLink"
         End
         Begin DMTLblLinkCtl.LabelLink LabelLink3 
            Height          =   255
            Left            =   4920
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            Caption         =   "Lavorazione"
            Name            =   "LabelLink"
         End
         Begin DMTDataCmb.DMTCombo cboSezionale 
            Height          =   315
            Left            =   120
            TabIndex        =   74
            Top             =   4560
            Width           =   2775
            _ExtentX        =   4895
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
         Begin DMTDATETIMELib.dmtDate txtDaLavMerce 
            Height          =   285
            Left            =   6360
            TabIndex        =   76
            Top             =   2805
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtALavMerce 
            Height          =   285
            Left            =   8040
            TabIndex        =   77
            Top             =   2805
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboPrezzoMedioConf 
            Height          =   315
            Left            =   13560
            TabIndex        =   81
            Top             =   2640
            Width           =   1575
            _ExtentX        =   2778
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
         Begin DMTDataCmb.DMTCombo cboTipoLavorazione 
            Height          =   315
            Left            =   12240
            TabIndex        =   83
            Top             =   3240
            Width           =   2895
            _ExtentX        =   5106
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
         Begin DMTDataCmb.DMTCombo cboTipoCategoria 
            Height          =   315
            Left            =   15240
            TabIndex        =   84
            Top             =   3240
            Width           =   1815
            _ExtentX        =   3201
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
         Begin DMTDataCmb.DMTCombo cboCalibro 
            Height          =   315
            Left            =   17160
            TabIndex        =   85
            Top             =   3240
            Width           =   1575
            _ExtentX        =   2778
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
         Begin DMTDataCmb.DMTCombo cboTipoLavConf 
            Height          =   315
            Left            =   9480
            TabIndex        =   89
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
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
         Begin DmtCodDescCtl.DmtCodDesc CDTipoPedana 
            Height          =   615
            Left            =   3960
            TabIndex        =   93
            Top             =   3110
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":4A545
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4A599
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4A60E
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
         Begin VB.CommandButton cmdRiduciGriglia 
            Height          =   375
            Left            =   5280
            Picture         =   "FrmMain.frx":4A668
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "Riduci maschera"
            Top             =   5160
            Visible         =   0   'False
            Width           =   375
         End
         Begin DMTDATETIMELib.dmtDate txtDaDataArrivoMerce 
            Height          =   285
            Left            =   6360
            TabIndex        =   98
            Top             =   3405
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtADataArrivoMerce 
            Height          =   285
            Left            =   8040
            TabIndex        =   99
            Top             =   3405
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDaDataOrdine 
            Height          =   285
            Left            =   6360
            TabIndex        =   103
            Top             =   4005
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtADataOrdine 
            Height          =   285
            Left            =   8040
            TabIndex        =   104
            Top             =   4005
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdine 
            Height          =   315
            Left            =   4440
            TabIndex        =   108
            Top             =   3975
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   253
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DmtCodDescCtl.DmtCodDesc CDClienteOrdine 
            Height          =   615
            Left            =   120
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   3720
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":4ABF2
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4AC40
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4AC99
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
         Begin DMTDataCmb.DMTCombo cboRiscontroPeso 
            Height          =   315
            Left            =   9480
            TabIndex        =   111
            Top             =   3840
            Width           =   1455
            _ExtentX        =   2566
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
         Begin DMTDATETIMELib.dmtDate txtDaDataOrdineCli 
            Height          =   285
            Left            =   6360
            TabIndex        =   113
            Top             =   4605
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtADataOrdineCli 
            Height          =   285
            Left            =   8040
            TabIndex        =   114
            Top             =   4605
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Left            =   4560
            Picture         =   "FrmMain.frx":4ACF3
            Style           =   1  'Graphical
            TabIndex        =   122
            ToolTipText     =   "Salva colonne griglia"
            Top             =   5100
            Width           =   615
         End
         Begin DmtCodDescCtl.DmtCodDesc CDImballoPrimario 
            Height          =   615
            Left            =   120
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   1920
            Width           =   5730
            _ExtentX        =   10107
            _ExtentY        =   1085
            PropCodice      =   $"FrmMain.frx":4B27D
            BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PropDescrizione =   $"FrmMain.frx":4B2CC
            BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MenuFunctions   =   $"FrmMain.frx":4B334
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
         Begin DMTDataCmb.DMTCombo cboTipoClassVend 
            Height          =   315
            Left            =   17040
            TabIndex        =   124
            Top             =   3840
            Width           =   1695
            _ExtentX        =   2990
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
         Begin DMTDataCmb.DMTCombo cboTipoClassConf 
            Height          =   315
            Left            =   15240
            TabIndex        =   125
            Top             =   3840
            Width           =   1695
            _ExtentX        =   2990
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
         Begin DMTDataCmb.DMTCombo cboCatLiq 
            Height          =   315
            Left            =   13440
            TabIndex        =   128
            Top             =   3840
            Width           =   1695
            _ExtentX        =   2990
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
         Begin DMTEDITNUMLib.dmtNumber txtNumeroProcesso 
            Height          =   315
            Left            =   15960
            TabIndex        =   131
            Top             =   4440
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   253
            Text            =   "0"
            BackColor       =   16777215
            Appearance      =   1
            AllowEmpty      =   0   'False
         End
         Begin DMTDATETIMELib.dmtDate txtDataProcesso 
            Height          =   315
            Left            =   17280
            TabIndex        =   132
            Top             =   4440
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtDaDataCompLiq 
            Height          =   285
            Left            =   6360
            TabIndex        =   136
            Top             =   5205
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDATETIMELib.dmtDate txtADataCompLiq 
            Height          =   285
            Left            =   8040
            TabIndex        =   137
            Top             =   5205
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   253
            ForeColor       =   0
            BackColor       =   16777215
            Appearance      =   1
         End
         Begin DMTDataCmb.DMTCombo cboSezionaleConf 
            Height          =   315
            Left            =   3000
            TabIndex        =   141
            Top             =   4560
            Width           =   2775
            _ExtentX        =   4895
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
         Begin DMTDataCmb.DMTCombo cboVettoreConf 
            Height          =   315
            Left            =   9480
            TabIndex        =   143
            Top             =   5040
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
         Begin DMTDataCmb.DMTCombo cboLuogoPresaMerceConf 
            Height          =   315
            Left            =   12840
            TabIndex        =   145
            Top             =   5040
            Width           =   2775
            _ExtentX        =   4895
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
         Begin DMTDataCmb.DMTCombo cboTipoConformeConf 
            Height          =   315
            Left            =   15720
            TabIndex        =   147
            Top             =   5040
            Width           =   3015
            _ExtentX        =   5318
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
         Begin VB.Line Line8 
            X1              =   18840
            X2              =   18840
            Y1              =   240
            Y2              =   5400
         End
         Begin VB.Label Label8 
            Caption         =   "Conformità"
            Height          =   255
            Index           =   3
            Left            =   15720
            TabIndex        =   148
            Top             =   4800
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Sede stoccaggio merce"
            Height          =   255
            Index           =   2
            Left            =   12840
            TabIndex        =   146
            Top             =   4800
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "Vettore conf."
            Height          =   255
            Index           =   14
            Left            =   9480
            TabIndex        =   144
            Top             =   4800
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Sezionale conf."
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   142
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data competenza liquidazione"
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
            Index           =   8
            Left            =   6000
            TabIndex        =   140
            Top             =   4920
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   8
            Left            =   6000
            TabIndex        =   139
            Top             =   5205
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   8
            Left            =   7800
            TabIndex        =   138
            Top             =   5205
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "Data processo"
            Height          =   255
            Index           =   13
            Left            =   17280
            TabIndex        =   135
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "N° processo"
            Height          =   255
            Index           =   12
            Left            =   15960
            TabIndex        =   134
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Anno processo"
            Height          =   255
            Index           =   11
            Left            =   14400
            TabIndex        =   133
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Cat. liq."
            Height          =   255
            Index           =   10
            Left            =   13440
            TabIndex        =   129
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo class. vend."
            Height          =   255
            Index           =   9
            Left            =   17040
            TabIndex        =   127
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo class. conf."
            Height          =   255
            Index           =   8
            Left            =   15240
            TabIndex        =   126
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Annotazioni aggiuntive"
            Height          =   255
            Index           =   5
            Left            =   9480
            TabIndex        =   121
            ToolTipText     =   "Nessuna forzatura in liquidazione"
            Top             =   4200
            Width           =   2655
         End
         Begin VB.Label Label13 
            Caption         =   "Sub lotto"
            Height          =   255
            Index           =   1
            Left            =   11040
            TabIndex        =   119
            Top             =   3600
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   7
            Left            =   7800
            TabIndex        =   117
            Top             =   4605
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   7
            Left            =   6000
            TabIndex        =   116
            Top             =   4605
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data ordine cliente"
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
            Index           =   7
            Left            =   6000
            TabIndex        =   115
            Top             =   4320
            Width           =   3255
         End
         Begin VB.Label Label10 
            Caption         =   "Riscontro peso"
            Height          =   255
            Index           =   4
            Left            =   9480
            TabIndex        =   112
            ToolTipText     =   "Nessuna forzatura in liquidazione"
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label Label10 
            Caption         =   "Numero ordine"
            Height          =   255
            Index           =   3
            Left            =   4440
            TabIndex        =   109
            Top             =   3765
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data ordine interno"
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
            Index           =   6
            Left            =   6000
            TabIndex        =   107
            Top             =   3720
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   6
            Left            =   6000
            TabIndex        =   106
            Top             =   4005
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   6
            Left            =   7800
            TabIndex        =   105
            Top             =   4005
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data arrivo merce"
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
            Index           =   5
            Left            =   6000
            TabIndex        =   102
            Top             =   3120
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   5
            Left            =   6000
            TabIndex        =   101
            Top             =   3405
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   5
            Left            =   7800
            TabIndex        =   100
            Top             =   3405
            Width           =   255
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            X1              =   120
            X2              =   5760
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Label Label11 
            Caption         =   "Codice pedana"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "Prz. medio conf."
            Height          =   255
            Index           =   3
            Left            =   13560
            TabIndex        =   82
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Line Line5 
            X1              =   9480
            X2              =   18720
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   4
            Left            =   7800
            TabIndex        =   80
            Top             =   2805
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   79
            Top             =   2805
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data lavorazione merce"
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
            Index           =   4
            Left            =   6000
            TabIndex        =   78
            Top             =   2520
            Width           =   3255
         End
         Begin VB.Label Label9 
            Caption         =   "Sezionale vend."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label lblArticoloConferito 
            Caption         =   "Descrizione articolo conferito"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   11280
            TabIndex        =   73
            Top             =   1230
            Width           =   3855
         End
         Begin VB.Label lblDescrizioneArticolo 
            Caption         =   "Descrizione articolo lavorato"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1920
            TabIndex        =   71
            Top             =   750
            Width           =   3855
         End
         Begin VB.Label lblDescrizioneImballo 
            Caption         =   "Descrizione imballo lavorato"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1920
            TabIndex        =   70
            Top             =   1350
            Width           =   3855
         End
         Begin VB.Label Label7 
            Caption         =   "Conf. Chiuso"
            Height          =   255
            Index           =   2
            Left            =   12240
            TabIndex        =   60
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   3
            Left            =   6000
            TabIndex        =   54
            Top             =   1000
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   3
            Left            =   7800
            TabIndex        =   53
            Top             =   1000
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Numero documento"
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
            Index           =   3
            Left            =   6000
            TabIndex        =   52
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label Label10 
            Caption         =   "Ness. Forz."
            Height          =   255
            Index           =   2
            Left            =   12960
            TabIndex        =   51
            ToolTipText     =   "Nessuna forzatura in liquidazione"
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Part. P.M."
            Height          =   255
            Index           =   1
            Left            =   14160
            TabIndex        =   48
            ToolTipText     =   "Partecipa al prezzo medio in liquidazione"
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo forzatura prezzo di liquidazione"
            Height          =   255
            Index           =   0
            Left            =   9480
            TabIndex        =   47
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label Label8 
            Caption         =   "Cat. merc."
            Height          =   255
            Index           =   0
            Left            =   17040
            TabIndex        =   43
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Line Line4 
            X1              =   6000
            X2              =   9240
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label Label9 
            Caption         =   "Categoria fiscale"
            Height          =   255
            Index           =   0
            Left            =   15360
            TabIndex        =   39
            Top             =   1815
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo doc. di consegna merce"
            Height          =   255
            Index           =   1
            Left            =   9480
            TabIndex        =   38
            Top             =   2400
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data di consegna merce"
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
            Index           =   2
            Left            =   6000
            TabIndex        =   36
            Top             =   1920
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   2
            Left            =   6000
            TabIndex        =   35
            Top             =   2205
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   2
            Left            =   7800
            TabIndex        =   34
            Top             =   2205
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "Lotto di vendita"
            Height          =   255
            Index           =   0
            Left            =   15360
            TabIndex        =   27
            Top             =   1240
            Width           =   3375
         End
         Begin VB.Label Label6 
            Caption         =   "Lotto di entrata"
            Height          =   255
            Left            =   15360
            TabIndex        =   25
            Top             =   700
            Width           =   3375
         End
         Begin VB.Label Label5 
            Caption         =   "Lotto di produzione"
            Height          =   255
            Left            =   15360
            TabIndex        =   23
            Top             =   120
            Width           =   3375
         End
         Begin VB.Line Line3 
            X1              =   15240
            X2              =   15240
            Y1              =   240
            Y2              =   2880
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo documento"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2520
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   1
            Left            =   7800
            TabIndex        =   21
            Top             =   1575
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   1
            Left            =   6000
            TabIndex        =   20
            Top             =   1575
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data di liquidazione"
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
            Index           =   1
            Left            =   6000
            TabIndex        =   19
            Top             =   1305
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "A"
            Height          =   255
            Index           =   0
            Left            =   7800
            TabIndex        =   18
            Top             =   390
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "DA"
            Height          =   255
            Index           =   0
            Left            =   6000
            TabIndex        =   17
            Top             =   390
            Width           =   375
         End
         Begin VB.Line Line2 
            X1              =   9360
            X2              =   9360
            Y1              =   240
            Y2              =   5520
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data di vendita"
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
            Left            =   6000
            TabIndex        =   16
            Top             =   120
            Width           =   3255
         End
         Begin VB.Line Line1 
            X1              =   5895
            X2              =   5895
            Y1              =   240
            Y2              =   5520
         End
         Begin VB.Label Label8 
            Caption         =   "Destinazione diversa"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   45
            Top             =   2520
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo categoria"
            Height          =   255
            Index           =   6
            Left            =   15240
            TabIndex        =   87
            Top             =   3040
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Calibro"
            Height          =   255
            Index           =   7
            Left            =   17160
            TabIndex        =   86
            Top             =   3040
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo lavorazione conferito"
            Height          =   255
            Index           =   4
            Left            =   9480
            TabIndex        =   90
            Top             =   3040
            Width           =   2655
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo lavorazione merce"
            Height          =   255
            Index           =   5
            Left            =   12240
            TabIndex        =   88
            Top             =   3040
            Width           =   2895
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8535
         Left            =   0
         ScaleHeight     =   8535
         ScaleWidth      =   19095
         TabIndex        =   12
         Top             =   4920
         Width           =   19095
         Begin DmtGridCtl.DmtGrid GrigliaControlloVendite 
            Height          =   7695
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   18975
            _ExtentX        =   33470
            _ExtentY        =   13573
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
         Begin MSComctlLib.StatusBar StatusBar1 
            Height          =   30
            Left            =   120
            TabIndex        =   58
            Top             =   8160
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   53
            _Version        =   393216
            BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
               NumPanels       =   1
               BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtVista 
            Height          =   7215
            Left            =   11400
            MultiLine       =   -1  'True
            TabIndex        =   94
            Text            =   "FrmMain.frx":4B38E
            Top             =   840
            Visible         =   0   'False
            Width           =   6135
         End
      End
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   5280
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'L'applicazione corrente
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
Public rsGriglia As ADODB.Recordset
Public gPaintNotify As PaintNotify
Public oReport As dmtReportLib.dmtReport

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


Private FILTRO_DA_DATA_VENDITA As Long
Private FILTRO_A_DATA_VENDITA As Long

Private FILTRO_DA_DATA_LIQUIDAZIONE As Long
Private FILTRO_A_DATA_LIQUIDAZIONE As Long

Private FILTRO_DA_DATA_CONSEGNA_MERCE As Long
Private FILTRO_A_DATA_CONSEGNA_MERCE As Long

Private BLoading As Boolean



'Variabili recordset per le visualizzazioni delle griglie

Public Sub ConnessioneADO()
Dim oActivity As IActivity
Dim o As Activity
Dim oFilter As Filter

    If Not (CnDMT Is Nothing) Then
        CnDMT.CloseConnection
        Set CnDMT = Nothing
    End If
    Set CnDMT = m_App.Database.Connection
    
    PrelevaAzienda
    
    VarPassword = m_App.Password
    VarUtente = m_App.User
    
    InitControlli
    
    Me.Caption = TheApp.FunctionName & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    'Inizializzazione del riquadro attività
    With ActivityBox
        .Activities.Clear
        
        'Aggiunge l'attività dei reports
        Set oActivity = .Activities.Add("DmtActBoxLib.ReportsActivity", "Reports")
        Set oActivity.Connection = CnDMT.InternalConnection
        
        oActivity.Load fnGetTipoOggetto, TheApp.IDFirm
        Set o = oActivity
        Set oReportsActivity = o.InternalClass
        
        
        'Imposta quale attività deve essere attivata per default
        If m_DefaultActivity <> "" Then
            Set .CurrentActivity = .Activities(m_DefaultActivity)
        End If
        
        'ridisegna il controllo
        .Redraw = True
        
        oReportsActivity.Is4DlgPrint = False
    End With

    Set gPaintNotify = New PaintNotify

End Sub
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

  
  


Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property

Public Sub InitControlli()
     With Me.cdCliente
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Clienti"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

     With Me.CDClienteOrdine
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Clienti"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Anagrafica"
        .CodeIsNumeric = False
    End With


    Set Me.ACSCliente.Connection = TheApp.Database.Connection
    ACSCliente.ApplicationName = App.Title
    ACSCliente.Client = App.EXEName
    ACSCliente.IDFirm = TheApp.IDFirm
    ACSCliente.IDUser = TheApp.IDUser
    ACSCliente.UserName = TheApp.User
    ACSCliente.SearchType = DmtSearchCustomers
    ACSCliente.HwndContainer = Me.hwnd


    With Me.CDSocio
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Socio/Fornitore"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Socio/Fornitore"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
     With Me.CDSocio
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Socio\Fornitore"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Socio\Fornitore"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    
     With Me.CDArticoloVenduto
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With


     With Me.CDImballoVendita
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    With Me.CDArticoloConferito
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    With Me.cboCategoriaMerceologica
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaMerceologica"
        .DisplayField = "CategoriaMerceologica"
        .Sql = "SELECT IDCategoriaMerceologica, CategoriaMerceologica FROM CategoriaMerceologica "
        .Sql = .Sql & " ORDER BY CategoriaMerceologica"
    End With
    
    With Me.cboTipoDocumentoCoop
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoDocumentoCoop"
        .DisplayField = "TipoDocumentoCoop"
        .Sql = "SELECT IDRV_POTipoDocumentoCoop, TipoDocumentoCoop FROM RV_POTipoDocumentoCoop "
        .Sql = .Sql & " ORDER BY TipoDocumentoCoop"
    End With

    With Me.cboCategoriaFiscale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCategoriaFiscale"
        .DisplayField = "CategoriaFiscale"
        .Sql = "SELECT IDCategoriaFiscale, CategoriaFiscale FROM CategoriaFiscale "
        .Sql = .Sql & " ORDER BY CategoriaFiscale"
    End With
    
    With Me.cboDestinazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica"
        .Sql = .Sql & " WHERE IDAnagrafica = " & Me.cdCliente.KeyFieldID
        .Sql = .Sql & " ORDER BY SitoPerAnagrafica"
    End With
    
    Me.cboTipoDocumento.AddItem "Tutti"
    Me.cboTipoDocumento.ItemData(0) = 0
    
    Me.cboTipoDocumento.AddItem "Documento di trasporto"
    Me.cboTipoDocumento.ItemData(1) = 2
    
    Me.cboTipoDocumento.AddItem "Fattura accompagnatoria"
    Me.cboTipoDocumento.ItemData(2) = 114

    Me.cboTipoDocumento.AddItem "Corrispettivi"
    Me.cboTipoDocumento.ItemData(3) = 8

    Me.cboTipoDocumento.AddItem "Nota di credito"
    Me.cboTipoDocumento.ItemData(4) = 11

    Me.cboTipoDocumento.AddItem "Nota di debito"
    Me.cboTipoDocumento.ItemData(5) = 107


    Set Me.LabelLink1.Application = TheApp
    Me.LabelLink1.WindowHandleClient = Me.hwnd
    Me.LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False

    Set Me.LabelLink2.Application = TheApp
    Me.LabelLink2.WindowHandleClient = Me.hwnd
    Me.LabelLink2.PopMenuItems("Mnu_SearchObject").Enabled = False

    Set Me.LabelLink3.Application = TheApp
    Me.LabelLink3.WindowHandleClient = Me.hwnd
    Me.LabelLink3.PopMenuItems("Mnu_SearchObject").Enabled = False
    
    With Me.CDSocioFatt
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & m_App.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Anagrafica"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Forza importo di liquidazione
    With Me.cboTipoImportoLiq
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoImportoVenditaLiq"
        .DisplayField = "TipoImportoVenditaLiq"
        .Sql = "SELECT * FROM RV_POTipoImportoVenditaLiq "
        .Fill
    End With

    'Forza importo di liquidazione
    With Me.cboPrezzoMedioInLiq
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SiNO"
        .Sql = "SELECT * FROM RV_POSiNO "
        .Fill
    End With

    'Non Forzare importo di liquidazione
    With Me.cboNessunaForzatura
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SiNO"
        .Sql = "SELECT * FROM RV_POSiNO "
        .Fill
    End With

    'Riscontro peso
    With Me.cboRiscontroPeso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SiNO"
        .Sql = "SELECT * FROM RV_POSiNO "
        .Fill
    End With

    'Conferimento chiuso
    With Me.cboConferimentoChiuso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SiNO"
        .Sql = "SELECT * FROM RV_POSiNO "
        .Fill
    End With
    
    'Sezionale
    If GET_NUMERO_FILIALE(TheApp.IDFirm) = 1 Then
        sSQL = "SELECT * FROM Sezionale "
        sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
        sSQL = sSQL & " ORDER BY Sezionale"
    Else
        sSQL = "SELECT Sezionale.IDSezionale, Sezionale.IDFiliale, AttivitaAzienda.IDAzienda, Sezionale.Sezionale + ' ( ' + Filiale.Filiale + ' )' AS Sezionale, Filiale.Filiale "
        sSQL = sSQL & "FROM Sezionale INNER JOIN "
        sSQL = sSQL & "Filiale ON Sezionale.IDFiliale = Filiale.IDFiliale INNER JOIN "
        sSQL = sSQL & "AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda "
        sSQL = sSQL & "WHERE (AttivitaAzienda.IDAzienda = " & TheApp.IDFirm & ")"
        sSQL = sSQL & " ORDER BY Sezionale"
    End If
    
    
    With Me.cboSezionale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .Sql = sSQL
        .Fill
    End With
    
    With Me.cboSezionaleConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .Sql = sSQL
        .Fill
    End With
    
    
    'Prezzo medio da conferimento
    With Me.cboPrezzoMedioConf
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POSINO"
        .DisplayField = "SiNO"
        .Sql = "SELECT * FROM RV_POSiNO "
        .Fill
    End With

    With Me.cboTipoLavorazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .Sql = "SELECT * FROM RV_POTipoLavorazione ORDER BY TipoLavorazione"
        .Fill
    End With

    With Me.cboCalibro
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POCalibro"
        .DisplayField = "Calibro"
        .Sql = "SELECT * FROM RV_POCalibro"
        .Sql = .Sql & " ORDER BY Calibro"
    End With
  
    With Me.cboTipoCategoria
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoCategoria"
        .DisplayField = "TipoCategoria"
        .Sql = "SELECT * FROM RV_POTipoCategoria"
        .Sql = .Sql & " ORDER BY TipoCategoria"
    End With

     With Me.cboTipoLavConf
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .Sql = "SELECT * FROM RV_POTipoLavorazione ORDER BY TipoLavorazione"
        .Fill
    End With

     With Me.CDTipoPedana
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceTipoPedana"
        .DescriptionField = "TipoPedana"
        .KeyField = "IDRV_POTipoPedana"
        .TableName = "RV_POTipoPedana"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice"
        .DescriptionCaption4Find = "Descrizione"
        .CodeIsNumeric = False
    End With

     With Me.CDImballoPrimario
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = fncTrovaIDFunzione("Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

     With Me.cboTipoClassConf
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDClassificazioneArticolo"
        .DisplayField = "ClassificazioneArticolo"
        .Sql = "SELECT * FROM ClassificazioneArticolo WHERE IDAzienda=" & TheApp.IDFirm & " ORDER BY ClassificazioneArticolo"
        .Fill
    End With

     With Me.cboTipoClassVend
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDClassificazioneArticolo"
        .DisplayField = "ClassificazioneArticolo"
        .Sql = "SELECT * FROM ClassificazioneArticolo WHERE IDAzienda=" & TheApp.IDFirm & " ORDER BY ClassificazioneArticolo"
        .Fill
    End With
    
     With Me.cboCatLiq
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POCategoriaLiquidazione"
        .DisplayField = "CategoriaLiquidazione"
        .Sql = "SELECT * FROM RV_POCategoriaLiquidazione  ORDER BY CategoriaLiquidazione"
        .Fill
    End With
    
    With Me.cboVettoreConf
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .Sql = "SELECT * FROM Vettore  ORDER BY Vettore"
        .Fill
    End With
        
    With Me.cboLuogoPresaMerceConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT * FROM SitoPerAnagrafica  "
        .Sql = .Sql & " WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .Sql = .Sql & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With
        
    With Me.cboTipoConformeConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoOrdineConforme"
        .DisplayField = "TipoOrdineConforme"
        .Sql = "SELECT IDRV_POTipoOrdineConforme, TipoOrdineConforme FROM RV_POTipoOrdineConforme"
        .Sql = .Sql & " ORDER BY TipoOrdineConforme"
    End With
        
    
End Sub


Private Sub ACSCliente_ChangedElement()
    Me.cdCliente.Load Me.ACSCliente.IDAnagrafica
End Sub

Private Sub cboCategoriaFiscale_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub cboCategoriaMerceologica_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub cboDestinazione_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub


Private Sub cboNessunaForzatura_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub cboPrezzoMedioInLiq_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub cboTipoDocumento_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub cboTipoDocumentoCoop_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub cboTipoImportoLiq_Click()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub CDArticoloConferito_ChangeElement()

    If Me.CDArticoloConferito.KeyFieldID > 0 Then
        Me.txtDescrizioneArticoloConf.Text = Me.CDArticoloConferito.Description
        Me.txtDescrizioneArticoloConf.Locked = True
    Else
        Me.txtDescrizioneArticoloConf.Locked = False
    End If
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
    
End Sub

Private Sub CDArticoloVenduto_ChangeElement()
    If Me.CDArticoloVenduto.KeyFieldID > 0 Then
        Me.txtDescrizioneArticolo.Text = Me.CDArticoloVenduto.Description
        Me.txtDescrizioneArticolo.Locked = True
    Else
        Me.txtDescrizioneArticolo.Locked = False
    End If
    Me.GrigliaControlloVendite.SaveUserSettings
End Sub

Private Sub cdCliente_ChangeElement()

    Me.GrigliaControlloVendite.SaveUserSettings

    With Me.cboDestinazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .Sql = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica"
        .Sql = .Sql & " WHERE IDAnagrafica = " & Me.cdCliente.KeyFieldID
        .Sql = .Sql & " ORDER BY SitoPerAnagrafica"
    End With


    'fnGriglia
    

End Sub

Private Sub CDImballoVendita_ChangeElement()
    If Me.CDImballoVendita.KeyFieldID > 0 Then
        Me.txtDescrizioneImballo.Text = Me.CDImballoVendita.Description
        Me.txtDescrizioneImballo.Locked = True
    Else
        Me.txtDescrizioneImballo.Locked = False
    End If

    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub CDSocio_ChangeElement()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub CDSocioFatt_ChangeElement()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub cmdConferimento_Click()
On Error Resume Next
    If fnNotNullN(Me.GrigliaControlloVendite("RV_POIDTipoDocumentoCoop").Value) = 1 Then
        LabelLink2.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POCaricoMerceL"))
    Else
        LabelLink2.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POFattAcqL"))
    End If
    
    Me.LabelLink2.IDReturn = fnNotNullN(Me.GrigliaControlloVendite("IDRV_POCaricoMerceTesta").Value)
    Me.LabelLink2.RunApplication
End Sub

Private Sub cmdEseguiRicerca_Click()
On Error GoTo ERR_cmdEseguiRicerca_Click
    Screen.MousePointer = 11
    frmAttesa.Show
    DoEvents
    Me.Enabled = False
    fnGriglia
    Unload frmAttesa
    Screen.MousePointer = 0
    Me.Enabled = True
    DoEvents
Exit Sub
ERR_cmdEseguiRicerca_Click:
    Unload frmAttesa
    Screen.MousePointer = 0
End Sub

Private Sub cmdEspandiGriglia_Click()
    If FraStampa.Visible = False Then
        Me.FraStampa.Height = Me.GrigliaControlloVendite.Height + 100
        'Me.cmdEspandiGriglia.Visible = False
        'Me.cmdRiduciGriglia.Visible = True
        Me.cmdStampa.Enabled = True
        Me.ActivityBox.Height = Me.FraStampa.Height - 700
        FraStampa.Visible = True
    Else
        FraStampa.Visible = False
    
    End If
End Sub

Private Sub cmdLavorazione_Click()
On Error Resume Next
    Me.LabelLink3.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POAssegnazioneMerce"))
    Me.LabelLink3.IDReturn = fnNotNullN(Me.GrigliaControlloVendite("RV_POIDCaricoMerceRighe").Value)
    Me.LabelLink3.RunApplication
End Sub

Private Sub cmdPulisciFiltri_Click()
    Me.CDArticoloConferito.Load 0
    Me.CDArticoloVenduto.Load 0
    'Me.cdCliente.Load 0
    
    
    Me.ACSCliente.IDAnagrafica = 0
    Me.ACSCliente.Description = ""
    Me.ACSCliente.SecondDescription = ""
    
    
    Me.CDImballoVendita.Load 0
    Me.CDSocio.Load 0
    Me.CDSocioFatt.Load 0
    
    Me.txtADataConferimento.Value = 0
    Me.txtDaDataConferimento.Value = 0
    Me.txtADataConsegnaMerce.Value = 0
    Me.txtDaDataConsegnaMerce.Value = 0
    Me.txtADataVendita.Value = 0
    Me.txtDaDataVendita.Value = 0
    Me.txtANumeroDocumento.Text = ""
    Me.txtDaNumeroDocumento.Text = ""
    
    
    Me.txtDescrizioneArticolo.Text = ""
    Me.txtDescrizioneImballo.Text = ""
    Me.txtDescrizioneArticoloConf.Text = ""
    
    Me.txtLottoConferimento.Text = ""
    Me.txtLottoDiCampagna.Text = ""
    Me.txtLottoVendita.Text = ""
    
    Me.cboCategoriaFiscale.WriteOn 0
    Me.cboCategoriaMerceologica.WriteOn 0
    Me.cboConferimentoChiuso.WriteOn 0
    Me.cboDestinazione.WriteOn 0
    Me.cboTipoDocumento.ListIndex = -1
    Me.cboNessunaForzatura.WriteOn 0
    Me.cboNessunaForzatura.WriteOn 0
    Me.cboPrezzoMedioInLiq.WriteOn 0
    
'    Me.txtDaDataVendita.Value = Date
'    Me.txtADataVendita.Value = Date
    
    Me.txtAnnoProcesso.Value = 0
    Me.txtNumeroProcesso.Value = 0
    Me.txtDataProcesso.Value = 0
    
    Me.txtADataCompLiq.Value = 0
    Me.txtDaDataCompLiq.Value = 0
    
End Sub

Private Sub cmdRiduciGriglia_Click()
    Me.FraStampa.Height = 615
    Me.cmdEspandiGriglia.Visible = True
    Me.cmdRiduciGriglia.Visible = False
    Me.cmdStampa.Enabled = False
    Me.ActivityBox.Height = 1095
End Sub

Private Sub StampaDocumento()
On Error GoTo ERR_StampaDocumento
Dim IDReport As Long

Set oReport = New dmtReportLib.dmtReport
    'parametri di accesso al database SQL Server
    Set oReport.Connection = TheApp.Database.Connection
    oReport.Password = TheApp.Password
    oReport.User = TheApp.User

'Imposta l'idfiliale di appartenenza del documento da stampare
    oReport.BranchID = TheApp.Branch  'IDFiliale

'Imposta l'identificativo del tipo di documento
    IDTipoOggettoPrg = fncIDTipoOggettoPrg
    
    oReport.DocTypeID = IDTipoOggettoPrg
    
    
    oReport.Where = GET_SQL_PER_STAMPA
    
    
    Me.txtVista.Text = oReport.Where
    
    If Len(oReportsActivity.SelectedReportName) > 0 Then
        IDReport = fncTrovaReport(oReportsActivity.SelectedReportName, GET_TIPO_OGGETTO(App.EXEName))
    Else
        IDReport = fncTrovaReport(oReportsActivity.DefaultReportName, GET_TIPO_OGGETTO(App.EXEName))
    End If
    
    If IDReport > 0 Then
        fncImpostaDefaultReport IDReport, GET_TIPO_OGGETTO(App.EXEName)
        
        oReport.Preview 0, 0, 0
        PrimaElaborazione = True
    Else
        MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
    End If
Exit Sub
ERR_StampaDocumento:
    MsgBox Err.Description, vbCritical, "StampaDocumento"
End Sub

Private Sub cmdStampa_Click()
    
    StampaDocumento
End Sub



Private Sub cmdVendita_Click()
On Error Resume Next
    Select Case Me.GrigliaControlloVendite("IDTipoOggetto").Value
        Case 2
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_PODDTL"))
            Me.LabelLink1.IDReturn = Me.GrigliaControlloVendite("IDOggetto").Value
            Me.LabelLink1.RunApplication
        Case 114
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POFAL"))
            Me.LabelLink1.IDReturn = Me.GrigliaControlloVendite("IDOggetto").Value
            Me.LabelLink1.RunApplication
        Case 8
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POSNFL"))
            Me.LabelLink1.IDReturn = Me.GrigliaControlloVendite("IDOggetto").Value
            Me.LabelLink1.RunApplication
        Case 11
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_PONCL"))
            Me.LabelLink1.IDReturn = Me.GrigliaControlloVendite("IDOggetto").Value
            Me.LabelLink1.RunApplication
        Case 107
            LabelLink1.IDFunction = GET_FUNZIONE(GET_TIPO_OGGETTO("RV_POFIL"))
            Me.LabelLink1.IDReturn = Me.GrigliaControlloVendite("IDOggetto").Value
            Me.LabelLink1.RunApplication
    End Select
        
End Sub

Private Sub Form_Activate()
On Error GoTo ERR_Form_Activate
    
    If BLoading = False Then
        Me.txtADataVendita.Value = Date
        Me.txtDaDataVendita.Value = Date
        BLoading = True
        
        frmAttesa.Show
        DoEvents
        fnGrigliaPartenza
        Unload frmAttesa
        DoEvents
    End If
    
    
    Me.WindowState = 2

Exit Sub
ERR_Form_Activate:
    MsgBox Err.Description, vbCritical, "Form_Activate"
    Unload frmAttesa
    Me.Enabled = True
End Sub

Private Sub Form_Load()
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
    
    BLoading = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
  If Me.WindowState <> 1 Then
            
        If Me.Width > 19200 Then
            Me.Pic1.Width = Me.Width - 240
            Me.Picture2.Width = Me.Pic1.Width - 120
            Me.GrigliaControlloVendite.Width = Me.Picture2.Width - 120
            Me.Frame1.Width = Me.GrigliaControlloVendite.Width
        End If
        If Me.Height > 12000 Then
            Me.Pic1.Height = Me.Height - 720
            Me.Picture2.Height = Me.Pic1.Height - 120
            Me.GrigliaControlloVendite.Height = Me.Picture2.Height - 120 - Me.Frame1.Height
        End If
            

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
Private Sub GrigliaControlloVendite_Reposition(ByVal AllColumns As dmtgridctl.dgColumns)
    If Not ((Me.GrigliaControlloVendite.Recordset.EOF) And (Me.GrigliaControlloVendite.Recordset.BOF)) Then
        Me.cmdVendita.ToolTipText = fnNotNull(Me.GrigliaControlloVendite("Oggetto").Value)
        Me.cmdVendita.ToolTipText = Me.cmdVendita.ToolTipText & " n° " & fnNotNull(Me.GrigliaControlloVendite("NumeroDocumento").Value)
        Me.cmdVendita.ToolTipText = Me.cmdVendita.ToolTipText & " del " & fnNotNull(Me.GrigliaControlloVendite("DataDocumento").Value)
        Me.Caption = TheApp.FunctionName & " - [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
        Me.Caption = Me.Caption & " [" & Me.cmdVendita.ToolTipText & "]"
        
        If fnNotNullN(Me.GrigliaControlloVendite("RV_POIDAssegnazioneMerce").Value) > 0 Then
            Me.cmdLavorazione.Enabled = True
        Else
            Me.cmdLavorazione.Enabled = False
        End If
        
        If fnNotNullN(Me.GrigliaControlloVendite("RV_POIDCaricoMerceRighe").Value) > 0 Then
            Me.cmdConferimento.Enabled = True
            'If fnNotNullN(Me.GrigliaControlloVendite("RV_POIDTipoDocumentoCoop").Value) = 1 Then
            '    Me.cmdConferimento.ToolTipText = "Conferimento "
            '    Me.cmdConferimento.ToolTipText = Me.cmdConferimento.ToolTipText & " n° " & fnNotNullN(Me.GrigliaControlloVendite.AllColumns("RV_PONumeroConferimento").Value)
            '    Me.cmdConferimento.ToolTipText = Me.cmdConferimento.ToolTipText & " del " & fnNotNull(Me.GrigliaControlloVendite.AllColumns("RV_PODataConferimento").Value)
            'Else
            '    Me.cmdConferimento.ToolTipText = "Acquisto merce "
            '    Me.cmdConferimento.ToolTipText = Me.cmdConferimento.ToolTipText & " n° " & fnNotNullN(Me.GrigliaControlloVendite.AllColumns("RV_PONumeroConferimento").Value)
            '    Me.cmdConferimento.ToolTipText = Me.cmdConferimento.ToolTipText & " del " & fnNotNull(Me.GrigliaControlloVendite.AllColumns("RV_PODataConferimento").Value)
            
            'End If

        Else
            Me.cmdConferimento.Enabled = False
            Me.cmdConferimento.ToolTipText = ""
        End If
        
    End If
    
    
End Sub

Private Sub LabelLink1_AfterRunServerApplication(ByVal lIDResultKey As Long)
On Error GoTo ERR_LabelLink1_AfterRunServerApplication
Dim NumeroRiga As Long

NumeroRiga = Me.GrigliaControlloVendite.ListIndex - 1
Me.GrigliaControlloVendite.SaveUserSettings
Screen.MousePointer = 11
frmAttesa.Show
DoEvents
fnGriglia
Unload frmAttesa
Screen.MousePointer = 0
DoEvents
Me.GrigliaControlloVendite.Recordset.Move NumeroRiga
DoEvents


Me.WindowState = 2
Exit Sub
ERR_LabelLink1_AfterRunServerApplication:
    Unload frmAttesa
    Screen.MousePointer = 0
End Sub

Private Sub LabelLink2_BeforeRunServerApplication()
    If fnNotNullN(Me.GrigliaControlloVendite("IDTipoDocumentoCoop").Value) = 1 Then
        SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POCaricoMerceL", "IDRigaConferimento", fnNotNullN(Me.GrigliaControlloVendite.AllColumns("RV_POIDCaricoMerceRighe").Value)
    End If
    If fnNotNullN(Me.GrigliaControlloVendite("IDTipoDocumentoCoop").Value) = 2 Then
        SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POFattAcqL", "IDRigaConferimento", fnNotNullN(Me.GrigliaControlloVendite.AllColumns("RV_POIDCaricoMerceRighe").Value)
    End If
End Sub

Private Sub LabelLink3_BeforeRunServerApplication()
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POAssegnazioneMerce", "IDLavorazione", fnNotNullN(Me.GrigliaControlloVendite.AllColumns("RV_POIDAssegnazioneMerce").Value)
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POAssegnazioneMerce", "Tipo", 1
End Sub

Private Sub lblArticoloConferito_Click()
On Error GoTo ERR_lblDescrizioneArticolo_Click
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtOleDbLib.adoResultset
If Me.CDArticoloVenduto.KeyFieldID > 0 Then Exit Sub

    'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = CnDMT
    
    'La Caption della finestra di ricerca
    oSearch.Caption = "Articoli"
    
    
    oSearch.AddDisplayField "Articolo", "Articolo", 1
    oSearch.AddDisplayField "Codice", "CodiceArticolo", 1
        
    
    oSearch.Filters.Add "Articolo", Me.txtDescrizioneArticoloConf.Text
    
    oSearch.Start = Me.txtDescrizioneArticoloConf.Text

    sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
        
    oSearch.Sql = sSQL
    
    Set oRes = oSearch.Exec
    
    
    If Not oRes.EOF Then
        Me.CDArticoloConferito.Load fnNotNullN(oRes!IDArticolo)
        Me.txtDescrizioneArticoloConf.Text = fnNotNull(oRes!Articolo)
    End If
            

    
Set oRes = Nothing
Set oSearch = Nothing
Exit Sub
ERR_lblDescrizioneArticolo_Click:
    MsgBox Err.Description, vbCritical, "lblDescrizioneArticoloConf_Click"
End Sub


Private Sub lblDescrizioneArticolo_Click()
On Error GoTo ERR_lblDescrizioneArticolo_Click
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtOleDbLib.adoResultset
If Me.CDArticoloVenduto.KeyFieldID > 0 Then Exit Sub

    'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = CnDMT
    
    'La Caption della finestra di ricerca
    oSearch.Caption = "Articoli"
    
    
    oSearch.AddDisplayField "Articolo", "Articolo", 1
    oSearch.AddDisplayField "Codice", "CodiceArticolo", 1
        
    
    oSearch.Filters.Add "Articolo", Me.txtDescrizioneArticolo.Text
    
    oSearch.Start = Me.txtDescrizioneArticolo.Text

    sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
        
    oSearch.Sql = sSQL
    
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        Me.CDArticoloVenduto.Load fnNotNullN(oRes!IDArticolo)
        Me.txtDescrizioneArticolo.Text = fnNotNull(oRes!Articolo)
    End If
        
Set oRes = Nothing
Set oSearch = Nothing
Exit Sub
ERR_lblDescrizioneArticolo_Click:
    MsgBox Err.Description, vbCritical, "lblDescrizioneArticolo_Click"
End Sub


Private Sub lblDescrizioneImballo_Click()
On Error GoTo ERR_lblDescrizioneArticolo_Click
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtOleDbLib.adoResultset

If Me.CDImballoVendita.KeyFieldID > 0 Then Exit Sub

    'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = CnDMT
    
    'La Caption della finestra di ricerca
    oSearch.Caption = "Articoli"
    
    oSearch.AddDisplayField "Imballo", "Articolo", 1
    oSearch.AddDisplayField "Codice", "CodiceArticolo", 1
        
    oSearch.Filters.Add "Articolo", Me.txtDescrizioneImballo.Text
    
    oSearch.Start = Me.txtDescrizioneImballo.Text

    sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
    sSQL = sSQL & "FROM Articolo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
        
    oSearch.Sql = sSQL
    
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        Me.CDImballoVendita.Load fnNotNullN(oRes!IDArticolo)
        Me.txtDescrizioneImballo.Text = fnNotNull(oRes!Articolo)
    End If
    
Set oRes = Nothing
Set oSearch = Nothing

Exit Sub
ERR_lblDescrizioneArticolo_Click:
    MsgBox Err.Description, vbCritical, "lblDescrizioneArticolo_Click"

End Sub


Private Sub txtADataConferimento_GotFocus()
    FILTRO_A_DATA_LIQUIDAZIONE = txtADataConferimento.Value
End Sub
Private Sub txtADataConferimento_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    If (FILTRO_A_DATA_LIQUIDAZIONE = Me.txtADataConferimento.Value) Then Exit Sub
    FILTRO_A_DATA_LIQUIDAZIONE = Me.txtADataConferimento.Value
    'fnGriglia
End Sub
Private Sub txtADataConsegnaMerce_GotFocus()
    FILTRO_A_DATA_CONSEGNA_MERCE = txtADataConsegnaMerce.Value
End Sub

Private Sub txtADataConsegnaMerce_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    If (FILTRO_A_DATA_CONSEGNA_MERCE = Me.txtADataConsegnaMerce.Value) Then Exit Sub
    FILTRO_A_DATA_CONSEGNA_MERCE = Me.txtADataConsegnaMerce.Value
    'fnGriglia
End Sub

Private Sub txtADataVendita_GotFocus()
    FILTRO_A_DATA_VENDITA = txtADataVendita.Value
End Sub

Private Sub txtADataVendita_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    If (FILTRO_A_DATA_VENDITA = Me.txtADataVendita.Value) Then Exit Sub
    FILTRO_A_DATA_VENDITA = Me.txtADataVendita.Value
    'fnGriglia
End Sub

Private Sub txtDaDataConferimento_GotFocus()
    FILTRO_DA_DATA_LIQUIDAZIONE = txtDaDataConferimento.Value
End Sub

Private Sub txtDaDataConferimento_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings

    If (FILTRO_DA_DATA_LIQUIDAZIONE = Me.txtDaDataConferimento.Value) Then Exit Sub
    FILTRO_DA_DATA_LIQUIDAZIONE = Me.txtDaDataConferimento.Value
    
    'fnGriglia
End Sub

Private Sub txtDaDataConsegnaMerce_GotFocus()
    FILTRO_DA_DATA_CONSEGNA_MERCE = Me.txtDaDataConsegnaMerce.Value
End Sub

Private Sub txtDaDataConsegnaMerce_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    
    If (FILTRO_DA_DATA_CONSEGNA_MERCE = Me.txtDaDataConsegnaMerce.Value) Then Exit Sub
    FILTRO_DA_DATA_CONSEGNA_MERCE = Me.txtDaDataConsegnaMerce.Value
    
    'fnGriglia
End Sub

Private Sub txtDaDataVendita_GotFocus()
    FILTRO_DA_DATA_VENDITA = txtDaDataVendita.Value
End Sub

Private Sub txtDaDataVendita_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    
    If FILTRO_DA_DATA_VENDITA = Me.txtDaDataVendita.Value Then Exit Sub
    FILTRO_DA_DATA_VENDITA = Me.txtDaDataVendita.Value
    
    'fnGriglia
End Sub

Private Sub txtLottoConferimento_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub txtLottoDiCampagna_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

Private Sub txtLottoVendita_LostFocus()
    Me.GrigliaControlloVendite.SaveUserSettings
    'fnGriglia
End Sub

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
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
End Sub
Private Sub fnGriglia()
On Error GoTo ERR_fnGrigliaAssegnazione
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    sSQL = "SELECT * FROM RV_POIEControlloVendite "
    sSQL = sSQL & "WHERE "
    sSQL = sSQL & GET_SQL
    sSQL = sSQL & " ORDER BY DataDocumento DESC, NumeroDocumento DESC"
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        If Not (rsGriglia Is Nothing) Then
            rsGriglia.Close
            Set rsGriglia = Nothing
        End If
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
        
        With Me.GrigliaControlloVendite
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            Set .PaintNotifyObj = gPaintNotify
            '.LoadUserSettings
                .ColumnsHeader.Add "IDMovimento", "IDMovimento", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "RV_POIDCaricoMerceRighe", "RV_POIDCaricoMerceRighe", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "RV_POIDAssegnazioneMerce", "RV_POIDAssegnazioneMerce", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "RV_POIDProcessoIVGamma", "RV_POIDProcessoIVGamma", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDAzienda", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDFunzione", "IDFunzione", dgInteger, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDProcessoPerFunzione", "IDProcessoPerFunzione", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDSezionale", "IDSezionale", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "DataMovimento", "Data Movimento", dgDate, False, 1500, dgAlignleft
                .ColumnsHeader.Add "DataDocumento", "Data documento", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "NumeroDocumento", "N° documento", dgNumeric, True, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_PODataCompetenzaLiq", "Data competenza liq.", dgDate, False, 1500, dgAlignleft
                .ColumnsHeader.Add "Oggetto", "Oggetto", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "TipoOggettoDescrizione", "Tipo documento vendita", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "AnagraficaCliente", "Cliente", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "NomeCliente", "Nome cliente", dgchar, False, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_POIDSitoPerAnagrafica", "RV_POIDSitoPerAnagrafica", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione diversa", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "IDArticoloVenduto", "IDArticoloVenuto", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "CodiceArticoloVenduto", "Codice Articolo", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "IDCategoriaMerceologica", "IDCategoriaMerceologica", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "CategoriaMerceologica", "Categoria merceologica", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "IDCategoriaFiscale", "IDCategoriaFiscale", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "CategoriaFiscale", "Categoria fiscale", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "DescrizioneArticoloVenduto", "Articolo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "IDUnitaDiMisuraVenduto", "IDUnitaDiMisura", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "UnitaDiMisuraVenduto", "U.M.", dgchar, True, 1000, dgAlignleft

                 Set cl = .ColumnsHeader.Add("QuantitaTotale", "Quantità", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PrezzoUnitario", "Imp. Uni.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("ScontoPerc1", "% Sc.1", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("ScontoPerc2", "% Sc.2", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PrezzoScontato", "Imp. uni. scont.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("Importo", "Importo riga", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                Set cl = .ColumnsHeader.Add("RV_PONumeroColli", "Colli", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("RV_POPesoLordo", "Peso lordo", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POTara", "Tara", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POPesoNetto", "Peso netto", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POQuantitaPezzi", "Numero Pezzi", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POCodiceLottoVendita", "Lotto di vendita", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "RV_POCodiceLotto", "Lotto di entrata", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "RV_POCodiceLottoCampagna", "Lotto di produzione", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDImballo", "RV_POIDImballo", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "CodiceArticoloImballo", "Codice Imballo", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "DescrizioneArticoloImballo", "Descrizione imballo", dgchar, True, 2000, dgAlignleft
                 Set cl = .ColumnsHeader.Add("RV_POImportoUnitarioImballo", "Imp. Uni. Imb.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 
                 Set cl = .ColumnsHeader.Add("RV_POQuantitaLiquidazione", "Q.ta liq.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POMerceInclusaImballo", "Merce incluso imballo", dgBoolean, True, 1000, dgAlignleft
                 Set cl = .ColumnsHeader.Add("RV_POImportoInclusoImballo", "Var. Liq.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POPrezzoMedioInLiq", "Prezzo medio", dgBoolean, True, 1000, dgAlignleft
                
                Set cl = .ColumnsHeader.Add("RV_POImportoLiquidazione", "Imp. Uni. Liq.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("RV_POImportoLiqDoc", "Imp. Uni. Liq. doc.", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POIDTipoImportoVenditaLiq", "RV_POIDTipoImportoVenditaLiq", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "TipoImportoVenditaLiq", "Tipo forzatura importo in Liq.", dgchar, True, 1000, dgAlignleft

                .ColumnsHeader.Add "IDTipoDocumentoCoop", "IDTipoDocumentoCoop", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "TipoDocumentoCoop", "Documento di consegna merce", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "DataConsegnaMerce", "Data consegna merce", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_PODataConferimento", "Data Conferimento", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_PONumeroConferimento", "N° Conferimento", dgNumeric, True, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_POIDAnagraficaSocio", "RV_POIDAnagraficaSocio", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "AnagraficaSocio", "Socio/Fornitore", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "NomeSocio", "Nome", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDAnagraficaFatturazione", "RV_POIDAnagraficaFatturazione", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "AnagraficaFatturazione", "Socio/Fornitore per F.C.S.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "NomeAnagraficaFatturazione", "Nome per F.C.S.", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "IDArticoloConferito", "IDArticoloConferito", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "CodiceArticoloConferito", "Codice Articolo Conf.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "ArticoloConferito", "Articolo Conf.", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "RV_POIDTipoLavorazione", "RV_POIDTipoLavorazione", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione merce", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "RV_POIDTipoCategoria", "RV_POIDTipoCategoria", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "TipoCategoria", "Tipo categoria", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "RV_POIDCalibro", "RV_POIDCalibro", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "Calibro", "Calibro", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "RV_POIDTipoLavorazioneConf", "RV_POIDTipoLavorazioneConf", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "TipoLavorazioneConferimento", "Tipo lavorazione conf.", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "RV_POIDPedana", "RV_POIDPedana", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "RV_POCodicePedana", "Codice pedana", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "CodiceTipoPedana", "Codice tipo pedana", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "TipoPedana", "Tipo pedana", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "IDClienteOrdine", "IDClienteOrdine", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "DataOrdine", "Data ordine", dgDate, True, 1700, dgAlignleft
                .ColumnsHeader.Add "NumeroOrdine", "Numero ordine", dgNumeric, True, 1700, dgAlignRight
                .ColumnsHeader.Add "IDTipoOggettoStatistica", "IDTipoOggettoStatistica", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "RV_POPrezzoMedioConf", "Prezzo medio conf.", dgBoolean, False, 1500, dgAligncenter
                .ColumnsHeader.Add "RV_PODataLavorazione", "Data lavorazione", dgDate, False, 1500, dgAlignleft
                .ColumnsHeader.Add "DataProcesso", "Data processo", dgDate, False, 1500, dgAlignleft
                .ColumnsHeader.Add "AnnoProcesso", "Anno processo", dgNumeric, False, 1500, dgAlignRight
                .ColumnsHeader.Add "NumeroProcesso", "Numero processo", dgNumeric, False, 1500, dgAlignRight
                Set cl = .ColumnsHeader.Add("ColliConferito", "Colli Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("PesoLordoConferito", "Peso lordo Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("TaraConferito", "Tara Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PesoNettoConferito", "Peso netto conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PezziConferito", "Numero Pezzi conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                 Set cl = .ColumnsHeader.Add("RV_POQuantitaMovimentata", "Q.tà Mov. Vend", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                 Set cl = .ColumnsHeader.Add("Qta_UMConferito", "Q.tà Mov. Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                .ColumnsHeader.Add "RV_POTipoRigaCollegata", "RV_POTipoRigaCollegata", dgNumeric, False, 500, dgAlignleft
                 Set cl = .ColumnsHeader.Add("RV_POPrezzoMerceNetta", "Prezzo netto merce", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POVariazionePrezzoImballo", "Var. Prezzo imballo", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                 Set cl = .ColumnsHeader.Add("RV_POQuantitaLiqPerPrezzoMedio", "Q.tà per prezzo medio", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POVariazionePrezzoManuale", "Var. Liq. Manuale", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POImportoRigaCommissioni", "Imp. commissioni", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                
                .ColumnsHeader.Add "RV_POIDIvaImballo", "RV_POIDIvaImballo", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "DescrizioneIvaImballo", "I.V.A. Imballo", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "CodiceIvaImballo", "Codice I.V.A. Imballo", dgchar, False, 1000, dgAlignleft
                 Set cl = .ColumnsHeader.Add("AliquotaIvaImballo", "Aliquota I.V.A. Imballo", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POAnnotazioniAggiuntiveLav", "Note aggiuntive", dgchar, False, 2500, dgAlignleft
                .ColumnsHeader.Add "RV_PONotaRigaOrdRaggr", "Raggr. ordine", dgchar, False, 2500, dgAlignleft
                .ColumnsHeader.Add "RV_PODataOrdineCliente", "Data ordine cliente", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "RV_PONumeroOrdineCliente", "N° ordine cliente", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "IDClassificazioneArticoloConf", "IDClassificazioneArticoloConf", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "ClassificazioneArticoloConf", "Class. conf.", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDClassificazioneArticoloVend", "IDClassificazioneArticoloVend", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "ClassificazioneArticoloVend", "Class. vend.", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDCategoriaLiquidazioneVend", "IDCategoriaLiquidazioneVend", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "CategoriaLiquidazioneVend", "Categoria liq.", dgchar, False, 2500, dgAlignleft
                
                
                'IMBALLO PRIMARIO
                .ColumnsHeader.Add "RV_POIDImballoPrim", "RV_POIDImballoPrim", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "RV_POCodiceImballoPrim", "Codice imb. prim.", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "RV_PODescrizioneImballoPrim", "Descr. imb. prim.", dgchar, False, 2500, dgAlignleft
                 Set cl = .ColumnsHeader.Add("RV_PONumeroConfezioniPerImballo", "N° confez.", dgDouble, False, 2000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POTaraConfezioneImballo", "Tara confez.", dgDouble, False, 2000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POCostoConfezioneImballo", "Costo confez.", dgDouble, False, 2000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POCostoConfezioneImballoLiq", "Costo confez. liq.", dgDouble, False, 2000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POCostoKitLiq", "Costo Kit. liq.", dgDouble, False, 2000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POIDLottoCampagnaLavorazione", "IDLottoCampagnaLavorazione", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDSezionaleConf", "IDSezionaleConf", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "SezionaleConf", "Sezionale conferimento", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDVettoreConf", "IDVettoreConf", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "VettoreConf", "Vettore conferimento", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDLuogoPresaMerce", "IDLuogoPresaMerce", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "LuogoPresaMerceConf", "Sede stoccaggio merce", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDRV_POTipoOrdineConforme", "IDRV_POTipoOrdineConforme", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "TipoOrdineConformeConf", "Conformità", dgchar, False, 2000, dgAlignleft
                
            Set .Recordset = rsGriglia
            .LoadUserSettings
            .Refresh
      
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati"
End Sub

Private Function fncTrovaReport(NomeReport As String, IDTipoOggetto As Long) As Long
On Error GoTo ERR_fncTrovaReport
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE ((ReportTipoOggetto=" & fnNormString(NomeReport) & ") AND (IDTipoOggetto=" & IDTipoOggetto & "))"

Set rs = CnDMT.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaReport = fnNotNullN(rs!IDReportTipoOggetto)
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
Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggetto & " AND IDFiliale = " & VarIDFiliale
    
    CnDMT.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
Private Function fncIDTipoOggettoPrg() As Long
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto, Gestore.Gestore"
    sSQL = sSQL & " FROM Gestore INNER JOIN TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore"
    sSQL = sSQL & " WHERE (((Gestore.Gestore)=" & fnNormString(App.EXEName) & "))"
    
    Set rs = CnDMT.OpenResultset(sSQL)
        
    If rs.EOF = False Then
        fncIDTipoOggettoPrg = rs!IDTipoOggetto
    Else
        fncIDTipoOggettoPrg = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function GET_SQL() As String
Dim sSQL_GENERALE As String
Dim sSQL_TIPO_DOCUMENTO As String
Dim sSQL As String

Dim I As Integer

sSQL_GENERALE = "IDAzienda=" & TheApp.IDFirm
sSQL_GENERALE = sSQL_GENERALE & " AND RV_POTipoRiga=1"

'sSQL_GENERALE = sSQL_GENERALE & " AND IDTipoDocumentoCoop=1"

If Me.ACSCliente.IDAnagrafica > 0 Then
    'sSQL_GENERALE = sSQL_GENERALE & " AND IDAnagraficaCliente=" & Me.cdCliente.KeyFieldID
    sSQL_GENERALE = sSQL_GENERALE & " AND IDAnagraficaCliente=" & Me.ACSCliente.IDAnagrafica
End If
If Me.CDClienteOrdine.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDClienteOrdine=" & Me.CDClienteOrdine.KeyFieldID
End If

If Me.CDArticoloVenduto.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDArticoloVenduto=" & Me.CDArticoloVenduto.KeyFieldID
End If
If (Me.CDArticoloVenduto.KeyFieldID = 0) And (Len(Me.txtDescrizioneArticolo.Text) > 0) Then
    sSQL_GENERALE = sSQL_GENERALE & " AND DescrizioneArticoloVenduto LIKE " & fnNormString(Me.txtDescrizioneArticolo.Text & "%")
End If
If Me.CDImballoVendita.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDImballo=" & Me.CDImballoVendita.KeyFieldID
End If
If (Me.CDImballoVendita.KeyFieldID = 0) And (Len(Me.txtDescrizioneImballo.Text) > 0) Then
    sSQL_GENERALE = sSQL_GENERALE & " AND DescrizioneArticoloImballo LIKE " & fnNormString(Me.txtDescrizioneImballo.Text & "%")
End If
If Me.CDArticoloConferito.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDArticoloConferito=" & Me.CDArticoloConferito.KeyFieldID
End If
If (Me.CDArticoloConferito.KeyFieldID = 0) And (Len(Me.txtDescrizioneArticoloConf.Text) > 0) Then
    sSQL_GENERALE = sSQL_GENERALE & " AND ArticoloConferito LIKE " & fnNormString(Me.txtDescrizioneArticoloConf.Text & "%")
End If
If Me.CDSocio.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
End If
If Len(Trim(Me.txtLottoDiCampagna.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodiceLottoCampagna LIKE " & fnNormString("%" & Me.txtLottoDiCampagna.Text & "%")
End If
If Len(Trim(Me.txtLottoConferimento.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodiceLotto LIKE " & fnNormString("%" & Me.txtLottoConferimento.Text & "%")
End If
If Len(Trim(Me.txtLottoVendita.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodiceLottoVendita LIKE " & fnNormString("%" & Me.txtLottoVendita.Text & "%")
End If
If Me.cboCategoriaMerceologica.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDCategoriaMerceologica = " & Me.cboCategoriaMerceologica.CurrentID
End If
If Me.cboTipoDocumentoCoop.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoDocumentoCoop = " & Me.cboTipoDocumentoCoop.CurrentID
End If
If Me.cboCategoriaFiscale.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDCategoriaFiscale = " & Me.cboCategoriaFiscale.CurrentID
End If
If Me.cboDestinazione.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDSitoPerAnagrafica = " & Me.cboDestinazione.CurrentID
End If
If Me.CDSocioFatt.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDAnagraficaFatturazione = " & Me.CDSocioFatt.KeyFieldID
End If
If Me.cboTipoImportoLiq.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoImportoVenditaLiq = " & Me.cboTipoImportoLiq.CurrentID
End If
If Me.cboSezionale.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDSezionale = " & Me.cboSezionale.CurrentID
End If
If Me.cboTipoLavConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoLavorazioneConf = " & Me.cboTipoLavConf.CurrentID
End If
If Me.cboTipoLavorazione.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoLavorazione = " & Me.cboTipoLavorazione.CurrentID
End If
If Me.cboCalibro.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDCalibro = " & Me.cboCalibro.CurrentID
End If
If Me.cboPrezzoMedioConf.CurrentID > 0 Then
    If Me.cboPrezzoMedioConf.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioConf=" & fnNormBoolean(0)
    End If
    If Me.cboPrezzoMedioConf.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioConf=" & fnNormBoolean(1)
    End If
End If
If Me.cboRiscontroPeso.CurrentID > 0 Then
    If Me.cboRiscontroPeso.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PORigaRiscontroPeso=" & 0
    End If
    If Me.cboRiscontroPeso.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PORigaRiscontroPeso=" & 1
    End If
End If
If Me.cboTipoCategoria.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoCategoria = " & Me.cboTipoCategoria.CurrentID
End If
If Me.CDTipoPedana.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoPedana = " & Me.CDTipoPedana.KeyFieldID
End If
If Len(Trim(Me.txtCodicePedana.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodicePedana LIKE " & fnNormString("%" & Me.txtCodicePedana.Text & "%")
End If
If Me.txtNumeroOrdine.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND NumeroOrdine = " & Me.txtNumeroOrdine.Value
End If

If Len(Me.txtRaggrRigaOrdine.Text) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_PONotaRigaOrdRaggr LIKE " & fnNormString("%" & Me.txtRaggrRigaOrdine.Text & "%")
End If

If Len(Me.txtNoteAgg.Text) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POAnnotazioniAggiuntiveLav LIKE " & fnNormString("%" & Me.txtNoteAgg.Text & "%")
End If
If Me.cboTipoClassConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDClassificazioneArticoloConf = " & Me.cboTipoClassConf.CurrentID
End If
If Me.cboTipoClassVend.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDClassificazioneArticoloVend = " & Me.cboTipoClassVend.CurrentID
End If
If Me.CDImballoPrimario.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDImballoPrim = " & Me.CDImballoPrimario.KeyFieldID
End If
If Me.cboCatLiq.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDCategoriaLiquidazioneVend = " & Me.cboCatLiq.CurrentID
End If
If Me.txtAnnoProcesso.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND AnnoProcesso = " & Me.txtAnnoProcesso.Value
End If
If Me.txtNumeroProcesso.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND NumeroProcesso = " & Me.txtNumeroProcesso.Value
End If
If Me.txtDataProcesso.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND DataProcesso = " & fnNormDate(Me.txtDataProcesso.Text)
End If

If Me.cboSezionaleConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDSezionaleConf = " & Me.cboSezionaleConf.CurrentID
End If
If Me.cboVettoreConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDSezionaleConf = " & Me.cboVettoreConf.CurrentID
End If
If Me.cboTipoConformeConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDRV_POTipoOrdineConforme = " & Me.cboTipoConformeConf.CurrentID
End If
If Me.cboLuogoPresaMerceConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDLuogoPresaMerce = " & Me.cboLuogoPresaMerceConf.CurrentID
End If




If Me.cboNessunaForzatura.CurrentID > 0 Then
    If Me.cboNessunaForzatura.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoImportoVenditaLiq > 0"
    End If
    If Me.cboNessunaForzatura.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoImportoVenditaLiq = 0"
    End If
End If
If Me.cboPrezzoMedioInLiq.CurrentID > 0 Then
    If Me.cboPrezzoMedioInLiq.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioInLiq = 0"
    End If
    If Me.cboPrezzoMedioInLiq.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioInLiq = 1"
    End If
End If



If Me.cboConferimentoChiuso.CurrentID > 0 Then
    If Me.cboConferimentoChiuso.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND ConferimentoChiuso=" & fnNormBoolean(0)
    End If
    If Me.cboConferimentoChiuso.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND ConferimentoChiuso=" & fnNormBoolean(1)
    End If
End If
If Me.txtDaDataVendita.Value > 0 Then
    If Me.txtADataVendita.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento=" & fnNormDate(Me.txtDaDataVendita.Text)
    Else
        If Me.txtADataVendita.Value <= Me.txtDaDataVendita.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento=" & fnNormDate(Me.txtDaDataVendita.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento>=" & fnNormDate(Me.txtDaDataVendita.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento<=" & fnNormDate(Me.txtADataVendita.Text)
        End If
    End If
End If
If Me.txtDaNumeroDocumento.Value > 0 Then
    If Me.txtANumeroDocumento.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento=" & fnNormNumber(Me.txtDaNumeroDocumento.Text)
    Else
        If Me.txtANumeroDocumento.Value <= Me.txtDaNumeroDocumento.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento=" & fnNormNumber(Me.txtDaNumeroDocumento.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento>=" & fnNormNumber(Me.txtDaNumeroDocumento.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento<=" & fnNormNumber(Me.txtANumeroDocumento.Text)
        End If
    End If
End If
If Me.txtDaDataConferimento.Value > 0 Then
    If Me.txtADataConferimento.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento=" & fnNormDate(Me.txtDaDataConferimento.Text)
    Else
        If Me.txtADataConferimento.Value <= Me.txtDaDataConferimento.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento=" & fnNormDate(Me.txtDaDataConferimento.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento>=" & fnNormDate(Me.txtDaDataConferimento.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento<=" & fnNormDate(Me.txtADataConferimento.Text)
        End If
    End If
End If
If Me.txtDaDataConsegnaMerce.Value > 0 Then
    If Me.txtADataConsegnaMerce.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce=" & fnNormDate(Me.txtDaDataConsegnaMerce.Text)
    Else
        If Me.txtADataConsegnaMerce.Value <= Me.txtDaDataConsegnaMerce.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce=" & fnNormDate(Me.txtDaDataConsegnaMerce.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce>=" & fnNormDate(Me.txtDaDataConsegnaMerce.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce<=" & fnNormDate(Me.txtADataConsegnaMerce.Text)
        End If
    End If
End If
If Me.txtDaLavMerce.Value > 0 Then
    If Me.txtALavMerce.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione=" & fnNormDate(Me.txtDaLavMerce.Text)
    Else
        If Me.txtALavMerce.Value <= Me.txtDaLavMerce.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione=" & fnNormDate(Me.txtDaLavMerce.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione>=" & fnNormDate(Me.txtDaLavMerce.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione<=" & fnNormDate(Me.txtALavMerce.Text)
        End If
    End If
End If
If Me.txtDaDataArrivoMerce.Value > 0 Then
    If Me.txtADataArrivoMerce.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce=" & fnNormDate(Me.txtDaDataArrivoMerce.Text)
    Else
        If Me.txtADataArrivoMerce.Value <= Me.txtDaDataArrivoMerce.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce=" & fnNormDate(Me.txtDaDataArrivoMerce.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce>=" & fnNormDate(Me.txtDaDataArrivoMerce.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce<=" & fnNormDate(Me.txtADataArrivoMerce.Text)
        End If
    End If
End If
If Me.txtDaDataOrdine.Value > 0 Then
    If Me.txtADataOrdine.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine=" & fnNormDate(Me.txtDaDataOrdine.Text)
    Else
        If Me.txtADataOrdine.Value <= Me.txtDaDataOrdine.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine=" & fnNormDate(Me.txtDaDataOrdine.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine>=" & fnNormDate(Me.txtDaDataOrdine.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine<=" & fnNormDate(Me.txtADataOrdine.Text)
        End If
    End If
End If

If Me.txtDaDataOrdineCli.Value > 0 Then
    If Me.txtADataOrdineCli.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente=" & fnNormDate(Me.txtDaDataOrdineCli.Text)
    Else
        If Me.txtADataOrdineCli.Value <= Me.txtDaDataOrdineCli.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente=" & fnNormDate(Me.txtDaDataOrdineCli.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente>=" & fnNormDate(Me.txtDaDataOrdineCli.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente<=" & fnNormDate(Me.txtADataOrdineCli.Text)
        End If
    End If
End If
If Me.txtDaDataCompLiq.Value > 0 Then
    If Me.txtADataCompLiq.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq=" & fnNormDate(Me.txtDaDataCompLiq.Text)
    Else
        If Me.txtADataCompLiq.Value <= Me.txtDaDataCompLiq.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq=" & fnNormDate(Me.txtDaDataCompLiq.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq>=" & fnNormDate(Me.txtDaDataCompLiq.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq<=" & fnNormDate(Me.txtADataCompLiq.Text)
        End If
    End If
End If

If Me.cboTipoDocumento.ListIndex > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDTipoOggetto=" & Me.cboTipoDocumento.ItemData(Me.cboTipoDocumento.ListIndex)
    GET_SQL = sSQL_GENERALE
    Exit Function
Else
    sSQL_TIPO_DOCUMENTO = ""
    sSQL = ""
    sSQL_GENERALE = "(" & sSQL_GENERALE
    For I = 1 To Me.cboTipoDocumento.ListCount - 1
        sSQL_TIPO_DOCUMENTO = " AND IDTipoOggetto=" & Me.cboTipoDocumento.ItemData(I) & ")"
        sSQL = sSQL & sSQL_GENERALE & sSQL_TIPO_DOCUMENTO
         If I < Me.cboTipoDocumento.ListCount - 1 Then
            sSQL = sSQL & " OR "
        End If
    Next
End If

GET_SQL = sSQL
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
Private Function GET_SQL_PER_STAMPA() As String
Dim sSQL_GENERALE As String
Dim sSQL_TIPO_DOCUMENTO As String
Dim sSQL As String

Dim I As Integer

sSQL_GENERALE = "IDAzienda=" & TheApp.IDFirm
sSQL_GENERALE = sSQL_GENERALE & " AND RV_POTipoRiga=1"
'sSQL_GENERALE = sSQL_GENERALE & " AND IDTipoDocumentoCoop=1"

If Me.ACSCliente.IDAnagrafica > 0 Then
    'sSQL_GENERALE = sSQL_GENERALE & " AND IDAnagraficaCliente=" & Me.cdCliente.KeyFieldID
    sSQL_GENERALE = sSQL_GENERALE & " AND IDAnagraficaCliente=" & Me.ACSCliente.IDAnagrafica
End If
If Me.CDClienteOrdine.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDClienteOrdine=" & Me.CDClienteOrdine.KeyFieldID
End If
If Me.CDArticoloVenduto.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDArticoloVenduto=" & Me.CDArticoloVenduto.KeyFieldID
End If
If (Me.CDArticoloVenduto.KeyFieldID = 0) And (Len(Me.txtDescrizioneArticolo.Text) > 0) Then
    sSQL_GENERALE = sSQL_GENERALE & " AND DescrizioneArticoloVenduto LIKE " & fnNormString(Me.txtDescrizioneArticolo.Text & "%")
End If
If Me.CDImballoVendita.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDImballo=" & Me.CDImballoVendita.KeyFieldID
End If
If (Me.CDImballoVendita.KeyFieldID = 0) And (Len(Me.txtDescrizioneImballo.Text) > 0) Then
    sSQL_GENERALE = sSQL_GENERALE & " AND DescrizioneArticoloImballo LIKE " & fnNormString(Me.txtDescrizioneImballo.Text & "%")
End If
If Me.CDArticoloConferito.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDArticoloConferito=" & Me.CDArticoloConferito.KeyFieldID
End If
If (Me.CDArticoloConferito.KeyFieldID = 0) And (Len(Me.txtDescrizioneArticoloConf.Text) > 0) Then
    sSQL_GENERALE = sSQL_GENERALE & " AND ArticoloConferito LIKE " & fnNormString(Me.txtDescrizioneArticoloConf.Text & "%")
End If
If Me.CDSocio.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDAnagraficaSocio=" & Me.CDSocio.KeyFieldID
End If
If Len(Trim(Me.txtLottoDiCampagna.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodiceLottoCampagna LIKE " & fnNormString("%" & Me.txtLottoDiCampagna.Text & "%")
End If
If Len(Trim(Me.txtLottoConferimento.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodiceLotto LIKE " & fnNormString("%" & Me.txtLottoConferimento.Text & "%")
End If
If Len(Trim(Me.txtLottoVendita.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodiceLottoVendita LIKE " & fnNormString("%" & Me.txtLottoVendita.Text & "%")
End If
If Me.cboCategoriaMerceologica.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDCategoriaMerceologica =  " & Me.cboCategoriaMerceologica.CurrentID
End If
If Me.cboTipoDocumentoCoop.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoDocumentoCoop = " & Me.cboTipoDocumentoCoop.CurrentID
End If
If Me.cboCategoriaFiscale.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDCategoriaFiscale = " & Me.cboCategoriaFiscale.CurrentID
End If
If Me.cboDestinazione.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDSitoPerAnagrafica = " & Me.cboDestinazione.CurrentID
End If
If Me.CDSocioFatt.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDAnagraficaFatturazione = " & Me.CDSocioFatt.KeyFieldID
End If
If Me.cboTipoImportoLiq.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoImportoVenditaLiq = " & Me.cboTipoImportoLiq.CurrentID
End If
If Me.cboNessunaForzatura.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoImportoVenditaLiq = 0"
End If
If Me.cboSezionale.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDSezionale = " & Me.cboSezionale.CurrentID
End If
If Me.cboTipoLavConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoLavorazioneConf = " & Me.cboTipoLavConf.CurrentID
End If
If Me.cboTipoLavorazione.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoLavorazione = " & Me.cboTipoLavorazione.CurrentID
End If
If Me.cboCalibro.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDCalibro = " & Me.cboCalibro.CurrentID
End If
If Me.cboPrezzoMedioConf.CurrentID > 0 Then
    If Me.cboPrezzoMedioConf.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioConf=" & fnNormBoolean(0)
    End If
    If Me.cboPrezzoMedioConf.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioConf=" & fnNormBoolean(1)
    End If
End If
If Me.cboTipoCategoria.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoCategoria = " & Me.cboTipoCategoria.CurrentID
End If

If Me.CDTipoPedana.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDTipoPedana = " & Me.CDTipoPedana.KeyFieldID
End If
If Len(Trim(Me.txtCodicePedana.Text)) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POCodicePedana LIKE " & fnNormString("%" & Me.txtCodicePedana.Text & "%")
End If

If Len(Me.txtRaggrRigaOrdine.Text) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_PONotaRigaOrdRaggr LIKE " & fnNormString("%" & Me.txtRaggrRigaOrdine.Text & "%")
End If

If Len(Me.txtNoteAgg.Text) > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POAnnotazioniAggiuntiveLav LIKE " & fnNormString("%" & Me.txtNoteAgg.Text & "%")
End If

If Me.cboTipoClassConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDClassificazioneArticoloConf = " & Me.cboTipoClassConf.CurrentID
End If
If Me.cboTipoClassVend.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDClassificazioneArticoloVend = " & Me.cboTipoClassVend.CurrentID
End If
If Me.CDImballoPrimario.KeyFieldID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND RV_POIDImballoPrim = " & Me.CDImballoPrimario.KeyFieldID
End If
If Me.txtNumeroOrdine.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND NumeroOrdine = " & Me.txtNumeroOrdine.Value
End If
If Me.cboCatLiq.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDCategoriaLiquidazioneVend = " & Me.cboCatLiq.CurrentID
End If
If Me.txtAnnoProcesso.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND AnnoProcesso = " & Me.txtAnnoProcesso.Value
End If
If Me.txtNumeroProcesso.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND NumeroProcesso = " & Me.txtNumeroProcesso.Value
End If
If Me.txtDataProcesso.Value > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND DataProcesso = " & fnNormDateString(Me.txtDataProcesso.Text)
End If
If Me.cboSezionaleConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDSezionaleConf = " & Me.cboSezionaleConf.CurrentID
End If
If Me.cboVettoreConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDSezionaleConf = " & Me.cboVettoreConf.CurrentID
End If
If Me.cboTipoConformeConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDRV_POTipoOrdineConforme = " & Me.cboTipoConformeConf.CurrentID
End If
If Me.cboLuogoPresaMerceConf.CurrentID > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDLuogoPresaMerce = " & Me.cboLuogoPresaMerceConf.CurrentID
End If

If Me.cboPrezzoMedioInLiq.CurrentID > 0 Then
    If Me.cboPrezzoMedioInLiq.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioInLiq = 0"
    End If
    If Me.cboPrezzoMedioInLiq.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_POPrezzoMedioInLiq = 1"
    End If
End If
If Me.cboConferimentoChiuso.CurrentID > 0 Then
    If Me.cboConferimentoChiuso.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND ConferimentoChiuso=" & fnNormBoolean(0)
    End If
    If Me.cboConferimentoChiuso.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND ConferimentoChiuso=" & fnNormBoolean(1)
    End If
End If
If Me.cboRiscontroPeso.CurrentID > 0 Then
    If Me.cboRiscontroPeso.CurrentID = 2 Then 'NO
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PORigaRiscontroPeso = 0"
    End If
    If Me.cboRiscontroPeso.CurrentID = 1 Then 'SI
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PORigaRiscontroPeso = 1"
    End If
End If
If Me.txtDaDataVendita.Value > 0 Then
    If Me.txtADataVendita.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento=" & fnNormDateString(Me.txtDaDataVendita.Text)
    Else
        If Me.txtADataVendita.Value <= Me.txtDaDataVendita.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento=" & fnNormDateString(Me.txtDaDataVendita.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento>=" & fnNormDateString(Me.txtDaDataVendita.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND DataDocumento<=" & fnNormDateString(Me.txtADataVendita.Text)
        End If
    End If
End If

If Me.txtDaNumeroDocumento.Value > 0 Then
    If Me.txtANumeroDocumento.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento=" & fnNormString(Me.txtDaNumeroDocumento.Text)
    Else
        If Me.txtANumeroDocumento.Value <= Me.txtDaNumeroDocumento.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento=" & fnNormString(Me.txtDaNumeroDocumento.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento>=" & fnNormString(Me.txtDaNumeroDocumento.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND NumeroDocumento<=" & fnNormString(Me.txtANumeroDocumento.Text)
        End If
    End If
End If

If Me.txtDaDataConferimento.Value > 0 Then
    If Me.txtADataConferimento.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento=" & fnNormDateString(Me.txtDaDataConferimento.Text)
    Else
        If Me.txtADataConferimento.Value <= Me.txtDaDataConferimento.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento=" & fnNormDateString(Me.txtDaDataConferimento.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento>=" & fnNormDateString(Me.txtDaDataConferimento.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataConferimento<=" & fnNormDateString(Me.txtADataConferimento.Text)
        End If
    End If
End If
If Me.txtDaDataConsegnaMerce.Value > 0 Then
    If Me.txtADataConsegnaMerce.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce=" & fnNormDateString(Me.txtDaDataConsegnaMerce.Text)
    Else
        If Me.txtADataConsegnaMerce.Value <= Me.txtDaDataConsegnaMerce.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce=" & fnNormDateString(Me.txtDaDataConsegnaMerce.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce>=" & fnNormDateString(Me.txtDaDataConsegnaMerce.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND DataConsegnaMerce<=" & fnNormDateString(Me.txtADataConsegnaMerce.Text)
        End If
    End If
End If
If Me.txtDaLavMerce.Value > 0 Then
    If Me.txtALavMerce.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione=" & fnNormDateString(Me.txtDaLavMerce.Text)
    Else
        If Me.txtALavMerce.Value <= Me.txtDaLavMerce.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione=" & fnNormDateString(Me.txtDaLavMerce.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione>=" & fnNormDateString(Me.txtDaLavMerce.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataLavorazione<=" & fnNormDateString(Me.txtALavMerce.Text)
        End If
    End If
End If
If Me.txtDaDataArrivoMerce.Value > 0 Then
    If Me.txtADataArrivoMerce.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce=" & fnNormDateString(Me.txtDaDataArrivoMerce.Text)
    Else
        If Me.txtADataArrivoMerce.Value <= Me.txtDaDataArrivoMerce.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce=" & fnNormDateString(Me.txtDaDataArrivoMerce.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce>=" & fnNormDateString(Me.txtDaDataArrivoMerce.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataArrivoMerce<=" & fnNormDateString(Me.txtADataArrivoMerce.Text)
        End If
    End If
End If
If Me.txtDaDataOrdine.Value > 0 Then
    If Me.txtADataOrdine.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine=" & fnNormDateString(Me.txtDaDataOrdine.Text)
    Else
        If Me.txtADataOrdine.Value <= Me.txtDaDataOrdine.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine=" & fnNormDateString(Me.txtDaDataOrdine.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine>=" & fnNormDateString(Me.txtDaDataOrdine.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND DataOrdine<=" & fnNormDateString(Me.txtADataOrdine.Text)
        End If
    End If
End If

If Me.txtDaDataOrdineCli.Value > 0 Then
    If Me.txtADataOrdineCli.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente=" & fnNormDateString(Me.txtDaDataOrdineCli.Text)
    Else
        If Me.txtADataOrdineCli.Value <= Me.txtDaDataOrdineCli.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente=" & fnNormDateString(Me.txtDaDataOrdineCli.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente>=" & fnNormDateString(Me.txtDaDataOrdineCli.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataOrdineCliente<=" & fnNormDateString(Me.txtADataOrdineCli.Text)
        End If
    End If
End If

If Me.txtDaDataCompLiq.Value > 0 Then
    If Me.txtADataCompLiq.Value = 0 Then
        sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq=" & fnNormDateString(Me.txtDaDataCompLiq.Text)
    Else
        If Me.txtADataCompLiq.Value <= Me.txtDaDataCompLiq.Value Then
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq=" & fnNormDateString(Me.txtDaDataCompLiq.Text)
        Else
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq>=" & fnNormDateString(Me.txtDaDataCompLiq.Text)
            sSQL_GENERALE = sSQL_GENERALE & " AND RV_PODataCompetenzaLiq<=" & fnNormDateString(Me.txtADataCompLiq.Text)
        End If
    End If
End If

If Me.cboTipoDocumento.ListIndex > 0 Then
    sSQL_GENERALE = sSQL_GENERALE & " AND IDTipoOggetto=" & Me.cboTipoDocumento.ItemData(Me.cboTipoDocumento.ListIndex)
    GET_SQL_PER_STAMPA = sSQL_GENERALE
    Exit Function
Else
    sSQL_TIPO_DOCUMENTO = " AND (IDTipoOggetto=2 OR IDTipoOggetto=114)"
    sSQL_TIPO_DOCUMENTO = " AND ("

    For I = 1 To Me.cboTipoDocumento.ListCount - 1
        sSQL_TIPO_DOCUMENTO = sSQL_TIPO_DOCUMENTO & "IDTipoOggetto=" & Me.cboTipoDocumento.ItemData(I) '& ")"
         If I < Me.cboTipoDocumento.ListCount - 1 Then
            sSQL_TIPO_DOCUMENTO = sSQL_TIPO_DOCUMENTO & " OR "
        End If
    Next
   
    sSQL_TIPO_DOCUMENTO = sSQL_TIPO_DOCUMENTO & ")"
    
End If

GET_SQL_PER_STAMPA = sSQL_GENERALE & sSQL_TIPO_DOCUMENTO

End Function
Private Sub fnGrigliaPartenza()
On Error GoTo ERR_fnGrigliaAssegnazione
    Dim sSQL As String
    Dim OLDCursor As Long
    Dim cl As dgColumnHeader
    
    sSQL = "SELECT * FROM RV_POIEControlloVendite "
    sSQL = sSQL & "WHERE IDMovimento=0"
    'sSQL = sSQL & GET_SQL
    
    OLDCursor = CnDMT.CursorLocation
    CnDMT.CursorLocation = 3
        
        If Not (rsGriglia Is Nothing) Then
            rsGriglia.Close
            Set rsGriglia = Nothing
        End If
        
        Set rsGriglia = New ADODB.Recordset
        rsGriglia.CursorLocation = adUseClient
        rsGriglia.Open sSQL, CnDMT.InternalConnection, adOpenKeyset, adLockPessimistic
    
        With Me.GrigliaControlloVendite
            .EnableMove = True
            .UpdatePosition = True
            .BooleanType = dgGraphic
            .SelectionMode = dgSelectRow
            .ColumnsHeader.Clear
            Set .PaintNotifyObj = gPaintNotify
            '.LoadUserSettings
                .ColumnsHeader.Add "IDMovimento", "IDMovimento", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDOggetto", "IDOggetto", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDTipoOggetto", "IDTipoOggetto", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDValoriOggettoDettaglio", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDAzienda", "IDValoriOggettoDettaglio", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDFunzione", "IDFunzione", dgInteger, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDProcessoPerFunzione", "IDProcessoPerFunzione", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "IDSezionale", "IDSezionale", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "Sezionale", "Sezionale", dgchar, True, 2000, dgAlignleft
                
                .ColumnsHeader.Add "DataMovimento", "Data Movimento", dgDate, False, 1500, dgAlignleft
                .ColumnsHeader.Add "DataDocumento", "Data documento", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "NumeroDocumento", "N° documento", dgNumeric, True, 1500, dgAlignleft
                .ColumnsHeader.Add "Oggetto", "Oggetto", dgchar, True, 3000, dgAlignleft
                .ColumnsHeader.Add "TipoOggettoDescrizione", "Tipo documento vendita", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDAnagraficaCliente", "IDAnagraficaCliente", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "AnagraficaCliente", "Cliente", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "NomeCliente", "Nome cliente", dgchar, False, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_POIDSitoPerAnagrafica", "RV_POIDSitoPerAnagrafica", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "SitoPerAnagrafica", "Destinazione diversa", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "IDArticoloVenduto", "IDArticoloVenuto", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceArticoloVenduto", "Codice Articolo", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "IDCategoriaMerceologica", "IDCategoriaMerceologica", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CategoriaMerceologica", "Categoria merceologica", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "IDCategoriaFiscale", "IDCategoriaFiscale", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CategoriaFiscale", "Categoria fiscale", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "DescrizioneArticoloVenduto", "Articolo", dgchar, True, 2000, dgAlignleft
                .ColumnsHeader.Add "IDUnitaDiMisuraVenduto", "IDUnitaDiMisura", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "UnitaDiMisuraVenduto", "U.M.", dgchar, True, 1000, dgAlignleft

                 Set cl = .ColumnsHeader.Add("QuantitaTotale", "Quantità", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PrezzoUnitario", "Imp. Uni.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("ScontoPerc1", "% Sc.1", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("ScontoPerc2", "% Sc.2", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PrezzoScontato", "Imp. uni. scont.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("Importo", "Importo riga", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                Set cl = .ColumnsHeader.Add("RV_PONumeroColli", "Colli", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("RV_POPesoLordo", "Peso lordo", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POTara", "Tara", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POPesoNetto", "Peso netto", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POQuantitaPezzi", "Numero Pezzi", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POCodiceLottoVendita", "Lotto di vendita", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "RV_POCodiceLotto", "Lotto di entrata", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "RV_POCodiceLottoCampagna", "Lotto di campagna", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDImballo", "RV_POIDImballo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceArticoloImballo", "Codice Imballo", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "DescrizioneArticoloImballo", "Descrizione imballo", dgchar, True, 2000, dgAlignleft
                 Set cl = .ColumnsHeader.Add("RV_POImportoUnitarioImballo", "Imp. Uni. Imb.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 
                 Set cl = .ColumnsHeader.Add("RV_POQuantitaLiquidazione", "Q.ta liq.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POMerceInclusaImballo", "Merce incluso imballo", dgBoolean, True, 1000, dgAlignleft
                 Set cl = .ColumnsHeader.Add("RV_POImportoInclusoImballo", "Var. Liq.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POPrezzoMedioInLiq", "Prezzo medio", dgBoolean, True, 1000, dgAlignleft
                
                Set cl = .ColumnsHeader.Add("RV_POImportoLiquidazione", "Imp. Uni. Liq.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                .ColumnsHeader.Add "RV_POIDTipoImportoVenditaLiq", "RV_POIDTipoImportoVenditaLiq", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "TipoImportoVenditaLiq", "Tipo forzatura importo in Liq.", dgchar, True, 1000, dgAlignleft

                .ColumnsHeader.Add "IDTipoDocumentoCoop", "IDTipoDocumentoCoop", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "TipoDocumentoCoop", "Documento di consegna merce", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "DataConsegnaMerce", "Data consegna merce", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_PODataConferimento", "Data Conferimento", dgDate, True, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_PONumeroConferimento", "N° Conferimento", dgNumeric, True, 1500, dgAlignleft
                .ColumnsHeader.Add "RV_POIDAnagraficaSocio", "RV_POIDAnagraficaSocio", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "AnagraficaSocio", "Socio/Fornitore", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "NomeSocio", "Nome", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDAnagraficaFatturazione", "RV_POIDAnagraficaFatturazione", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "AnagraficaFatturazione", "Socio/Fornitore per F.C.S.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "NomeAnagraficaFatturazione", "Nome per F.C.S.", dgchar, False, 2000, dgAlignleft
                
                .ColumnsHeader.Add "IDArticoloConferito", "IDArticoloConferito", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceArticoloConferito", "Codice Articolo Conf.", dgchar, True, 1700, dgAlignleft
                .ColumnsHeader.Add "ArticoloConferito", "Articolo Conf.", dgchar, True, 2000, dgAlignleft

                
                .ColumnsHeader.Add "RV_POIDTipoLavorazione", "RV_POIDTipoLavorazione", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "TipoLavorazione", "Tipo lavorazione merce", dgchar, False, 1700, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDTipoCategoria", "RV_POIDTipoCategoria", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "TipoCategoria", "Tipo categoria", dgchar, False, 1700, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDCalibro", "RV_POIDCalibro", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "Calibro", "Calibro", dgchar, False, 1700, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDTipoLavorazioneConf", "RV_POIDTipoLavorazioneConf", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "TipoLavorazioneConferimento", "Tipo lavorazione conf.", dgchar, False, 1700, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDPedana", "RV_POIDPedana", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "RV_POCodicePedana", "Codice pedana", dgchar, False, 1700, dgAlignleft
                
                .ColumnsHeader.Add "RV_POIDTipoPedana", "RV_POIDTipoPedana", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "CodiceTipoPedana", "Codice tipo pedana", dgchar, False, 1700, dgAlignleft
                .ColumnsHeader.Add "TipoPedana", "Tipo pedana", dgchar, False, 1700, dgAlignleft
                
                .ColumnsHeader.Add "IDTipoOggettoStatistica", "IDTipoOggettoStatistica", dgNumeric, False, 500, dgAlignleft

                .ColumnsHeader.Add "RV_POPrezzoMedioConf", "Prezzo medio conf.", dgBoolean, False, 1500, dgAligncenter
                .ColumnsHeader.Add "RV_PODataLavorazione", "Data lavorazione", dgDate, False, 1500, dgAlignleft


                Set cl = .ColumnsHeader.Add("ColliConferito", "Colli Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .ColumnsHeader.Add("PesoLordoConferito", "Peso lordo Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("TaraConferito", "Tara Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PesoNettoConferito", "Peso netto conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("PezziConferito", "Numero Pezzi conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                 Set cl = .ColumnsHeader.Add("RV_POQuantitaMovimentata", "Q.tà Mov. Vend", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                 Set cl = .ColumnsHeader.Add("Qta_UMConferito", "Q.tà Mov. Conf.", dgDouble, True, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                .ColumnsHeader.Add "RV_POTipoRigaCollegata", "RV_POTipoRigaCollegata", dgNumeric, False, 500, dgAlignleft
                 Set cl = .ColumnsHeader.Add("RV_POPrezzoMerceNetta", "Prezzo netto merce", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POVariazionePrezzoImballo", "Var. Prezzo imballo", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."

                 Set cl = .ColumnsHeader.Add("RV_POQuantitaLiqPerPrezzoMedio", "Q.tà per prezzo medio", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POVariazionePrezzoManuale", "Var. Liq. Manuale", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                 Set cl = .ColumnsHeader.Add("RV_POImportoRigaCommissioni", "Imp. commissioni", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 5
                    cl.FormatOptions.FormatNumericThousandSep = "."
                
                .ColumnsHeader.Add "RV_POIDIvaImballo", "RV_POIDIvaImballo", dgNumeric, False, 500, dgAlignleft
                .ColumnsHeader.Add "DescrizioneIvaImballo", "I.V.A. Imballo", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "CodiceIvaImballo", "Codice I.V.A. Imballo", dgchar, False, 1000, dgAlignleft
                 Set cl = .ColumnsHeader.Add("AliquotaIvaImballo", "Aliquota I.V.A. Imballo", dgDouble, False, 1000, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                    
                .ColumnsHeader.Add "RV_POIDLottoCampagnaLavorazione", "IDLottoCampagnaLavorazione", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "IDSezionaleConf", "IDSezionaleConf", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "SezionaleConf", "Sezionale conferimento", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDVettoreConf", "IDVettoreConf", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "VettoreConf", "Vettore conferimento", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDLuogoPresaMerce", "IDLuogoPresaMerce", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "LuogoPresaMerceConf", "Sede stoccaggio merce", dgchar, False, 2000, dgAlignleft
                .ColumnsHeader.Add "IDRV_POTipoOrdineConforme", "IDRV_POTipoOrdineConforme", dgNumeric, False, 500, dgAlignRight
                .ColumnsHeader.Add "TipoOrdineConformeConf", "Conformità", dgchar, False, 2000, dgAlignleft
                    
                    
            Set .Recordset = rsGriglia
            .LoadUserSettings
            .Refresh
      
        End With
    
    CnDMT.CursorLocation = OLDCursor
Exit Sub
ERR_fnGrigliaAssegnazione:
    MsgBox Err.Description, vbCritical, "Reperimento dati assegnazione"
End Sub
Private Function GET_NUMERO_FILIALE(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT COUNT(Filiale.IDFiliale) AS NumeroRecord "
sSQL = sSQL & "FROM Filiale INNER JOIN "
sSQL = sSQL & "AttivitaAzienda ON Filiale.IDAttivitaAzienda = AttivitaAzienda.IDAttivitaAzienda "
sSQL = sSQL & "WHERE AttivitaAzienda.IDAzienda=" & IDAzienda

Set rs = New ADODB.Recordset

rs.Open sSQL, TheApp.Database.InternalConnection


If rs.EOF Then
    GET_NUMERO_FILIALE = 1
Else
    GET_NUMERO_FILIALE = fnNotNullN(rs!NumeroRecord)
    
    If GET_NUMERO_FILIALE = 0 Then GET_NUMERO_FILIALE = 1
End If

rs.Close
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
