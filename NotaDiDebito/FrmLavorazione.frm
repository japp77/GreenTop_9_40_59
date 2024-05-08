VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#9.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.6#0"; "DmtCodDesc.ocx"
Begin VB.Form FrmLavorazione 
   Caption         =   "Creazione lotto di vendita"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   5760
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdConferma 
      Caption         =   "Conferma"
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lotto di conferimento"
      Height          =   2655
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   11295
      Begin DMTEDITNUMLib.dmtNumber txtNumeroConferimento 
         Height          =   285
         Left            =   7200
         TabIndex        =   72
         Top             =   480
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   253
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtIDLotto_Conf 
         Height          =   255
         Left            =   10320
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin VB.TextBox txtLottoArticolo_Conf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   68
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtCodiceLottoArticolo_Conf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   66
         Top             =   1680
         Width           =   3375
      End
      Begin DMTEDITNUMLib.dmtNumber txtIDArticolo_Conf 
         Height          =   255
         Left            =   9480
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTDATETIMELib.dmtDate txtDataConferimento 
         Height          =   285
         Left            =   5640
         TabIndex        =   64
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   253
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtIDConferimento 
         Height          =   255
         Left            =   8640
         TabIndex        =   62
         Top             =   360
         Visible         =   0   'False
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtDispLotto 
         Height          =   285
         Left            =   5880
         TabIndex        =   60
         Top             =   2250
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   503
         _StockProps     =   253
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
      End
      Begin VB.TextBox txtSocio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtImballo_Conf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8640
         TabIndex        =   42
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtCodiceImballo_Conf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   41
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtArticolo_Conf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   40
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtCodiceArticolo_Conf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   2895
      End
      Begin DMTEDITNUMLib.dmtNumber txtQtaUM_Conf 
         Height          =   285
         Left            =   4920
         TabIndex        =   43
         Top             =   2250
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPezzi_Conf 
         Height          =   285
         Left            =   3960
         TabIndex        =   44
         Top             =   2250
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTara_Conf 
         Height          =   285
         Left            =   3000
         TabIndex        =   45
         Top             =   2250
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoNetto_Conf 
         Height          =   285
         Left            =   2040
         TabIndex        =   46
         Top             =   2250
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoLordo_Conf 
         Height          =   285
         Left            =   1080
         TabIndex        =   47
         Top             =   2250
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtColli_Conf 
         Height          =   285
         Left            =   120
         TabIndex        =   48
         Top             =   2250
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "N° conferimento"
         Height          =   255
         Index           =   24
         Left            =   7200
         TabIndex        =   71
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Descrizione lotto articolo"
         Height          =   255
         Index           =   23
         Left            =   3600
         TabIndex        =   69
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Codice lotto articolo"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   67
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Data conferimento"
         Height          =   255
         Index           =   21
         Left            =   5640
         TabIndex        =   63
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Disp. lotto"
         Height          =   255
         Index           =   20
         Left            =   5880
         TabIndex        =   61
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Imballo"
         Height          =   255
         Index           =   4
         Left            =   8640
         TabIndex        =   59
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Codice imballo"
         Height          =   255
         Index           =   3
         Left            =   6720
         TabIndex        =   58
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Articolo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3120
         MouseIcon       =   "FrmLavorazione.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         MouseIcon       =   "FrmLavorazione.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label2 
         Caption         =   "Q.tà"
         Height          =   255
         Index           =   19
         Left            =   4920
         TabIndex        =   54
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Peso netto"
         Height          =   255
         Index           =   18
         Left            =   2040
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tara "
         Height          =   255
         Index           =   17
         Left            =   3000
         TabIndex        =   52
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Pezzi"
         Height          =   255
         Index           =   16
         Left            =   3960
         TabIndex        =   51
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Peso lordo"
         Height          =   255
         Index           =   13
         Left            =   1080
         TabIndex        =   50
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Colli"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   49
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Codice articolo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "FrmLavorazione.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Creazione lotto vendita"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   11295
      Begin DMTDataCmb.DMTCombo cboCalibro 
         Height          =   315
         Left            =   1920
         TabIndex        =   73
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.TextBox txtImballo 
         Height          =   285
         Left            =   7440
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtDescrizioneLotto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtCodiceLotto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtArticolo 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin DMTDataCmb.DMTCombo cboCategoria 
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Top             =   1200
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
      Begin DmtCodDescCtl.DmtCodDesc CDPedana 
         Height          =   615
         Left            =   7440
         TabIndex        =   15
         Top             =   1500
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         PropCodice      =   $"FrmLavorazione.frx":091E
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmLavorazione.frx":0967
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmLavorazione.frx":09B1
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
      End
      Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         PropCodice      =   $"FrmLavorazione.frx":0A0B
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmLavorazione.frx":0A4D
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmLavorazione.frx":0A98
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
      End
      Begin DMTDataCmb.DMTCombo cboTipoLavorazione 
         Height          =   315
         Left            =   9120
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
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
      Begin DMTDATETIMELib.dmtDate txtDataLavorazione 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1770
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   253
         BackColor       =   16777215
         Appearance      =   1
      End
      Begin DMTEDITNUMLib.dmtNumber txtTaraUnitaria 
         Height          =   285
         Left            =   9960
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   253
         BackColor       =   16777215
         Enabled         =   0   'False
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
      End
      Begin DMTDataCmb.DMTCombo cboUM 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin DmtCodDescCtl.DmtCodDesc CDCodiceImballo 
         Height          =   615
         Left            =   5640
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         PropCodice      =   $"FrmLavorazione.frx":0AF2
         BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PropDescrizione =   $"FrmLavorazione.frx":0B3C
         BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MenuFunctions   =   $"FrmLavorazione.frx":0B86
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
      End
      Begin DMTEDITNUMLib.dmtNumber txtQta_UM 
         Height          =   285
         Left            =   6480
         TabIndex        =   14
         Top             =   1770
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPezzi 
         Height          =   285
         Left            =   5520
         TabIndex        =   13
         Top             =   1770
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtTara 
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Top             =   1770
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoNetto 
         Height          =   285
         Left            =   4560
         TabIndex        =   11
         Top             =   1800
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtPesoLordo 
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   1770
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin DMTEDITNUMLib.dmtNumber txtColli 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   1770
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   253
         Text            =   "0"
         BackColor       =   16777215
         Appearance      =   1
         UseSeparator    =   -1  'True
         DecFinalZeros   =   -1  'True
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Calibro"
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   35
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data lavorazione"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Tara unitaria"
         Height          =   255
         Index           =   10
         Left            =   9960
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Unità di misura"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblImballo 
         Caption         =   "Imballo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7440
         TabIndex        =   31
         Top             =   390
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Descrizione lotto"
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   30
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Codice lotto"
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   29
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Q.tà"
         Height          =   255
         Index           =   5
         Left            =   6480
         TabIndex        =   28
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Peso netto"
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   27
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Tara "
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   26
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Pezzi"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   25
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Peso lordo"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Colli"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   23
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblArticolo 
         Caption         =   "Articolo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   22
         Top             =   390
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo lavorazione"
         Height          =   255
         Index           =   14
         Left            =   9120
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Categoria"
         Height          =   255
         Index           =   15
         Left            =   3120
         TabIndex        =   20
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmLavorazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Link_Magazzino_Vendita As Long
Private Link_Magazzino_Conferimento As Long
Private Link_Causale_MagCar_Conf As Long
Private Link_Causale_MagScar_Conf As Long
Private Link_Causale_MagCar_Vend As Long
Private Link_Causale_MagScar_Vend As Long

Private Link_Magazzino_Carico As Long
Private Link_Magazzino_Scarico As Long
Private Link_CausaleCarico As Long
Private Link_CausaleScarico As Long
Private Link_Esercizio As Long

Private LINK_LAVORAZIONE As Long
Private Link_UM_Conf As Long
Private Link_TipoCaloPeso As Long
Private Link_TipoAumentoPeso As Long
Private Link_TipoScarto As Long
Private Link_MovimentoVendita As Long
Private Link_MovimentoConferimento As Long
Private Link_LottoArticolo As Long
Private Link_Socio As Long
Private Flag_Close_Conf As Integer
Private Link_UM_Coop As Long
Private Link_UM_Vendita As Long
Private NomeSocio As String
Private CodiceSocio As String
Private LottoDiRiferimento As String

Private NonRiportoInstrastat As Boolean






Private Sub cboUM_Click()
    Link_UM_Coop = fnGetUMCoop(Me.cboUM.CurrentID)
End Sub

Private Sub CDArticolo_ChangeElement()
On Error Resume Next
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset


    If CDArticolo.KeyFieldID > 0 Then
        Me.TxtArticolo.Text = Me.CDArticolo.Description
        sSQL = "SELECT IDUnitaDiMisuraVendita, RV_POIDImballoVendita, RV_POIDCalibro, "
        sSQL = sSQL & "IDTipoProdotto, RV_POIDTipoCategoria, RV_POIDTipoLavorazione "
        sSQL = sSQL & "FROM Articolo WHERE IDArticolo=" & Me.CDArticolo.KeyFieldID
        Set rs = Cn.OpenResultset(sSQL)
        If rs.EOF = False Then
            If (rs!IDTipoProdotto = Link_TipoCaloPeso) Or (rs!IDTipoProdotto = Link_TipoAumentoPeso) Or (rs!IDTipoProdotto = Link_TipoScarto) Then
                MsgBox "Il prodotto selezionato non è configurato come prodotto di vendita", vbInformation, "Creazione lavorazione"
                Me.CDArticolo.SetFocus
                Exit Sub
            End If
            
            Me.cboUM.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
            Me.CDCodiceImballo.Load IIf(IsNull(rs!RV_POIDImballoVendita), 0, rs!RV_POIDImballoVendita)
            Me.cboCalibro.WriteOn fnNotNullN(rs!RV_POIDCalibro)
            Me.cboCategoria.WriteOn fnNotNullN(rs!RV_POIDTipoCategoria)
            Me.cboTipoLavorazione.WriteOn fnNotNullN(rs!RV_POIDTipoLavorazione)
            
        End If
        
        rs.CloseResultset
        Set rs = Nothing

    End If
    
    ControllaTabulazione
    

End Sub

Private Sub CDCodiceImballo_ChangeElement()
    If Me.CDCodiceImballo.KeyFieldID > 0 Then
        Me.txtImballo.Text = Me.CDCodiceImballo.Description
        Me.txtTaraUnitaria.Value = fnGetTaraImballo
    End If
    
    

End Sub



Private Sub cmdConferma_Click()
On Error GoTo ERR_cmdConferma_Click
Dim StringaPermesso As String
StringaPermesso = PermessoSalvataggio
If Len(StringaPermesso) > 0 Then
    MsgBox StringaPermesso, vbInformation, "Creazione lotto di vendita"
    Exit Sub
End If
Screen.MousePointer = 11
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 4

Link_LottoArticolo = fnGetLottoArticolo
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1


Movimenti

If (Me.txtDispLotto.Value - Me.txtQta_UM.Value) <= 0 Then
    If MsgBox("ATTENZIONE!!" & "Vuoi chiudere il lotto di conferimento?", vbQuestion + vbYesNo, "Chiusura lotto di conferimento") = vbYes Then
        Flag_Close_Conf = 1
    Else
        Flag_Close_Conf = 0
    End If
Else
    Flag_Close_Conf = 0
End If


SalvaLavorazione
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1

If MsgBox("Riportare la lavorazione nel documento?", vbInformation + vbYesNo, "Creazione lotto di vendita") = vbYes Then
    RiportaLavorazioneInDocumento
    Var_RiportoRiga_Da_Nuova_Lavorazione = 1
End If
Screen.MousePointer = 0
Unload Me
Exit Sub
ERR_cmdConferma_Click:
    MsgBox Err.Description, vbCritical, "Creazione lotto di vendita"
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
    Link_ArticoloPadre = Me.txtIDArticolo_Conf.Value
    frmArticoliDerivati.Show vbModal
    ControllaTabulazione
End If
End Sub

Private Sub Form_Load()
    Me.Icon = gResource.GetIcon(IDI_DIAMANTE16)
    ParametroTipoCaloPeso
    ParametroTipoAumentoPeso
    ParametroTipoScarto
    
    'Parametri di default
    
    Link_Magazzino_Conferimento = fnGetParametriMagazzino("IDMagazzino_Carico")
    Link_Magazzino_Vendita = fnGetParametriMagazzino("IDMagazzino_Vendita")
    Link_Causale_MagCar_Conf = fnGetParametriMagazzino("IDCausale_Carico_Mag_Carico")
    Link_Causale_MagScar_Conf = fnGetParametriMagazzino("IDCausale_Scarico_Mag_Carico")
    Link_Causale_MagCar_Vend = fnGetParametriMagazzino("IDCausale_Carico_Mag_Vendita")
    Link_Causale_MagScar_Vend = fnGetParametriMagazzino("IDCausale_Scarico_Mag_vendita")
    fnGetMagazzinoScarico
    fnGetMagazzinoCarico
    Link_Esercizio = fnGetEsercizio(Me.txtDataLavorazione.Text)
    
    INIT_CONTROLLI
    
    Me.Frame1.Enabled = False
End Sub
Private Sub INIT_CONTROLLI()
     With Me.CDArticolo
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))" & " AND GestioneLotti = " & fnNormBoolean(1)
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Imballo
    With Me.CDCodiceImballo
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & TheApp.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
 
    

    'Unita di misura
    With Me.cboUM
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With
    
    With Me.cboTipoLavorazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .SQL = "SELECT * FROM RV_POTipoLavorazione"
        .Fill
    End With

     With Me.CDPedana
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Descrizione"
        .KeyField = "IDRV_POPedana"
        .TableName = "RV_POPedana"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))" & " AND GestioneLotti = " & fnNormBoolean(1)
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice pedana"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione padana"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("RV_POPedana")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    With Me.cboCategoria
       Set .Database = TheApp.Database.Connection
       .AddFieldKey "IDRV_POTipoCategoria"
       .DisplayField = "TipoCategoria"
       .SQL = "SELECT * FROM RV_POTipoCategoria"
       .Fill
    End With

    With Me.cboCalibro
       Set .Database = TheApp.Database.Connection
       .AddFieldKey "IDRV_POCalibro"
       .DisplayField = "Calibro"
       .SQL = "SELECT * FROM RV_POCalibro"
       .Fill
    End With

End Sub

Private Function fncTrovaIDFunzione(Gestore As String) As Long
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE (Gestore.Gestore = " & fnNormString(Gestore) & ") "
sSQL = sSQL & "AND (Funzione.IDFunzione >= 10000)"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub Label1_Click(Index As Integer)
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtADOLib.adoResultset


    Set oSearch = New dmtFind.Find
    oSearch.Database = Cn
    oSearch.Caption = "Conferimento"
    oSearch.AddDisplayField "Chiuso", "Chiuso", 1
    oSearch.AddDisplayField "Data conferimento", "DataDocumento", 2
    oSearch.AddDisplayField "Socio", "Anagrafica", 1
    oSearch.AddDisplayField "Codice Articolo", "CodiceArticolo", 1
    oSearch.AddDisplayField "Articolo", "Articolo", 1
    oSearch.AddDisplayField "Codice lotto", "CodiceLotto", 1
    oSearch.AddDisplayField "Descrizione lotto", "DescrizioneLotto", 1
    oSearch.AddDisplayField "Q.tà", "Qta_UM", 0
    oSearch.AddDisplayField "Colli", "Colli", 0
    oSearch.AddDisplayField "PesoLordo", "PesoLordo", 0
    oSearch.AddDisplayField "PesoNetto", "PesoNetto", 0
    oSearch.AddDisplayField "Tara", "Tara", 0
    
    
    oSearch.Filters.Add "Chiuso", 0
    oSearch.Filters.Add "Anagrafica", ""
    oSearch.Filters.Add "DataDocumento", ""
    oSearch.Filters.Add "Articolo", ""
    oSearch.Filters.Add "CodiceArticolo", ""
   
    
    oSearch.Start = 0
   
            
    sSQL = "SELECT RV_POCaricoMerceRighe.*, RV_POCaricoMerceTesta.*, Fornitore.Codice AS CodiceFornitore "
    sSQL = sSQL & "FROM RV_POCaricoMerceTesta INNER JOIN "
    sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta INNER JOIN "
    sSQL = sSQL & "Fornitore ON RV_POCaricoMerceTesta.IDAnagrafica = Fornitore.IDAnagrafica "
    sSQL = sSQL & "WHERE RV_POCaricoMerceRighe.Chiuso=" & fnNormBoolean(0)
    sSQL = sSQL & " AND RV_POCaricoMerceTesta.IDAzienda=" & TheApp.IDFirm

            oSearch.SQL = fnAnsi2Jet(sSQL)
        
            Set oRes = oSearch.Exec
    
    
            If Not oRes.EOF Then
                Me.txtIDConferimento.Value = fnNotNullN(oRes!IDRV_POCaricoMerceRighe)
                Me.txtIDArticolo_Conf.Value = fnNotNullN(oRes!IDArticolo)
                Me.txtCodiceArticolo_Conf.Text = fnNotNull(oRes!CodiceArticolo)
                Me.txtArticolo_Conf.Text = fnNotNull(oRes!Articolo)
                Link_Socio = fnNotNullN(oRes!IDAnagrafica)
                CodiceSocio = fnNotNull(oRes!CodiceFornitore)
                NomeSocio = fnNotNull(oRes!Nome)
                Me.txtSocio.Text = fnNotNull(oRes!Anagrafica)
                Me.txtIDLotto_Conf.Value = fnNotNullN(oRes!IDCodiceLotto)
                Me.txtCodiceImballo_Conf.Text = fnNotNull(oRes!CodiceImballo)
                Me.txtImballo_Conf.Text = fnNotNull(oRes!DescrizioneImballo)
                Me.txtLottoArticolo_Conf.Text = fnNotNull(oRes!DescrizioneLotto)
                Me.txtCodiceLottoArticolo_Conf.Text = fnNotNull(oRes!CodiceLotto)
                Link_UM_Conf = fnNotNullN(oRes!IDUnitaDiMisuraDiamante)
                Me.txtColli_Conf.Value = fnNotNullN(oRes!Colli)
                Me.txtPesoLordo_Conf.Value = fnNotNullN(oRes!PesoLordo)
                Me.txtTara_Conf.Value = fnNotNullN(oRes!Tara)
                Me.txtPesoNetto_Conf.Value = fnNotNullN(oRes!PesoNetto)
                Me.txtPezzi_Conf.Value = fnNotNullN(oRes!Pezzi)
                Me.txtQtaUM_Conf.Value = fnNotNull(oRes!Qta_UM)
                Me.txtDispLotto.Value = DispLotto(Me.txtIDLotto_Conf.Value)
                Me.txtDataConferimento.Value = fnNotNull(oRes!DataDocumento)
                Me.txtNumeroConferimento.Value = fnNotNull(oRes!NumeroDocumento)
                LottoDiRiferimento = fnNotNull(oRes!LottoDiConferimento)
                Me.txtDataLavorazione.Value = Date
                Me.Frame1.Enabled = True
            End If
            
       

End Sub
Private Sub ParametroTipoCaloPeso()
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT IDTipoCaloPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoCaloPeso = fnNotNullN(rs!IDTipoCaloPeso)
Else
    Link_TipoCaloPeso = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroTipoAumentoPeso()
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT IDTipoAumentoPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoAumentoPeso = fnNotNullN(rs!IDTipoAumentoPeso)
Else
    Link_TipoAumentoPeso = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroTipoScarto()
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT IDTipoScarto FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoScarto = fnNotNullN(rs!IDTipoScarto)
Else
    Link_TipoScarto = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function fnGetTaraImballo() As Double
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT Tara FROM Articolo WHERE "
    sSQL = sSQL & "IDArticolo = " & Me.CDCodiceImballo.KeyFieldID
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        If IsNull(rs!Tara) Then
            fnGetTaraImballo = 0
        Else
            fnGetTaraImballo = rs!Tara
        End If
        
    Else
        
            fnGetTaraImballo = 0
        
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Function DispLotto(Link_Lotto As Long) As Double
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT Giacenza FROM LottoArticoloPerMagazzino "
sSQL = sSQL & "WHERE IDLottoArticolo=" & Link_Lotto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = True Then
    DispLotto = 0
Else
    If IsNull(rs!Giacenza) Then
        DispLotto = 0
    Else
        DispLotto = rs!Giacenza
    End If
End If
End Function
Private Function PermessoSalvataggio() As String
PermessoSalvataggio = ""

If Me.txtIDConferimento.Value = 0 Then
    PermessoSalvataggio = "Selezionare il conferimento da agganciare alla lavorazione"
    Exit Function
End If
If Me.CDArticolo.KeyFieldID = 0 Then
    PermessoSalvataggio = "Selezionare il prodotto di vendita"
    Exit Function
End If
If Me.cboUM.CurrentID = 0 Then
    PermessoSalvataggio = "Manca l'unita di misura"
    Exit Function
End If
If Me.txtQta_UM.Value = 0 Then
    PermessoSalvataggio = "Manca la quantità di movimentazione"
    Exit Function
End If


End Function
Public Function fnGetParametriMagazzino(NomeCampo As String) As Long
    Dim rsEse As DmtADOLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
    sSQL = sSQL & "WHERE ((IDUtente=" & TheApp.IDUser & ") "
    sSQL = sSQL & "AND (IDFiliale=" & TheApp.Branch & "))"
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
            fnGetParametriMagazzino = rsEse.adoColumns(NomeCampo).Value
        Else
            sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
            sSQL = sSQL & "WHERE ((IDFiliale=" & m_App.Branch & ") "
            sSQL = sSQL & "AND (IDUtente=0))"
        
            Set rsEse = Cn.OpenResultset(sSQL)
        
            If rsEse.EOF = False Then
                If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
                    fnGetParametriMagazzino = rsEse.adoColumns(NomeCampo).Value
                Else
                    fnGetParametriMagazzino = 0
                End If
            Else
                fnGetParametriMagazzino = 0
            End If
            
        End If
    Else
        sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
        sSQL = sSQL & "WHERE ((IDFiliale=" & TheApp.Branch & ") "
        sSQL = sSQL & "AND (IDUtente=0))"
        
        Set rsEse = Cn.OpenResultset(sSQL)
        
        If rsEse.EOF = False Then
            If IsNull(rsEse.adoColumns(NomeCampo).Value) = False Then
                fnGetParametriMagazzino = rsEse.adoColumns(NomeCampo).Value
            Else
                fnGetParametriMagazzino = 0
            End If
        Else
            fnGetParametriMagazzino = 0
        End If
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Private Function fnGetCodiceLotto(TipoLotto As Integer, TipoStringaLotto As Integer, StringaLotto As String, Link_LottoArticolo As Long) As String
Dim rs As DmtADOLib.adoResultset
Dim sSQL As String
Dim Codice As String
Dim I As Integer
Dim PosizioneCodice As Integer
Dim Stringa As String
Dim StringaElaborata As String

sSQL = "SELECT RV_POLottoCostruzioneRighe.IDRV_POLottoComp, RV_POLottoCostruzioneRighe.Posizione, RV_POLottoCostruzioneRighe.Lunghezza, "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.Testo, RV_POLottoCostruzioneTesta.PosVendita "
sSQL = sSQL & "FROM RV_POLottoCostruzioneRighe INNER JOIN "
sSQL = sSQL & "RV_POLottoCostruzioneTesta ON "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.IDRV_POLottoCostruzioneTesta = RV_POLottoCostruzioneTesta.IDRV_POLottoCostruzioneTesta "
sSQL = sSQL & "WHERE RV_POLottoCostruzioneTesta.IDFiliale=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoLotto=" & TipoLotto
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoStringaLotto=" & TipoStringaLotto
sSQL = sSQL & " ORDER BY Posizione"

fnGetCodiceLotto = ""
StringaElaborata = ""
Set rs = Cn.OpenResultset(sSQL)

If Link_LottoArticolo > 0 Then
    For I = Len(CStr(Link_LottoArticolo)) To 7
        Codice = Codice & "0"
    Next
    Codice = Codice & Link_LottoArticolo
    If PosizioneCodice = 1 Then
        Codice = Codice & "_"
    Else
        Codice = "_" & Codice
    End If
End If
            

If rs.EOF Then
    StringaElaborata = ""
Else
    fnGetCodiceLotto = StringaLotto
    While Not rs.EOF
        Select Case fnNotNullN(rs!IDRV_POLottoComp)
            Case 1 'Codice socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(CodiceSocio), 0)
            Case 2 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtSocio.Text), 1)
            Case 3 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(NomeSocio), 1)
            Case 4 'Giorno conferimento

            Case 5 'Mese del conferimento
            
            Case 6 'Anno del conferimento
            
            Case 7 'Giorno lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("d", Me.txtDataLavorazione.Text)), 0)
            Case 8 'Mese lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("m", Me.txtDataLavorazione.Text)), 0)
            Case 9 'Anno lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("yyyy", Me.txtDataLavorazione.Text)), 0)
            Case 10 'calibro
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.cboCalibro.Text), 1)
            Case 11 'Tipo lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.cboTipoLavorazione.Text), 1)
            Case 12 'Tipo categoria
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.cboCategoria.Text), 1)
            Case 13 'Carattere speciale "_"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("_"), 1)
            Case 14 'Carattere speciale "-"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(rs!LottoComp)), 1)
            Case 15 'Stringa personalizzata
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(rs!Testo)), 1)
            Case 16 'Codice imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.CDCodiceImballo.Code), 1)
            Case 17 'Descrizione imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtImballo.Text), 1)
            Case 18 'Codice pedana
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.CDPedana.Code), 1)
            Case 19 'Codice articolo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.CDArticolo.Code), 1)
            Case 20 'Descrizione articolo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.TxtArticolo.Text), 1)
            Case 22 'Numero della settimana
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("ww", Me.txtDataLavorazione.Text)), 0)
            Case 23 'giorno dell'anno
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", Me.txtDataLavorazione.Text)), 0)
            Case 24 'Lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(LottoDiRiferimento), 1)
            Case 25 'Lotto di conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtCodiceLottoArticolo_Conf.Text), 1)
        End Select
    rs.MoveNext
    Wend
End If

    If StringaElaborata = "" Then
        If PosizioneCodice = 1 Then
            fnGetCodiceLotto = Mid(Codice, 1, Len(Codice) - 1)
        Else
            fnGetCodiceLotto = Mid(Codice, 2, Len(Codice))
        End If
        
    Else
        If PosizioneCodice = 1 Then
            fnGetCodiceLotto = Codice & StringaElaborata
        Else
            fnGetCodiceLotto = StringaElaborata & Codice
        End If
        
    End If



    
    
End Function
Private Function GET_STRINGALOTTO(IDRV_POLottoComp As Long, Lunghezza As Integer, Stringa As String, TipoStringa As Integer, Optional TestoPersonalizzato As String) As String
Dim Parziale As String
Dim I As Integer
Parziale = ""

        If Len(Stringa) >= Lunghezza Then
            GET_STRINGALOTTO = Mid(Stringa, 1, Lunghezza)
        Else
            For I = Len(Stringa) To Lunghezza - 1
                If TipoStringa = 0 Then
                    Parziale = "0" & Parziale
                Else
                    Parziale = "" & Parziale
                End If
            Next
            GET_STRINGALOTTO = Parziale & Stringa
        End If
End Function

Private Function fnGetLottoArticolo() As Long
Dim sSQL As String
Dim IDLotto As Long
    
    IDLotto = fnGetNewKey("LottoArticolo", "IDLottoArticolo")
    
    Me.txtCodiceLotto.Text = fnGetCodiceLotto(2, 1, "", IDLotto)
    Me.txtDescrizioneLotto.Text = fnGetCodiceLotto(2, 2, "", IDLotto)
    
    
    Link_LottoArticolo = IDLotto
    
    
    sSQL = "INSERT INTO LottoArticolo (IDLottoArticolo, IDArticolo, Codice, LottoArticolo, DataScadenza) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & IDLotto & ", "
    sSQL = sSQL & Me.CDArticolo.KeyFieldID & ", "
    sSQL = sSQL & fnNormString(Me.txtCodiceLotto.Text) & ", "
    sSQL = sSQL & fnNormString(Me.txtDescrizioneLotto.Text) & ", "
    sSQL = sSQL & fnNormDate(DateAdd("m", 1, Me.txtDataLavorazione.Text)) & ")"
    Cn.Execute sSQL
    
    fnGetLottoArticolo = IDLotto
    
    
End Function
Private Function InserimentoLottoInMagazzino() As Boolean
'Restituisce True se tutte le operazioni sono andate a buon fine
'Altrimenti restiruisce False

    Dim sSQL As String
        
        sSQL = "INSERT INTO LottoArticoloPerMagazzino ("
        sSQL = sSQL & "IDLottoArticolo, IDMagazzino, Giacenza, DataUltimoCarico) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & Link_LottoArticolo & ", "
        sSQL = sSQL & Link_Magazzino_Carico & ", "
        sSQL = sSQL & fnNormNumber(Me.txtQta_UM.Value) & ", "
        sSQL = sSQL & fnNormDate(Date) & ")"
        
        Cn.Execute sSQL
        InserimentoLottoInMagazzino = True
    
End Function


Public Sub Movimenti()
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    sSQL = "SELECT RV_POSchemaCoop.IDFiliale, RV_POProcessiDocumentoCoop.IDRV_POProcessiDocumentoCoop, "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop, RV_POProcessiDocumentoCoop.IDDocumentoCoop,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoProcessoCoop, RV_POProcessiDocumentoCoop.IDMagazzino,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoMagazzino , RV_POSchemaCoop.IDUtente "
    sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop ON "
    sSQL = sSQL & "RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop "
    sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale = " & TheApp.Branch & ") AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDDocumentoCoop = 10) AND (RV_POSchemaCoop.IDUtente = 0)"
    Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
        Select Case rs!IDTipoProcessoCoop
            Case 1
                If GeneraMovimentoDiCarico = True Then
                    Link_MovimentoVendita = fnGetIDMovimentoMagazzino
                    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                End If
            Case 2
                If GeneraMovimentoDiScarico = True Then
                    Link_MovimentoConferimento = fnGetIDMovimentoMagazzino
                    Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                End If
        End Select
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
     
End Sub
Private Function GeneraMovimentoDiCarico() As Boolean

Dim Mov As DmtMovim.cMovimentazione
Set Mov = New DmtMovim.cMovimentazione
Set Mov.Connection = TheApp.Database.Connection

Mov.DataMovimento = Date
Mov.FattoreDiConversione = Null
Mov.GestioneLotti = True
Mov.GestioneMatricole = False
Mov.IDEsercizio = fnGetEsercizio(Me.txtDataConferimento.Text)
Mov.IDTipoOggetto = fnGetTipoOggetto("RV_POLavorazione")

Mov.IDFunzione = Link_CausaleCarico
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Link_Magazzino_Vendita
Mov.Cessione = 0

Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Link_Socio
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", Me.CDArticolo.KeyFieldID
Mov.Field "IDUnitaDiMisura", Me.cboUM.CurrentID
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Me.TxtArticolo.Text
Mov.Field "QuantitaTotale", Me.txtQta_UM.Value
Mov.Field "Importo", 0
Mov.Field "DataDocumento", Date
Mov.Field "Oggetto", Mid("Lavorazione del lotto dal conferimento " & Me.txtDataConferimento.Text & " Numero " & Me.txtNumeroConferimento.Value & " del socio " & Me.txtSocio.Text, 1, 100)
Mov.Field "IDTipoMovimento", 1
Mov.Field "IDLottoArticolo", Link_LottoArticolo
Mov.Field "TipoRiga", trcNessuno

GeneraMovimentoDiCarico = Mov.Insert
Set Mov = Nothing
End Function
Private Function GeneraMovimentoDiScarico() As Boolean

Dim Mov As DmtMovim.cMovimentazione
Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection

Mov.DataMovimento = Date
Mov.FattoreDiConversione = Null
Mov.GestioneLotti = True
Mov.GestioneMatricole = False
Mov.IDEsercizio = fnGetEsercizio(Me.txtDataConferimento.Text)
Mov.IDTipoOggetto = fnGetTipoOggetto("RV_POLavorazione")
Mov.IDFunzione = Link_CausaleScarico
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoUscita = Link_Magazzino_Scarico
Mov.IDMagazzinoEntrata = Link_Magazzino_Scarico
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Link_Socio
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", Me.txtIDArticolo_Conf.Value
Mov.Field "DescrizioneArticolo", Me.txtArticolo_Conf.Text
Mov.Field "IDUnitaDiMisura", Link_UM_Conf
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Me.txtArticolo_Conf.Text
Mov.Field "QuantitaTotale", Me.txtQta_UM.Value
Mov.Field "Importo", 0
Mov.Field "DataDocumento", Date
Mov.Field "Oggetto", Mid("Scarico della lavorazione del lotto in data " & Date, 1, 100)
Mov.Field "IDTipoMovimento", 1
Mov.Field "IDLottoArticolo", Me.txtIDLotto_Conf.Value
Mov.Field "TipoRiga", trcNessuno


GeneraMovimentoDiScarico = Mov.Insert
Set Mov = Nothing
End Function

Public Function fnGetMagazzinoCarico() As Long
   Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT RV_POSchemaCoop.IDFiliale, RV_POProcessiDocumentoCoop.IDRV_POProcessiDocumentoCoop, "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop, RV_POProcessiDocumentoCoop.IDDocumentoCoop,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoProcessoCoop, RV_POProcessiDocumentoCoop.IDMagazzino,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoMagazzino , RV_POSchemaCoop.IDUtente "
    sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop ON "
    sSQL = sSQL & "RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop "
    sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale = " & TheApp.Branch & ") AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDDocumentoCoop = 10) AND (RV_POSchemaCoop.IDUtente = 0) AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDTipoProcessoCoop = 1)"
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Select Case rs!IDTipoMagazzino
            Case 1
                Link_Magazzino_Carico = Link_Magazzino_Conferimento
                Link_CausaleCarico = Link_Causale_MagCar_Conf
            Case 2
                Link_Magazzino_Carico = Link_Magazzino_Vendita
                Link_CausaleCarico = Link_Causale_MagCar_Vend
                
        End Select
    Else
        Link_Magazzino_Carico = Link_Magazzino_Vendita
        Link_CausaleCarico = Link_Causale_MagCar_Vend
       
    End If
        
    
End Function
Public Function fnGetMagazzinoScarico() As Long
   Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT RV_POSchemaCoop.IDFiliale, RV_POProcessiDocumentoCoop.IDRV_POProcessiDocumentoCoop, "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop, RV_POProcessiDocumentoCoop.IDDocumentoCoop,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoProcessoCoop, RV_POProcessiDocumentoCoop.IDMagazzino,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoMagazzino , RV_POSchemaCoop.IDUtente "
    sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop ON "
    sSQL = sSQL & "RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop "
    sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale = " & TheApp.Branch & ") AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDDocumentoCoop = 10) AND (RV_POSchemaCoop.IDUtente = 0) AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDTipoProcessoCoop = 2)"
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Select Case rs!IDTipoMagazzino
            Case 1
                Link_Magazzino_Scarico = Link_Magazzino_Conferimento
                Link_CausaleScarico = Link_Causale_MagScar_Conf
            Case 2
                Link_Magazzino_Scarico = Link_Magazzino_Vendita
                Link_CausaleScarico = Link_Causale_MagScar_Vend
                
        End Select
    Else
        Link_Magazzino_Scarico = Link_Magazzino_Conferimento
        Link_CausaleScarico = Link_Causale_MagScar_Conf
    End If
        
    
End Function
Private Function fnGetEsercizio(dData As String) As Long
    Dim rsEse As DmtADOLib.adoResultset
    Dim sSQL As String
    
    sSQL = "Select IDEsercizio, DataInizio, DataFine FROM Esercizio WHERE "
    sSQL = sSQL & "((IDAzienda = " & TheApp.IDFirm & ")"
    sSQL = sSQL & " AND (DataInizio <= " & fnNormDate(dData) & ")"
    sSQL = sSQL & " AND (DataFine >= " & fnNormDate(dData) & "))"
   

    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetEsercizio = rsEse!IDEsercizio
    Else
        fnGetEsercizio = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Private Function fnGetNewKey(Tabella As String, CampoKey As String) As Long
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
        
        sSQL = "SELECT " & CampoKey & " FROM " & Tabella & " ORDER BY " & CampoKey & " DESC"
        
        Set rs = Cn.OpenResultset(fnAnsi2Jet(sSQL))
    
        If rs.EOF = True Then
        
            fnGetNewKey = 1
    
        Else
            
            fnGetNewKey = fnNotNullN(rs.adoColumns(CampoKey)) + 1
    
        End If

        rs.CloseResultset
        Set rs = Nothing
    

    
End Function
Private Function fnGetIDMovimentoMagazzino() As Long
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT IDMovimento FROM Movimento ORDER BY IDMovimento DESC"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        fnGetIDMovimentoMagazzino = rs!IDMovimento
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub SalvaLavorazione()
On Error GoTo ERR_SalvaLavorazione
   Dim sSQL As String
    
    LINK_LAVORAZIONE = fnGetNewKey("RV_POLavorazione", "IDRV_POLavorazione")
    sSQL = "INSERT INTO RV_POLavorazione ("
    sSQL = sSQL & "IDRV_POLavorazione, IDRV_POCaricoMerceRighe, Link_Riga_Interno, IDTipoLavorazione, IDRV_POTipoCategoria, IDRV_POCalibro, Calibro, "
    sSQL = sSQL & "DataDocumento, IDArticolo, CodiceArticolo, Articolo, "
    sSQL = sSQL & "IDUnitaDiMisura, IDUnitaDiMisuraCoop, Colli, PesoLordo, PesoNetto, Tara, Pezzi, Qta_UM, "
    sSQL = sSQL & "IDCodiceLotto_Vendita, CodiceLottoVendita, DescrizioneLottoVendita, "
    sSQL = sSQL & "IDImballoVendita, CodiceImballoVendita, ImballoVendita, "
    sSQL = sSQL & "IDMovimento_Vendita, IDMovimento_Carico, TaraUnitaria, Chiuso) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & LINK_LAVORAZIONE & ", "
    sSQL = sSQL & Me.txtIDConferimento.Value & ", "
    sSQL = sSQL & fnGetNewKey("RV_POLavorazione", "Link_Riga_Interno") & ", "
    sSQL = sSQL & Me.cboTipoLavorazione.CurrentID & ", "
    sSQL = sSQL & Me.cboCategoria.CurrentID & ", "
    sSQL = sSQL & Me.cboCalibro.CurrentID & ", "
    sSQL = sSQL & fnNormString(Me.cboCalibro.Text) & ", "
    sSQL = sSQL & fnNormDate(Me.txtDataLavorazione.Text) & ", "
    sSQL = sSQL & Me.CDArticolo.KeyFieldID & ", "
    sSQL = sSQL & fnNormString(Me.CDArticolo.Code) & ", "
    sSQL = sSQL & fnNormString(Me.TxtArticolo.Text) & ", "
    sSQL = sSQL & Me.cboUM.CurrentID & ", "
    sSQL = sSQL & Link_UM_Coop & ", "
    sSQL = sSQL & fnNormNumber(Me.txtColli.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtPesoLordo.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtPesoNetto.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtTara.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtPezzi.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtQta_UM.Value) & ", "
    sSQL = sSQL & Link_LottoArticolo & ", "
    sSQL = sSQL & fnNormString(Me.txtCodiceLotto.Text) & ", "
    sSQL = sSQL & fnNormString(Me.txtDescrizioneLotto.Text) & ", "
    sSQL = sSQL & Me.CDCodiceImballo.KeyFieldID & ", "
    sSQL = sSQL & fnNormString(Me.CDCodiceImballo.Code) & ", "
    sSQL = sSQL & fnNormString(Me.txtImballo.Text) & ", "
    sSQL = sSQL & Link_MovimentoVendita & ", "
    sSQL = sSQL & Link_MovimentoConferimento & ", "
    sSQL = sSQL & fnNormNumber(Me.txtTaraUnitaria.Value) & ", "
    sSQL = sSQL & fnNormBoolean(0) & ")"
    
    Cn.Execute sSQL
    
    
    fncTabellaDiCollegamentoVendita
    
    If Flag_Close_Conf = 1 Then
        AggiornaFlagChiuso_LottoConferimento
    End If

Exit Sub
ERR_SalvaLavorazione:
    MsgBox Err.Description, vbCritical, "Salva Lavorazione"

    
End Sub
Private Sub fncTabellaDiCollegamentoVendita()
On Error GoTo ERR_fncTabellaDiCollegamentoVendita
Dim sSQL As String
Dim rsCtrl As DmtADOLib.adoResultset

sSQL = "SELECT IDLottoArticolo FROM RV_POCollegamento "
sSQL = sSQL & "WHERE IDLottoArticolo = " & Link_LottoArticolo


Set rsCtrl = Cn.OpenResultset(sSQL)

If rsCtrl.EOF Then
    sSQL = "INSERT INTO RV_POCollegamento (IDLottoArticolo, QtaColliCaricati, QtaColliVenduti) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & Link_LottoArticolo & ", "
    sSQL = sSQL & fnNormNumber(Me.txtColli.Value) & ", "
    sSQL = sSQL & 0 & ")"
    
Else
    sSQL = "UPDATE RV_POCollegamento SET "
    sSQL = sSQL & "QtaColliCaricati=" & fnNormNumber(Me.txtColli.Value) & " "
    sSQL = sSQL & "WHERE IDLottoArticolo=" & Link_LottoArticolo
End If

Cn.Execute sSQL

rsCtrl.CloseResultset
Set rsCtrl = Nothing
Exit Sub
ERR_fncTabellaDiCollegamentoVendita:
    MsgBox Err.Description, vbCritical, "fncTabellaDiCollegamentoVendita"

End Sub
Private Sub AggiornaFlagChiuso_LottoConferimento()
On Error GoTo ERR_AggiornaFlagChiuso_LottoConferimento
Dim sSQL As String

sSQL = "UPDATE RV_POCaricoMerceRighe SET "
sSQL = sSQL & "Chiuso=" & fnNormBoolean(Flag_Close_Conf) & " "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & Me.txtIDConferimento.Value

Cn.Execute sSQL

Exit Sub
ERR_AggiornaFlagChiuso_LottoConferimento:
    MsgBox Err.Description, vbCritical, "AggiornaFlagChiuso_LottoConferimento"

End Sub
Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
    sSQL = sSQL & "IDUnitaDiMisura = " & fnNotNullN(Link_UMAcq)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetUMCoop = fnNotNullN(rs!RV_POIDUnitaDiMisuraCoop)
    Else
        fnGetUMCoop = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub txtColli_Change()
    If Link_UM_Coop = 1 Then
        Me.txtQta_UM.Value = Me.txtColli.Value
    End If
End Sub

Private Sub txtColli_LostFocus()
    Me.txtTara.Value = Me.txtColli.Value * Me.txtTaraUnitaria.Value
End Sub
Private Sub txtPesoLordo_Change()
    If Link_UM_Coop = 2 Then
        Me.txtQta_UM.Value = Me.txtPesoLordo.Value
    End If
End Sub

Private Sub txtPesoLordo_LostFocus()
    Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
End Sub

Private Sub txtPesoNetto_Change()
    If Link_UM_Coop = 3 Then
        Me.txtQta_UM.Value = Me.txtPesoNetto.Value
    End If
    
End Sub

Private Sub txtPesoNetto_LostFocus()
    Me.txtPesoLordo.Value = Me.txtPesoNetto.Value + Me.txtTara.Value
End Sub

Private Sub txtPezzi_Change()
    If Link_UM_Coop = 5 Then
        Me.txtQta_UM.Value = Me.txtPezzi.Value
    End If
    
End Sub

Private Sub txtTara_Change()
    If Link_UM_Coop = 4 Then
        Me.txtQta_UM.Value = Me.txtTara.Value
    End If
    
End Sub
Private Sub RiportaLavorazioneInDocumento()
        Link_Articolo = Me.CDArticolo.KeyFieldID
        frmMain.txtCodiceArticolo.Text = Me.CDArticolo.Code
        frmMain.txtDescrizioneArticolo.Text = Me.TxtArticolo.Text
        frmMain.cboAliquotaArticolo.WriteOn GET_AliquotaIvaVendita
        frmMain.cboUnitaDiMisura.WriteOn Me.cboUM.CurrentID
        frmMain.AGGIORNA_CONTENITORE_DATI_LOTTO
        frmMain.CDCodiceLotto.Load Link_LottoArticolo
        frmMain.txtDescrizioneLotto.Text = Me.txtDescrizioneLotto.Text
        frmMain.CDImballo.Load Me.CDCodiceImballo.KeyFieldID
        frmMain.txtDescrizioneImballo.Text = Me.txtImballo.Text
        frmMain.txtTaraUnitaria.Value = Me.txtTaraUnitaria.Value
        frmMain.txtColli.Value = Me.txtColli.Value
        frmMain.txtPesoLordo.Value = Me.txtPesoLordo.Value
            Val_PesoLordo = frmMain.txtPesoLordo.Value
        
        frmMain.txtTara.Value = Me.txtTara.Value
        frmMain.txtPesoNetto.Value = Me.txtPesoNetto.Value
        
        frmMain.txtPezzi.Value = Me.txtPezzi.Value
            Val_Pezzi = frmMain.txtPezzi.Value
        
        frmMain.txtQta_UM.Value = Me.txtQta_UM.Value
                'Instrastat
                If frmMain.chkCessione.Value = Checked Then
                    
                    If NonRiportoInstrastat = True Then
                        frmMain.chkRiportoIntra_Art = Checked
                    Else
                        frmMain.chkRiportoIntra_Art = Unchecked
                        frmMain.CDIntra_art_Nat_Trans.Load Link_Nat_Trans_Art
                        frmMain.CDIntra_art_Nom_Comb.Load Link_Nom_Comb_Art
                        frmMain.txtIntra_Art_MassaNetta.Value = MassaNetta_Art
                
                    End If
               
                Else
                
                End If
        
        
Exit Sub

End Sub
Private Function GET_AliquotaIvaVendita() As Long
Dim sSQL As String
Dim rs As DmtADOLib.adoResultset

sSQL = "SELECT * FROM Articolo WHERE IDArticolo=" & Me.CDArticolo.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_AliquotaIvaVendita = 0
Else
    GET_AliquotaIvaVendita = fnNotNullN(rs!IDIvaVendita)
    MassaNetta_Art = fnNotNullN(rs!MassaNettaInKg)
    Link_Nom_Comb_Art = fnNotNullN(IDNomenclaturaCombinata)
    Link_Nat_Trans_Art = fnNotNullN(RV_POIDNaturaTransazione)
    NonRiportoInstrastat = IIf(IsNull(rs!NonRiportoIntrastat), 0, rs!NonRiportoIntrastat)
End If
rs.CloseResultset
Set rs = Nothing
End Function

Private Sub txtTara_LostFocus()
    Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
End Sub
Private Sub ControllaTabulazione()
If Len(Me.CDCodiceImballo.Code) = 0 Then
    Me.CDCodiceImballo.SetFocus
    Exit Sub
End If
If Me.cboCalibro.CurrentID = 0 Then
    Me.cboCalibro.SetFocus
    Exit Sub
End If
If Me.cboCategoria.CurrentID = 0 Then
    Me.cboCategoria.SetFocus
    Exit Sub
End If
If Me.cboTipoLavorazione.CurrentID = 0 Then
    Me.cboTipoLavorazione.SetFocus
    Exit Sub
End If
If Len(Me.txtDataLavorazione.Text) = 0 Then
    MMe.txtDataLavorazione.SetFocus
    Exit Sub
End If
If Len(Me.txtDataLavorazione.Text) = 0 Then
    MMe.txtDataLavorazione.SetFocus
    Exit Sub
End If
Me.txtColli.SetFocus

End Sub
Private Function fnGetTipoOggetto(Optional Gestore As String) As Long
    Dim sSQL As String
    Dim rs As DmtADOLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    If Gestore = "" Then
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    Else
        sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)
    End If
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
