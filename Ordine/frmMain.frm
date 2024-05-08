VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{7A1D73E4-F461-11D0-8F01-004033A00AF2}#1.0#0"; "DmtWheel.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{9385BB2E-6637-11D1-850D-002018802E11}#3.1#0"; "Dmtsplit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{FCA49525-5F72-11D2-B9EB-00201880103B}#18.1#0"; "DMTPrinterDialog.OCX"
Object = "{A83BB158-4E50-11D2-B95E-002018813989}#8.3#0"; "DmtSearchAccount.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   11955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11955
   ScaleWidth      =   18135
   WhatsThisHelp   =   -1  'True
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   11610
      Left            =   0
      TabIndex        =   84
      Top             =   0
      Width           =   18135
      _LayoutVersion  =   2
      _ExtentX        =   31988
      _ExtentY        =   20479
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
      Begin DMTPrinterDialog.DMTDialog DmtPrnDlg 
         Left            =   11880
         Top             =   10320
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
      End
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   0
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.PictureBox PicForm 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   11295
         Left            =   0
         ScaleHeight     =   11265
         ScaleWidth      =   20145
         TabIndex        =   86
         Top             =   0
         Width           =   20175
         Begin VB.PictureBox PicForm2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   11115
            Left            =   0
            ScaleHeight     =   11085
            ScaleWidth      =   20025
            TabIndex        =   87
            Top             =   0
            Width           =   20055
            Begin VB.Frame FraTab 
               Caption         =   "Totale Documento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   900
               Index           =   6
               Left            =   0
               TabIndex        =   281
               Top             =   10080
               Width           =   17415
               Begin DMTEDITNUMLib.dmtCurrency curTotImponibile 
                  Height          =   315
                  Left            =   3480
                  TabIndex        =   282
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curTotImposta 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   283
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curTotDocumento 
                  Height          =   315
                  Left            =   5160
                  TabIndex        =   284
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   12648384
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curTotArrotondamenti 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   285
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   12648384
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curNettoAPagare 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   286
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   12648384
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curNettoAPagare_naz 
                  Height          =   315
                  Left            =   8520
                  TabIndex        =   287
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   12648384
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtPesoTotale 
                  Height          =   315
                  Left            =   11160
                  TabIndex        =   288
                  Top             =   480
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtColliTotali 
                  Height          =   315
                  Left            =   10200
                  TabIndex        =   289
                  Top             =   480
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtNPedaneTesta 
                  Height          =   315
                  Left            =   16440
                  TabIndex        =   290
                  Top             =   480
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtNPedaneTotale 
                  Height          =   315
                  Left            =   15360
                  TabIndex        =   291
                  Top             =   480
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtPezziTotale 
                  Height          =   315
                  Left            =   14280
                  TabIndex        =   292
                  Top             =   480
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtTaraTotale 
                  Height          =   315
                  Left            =   12240
                  TabIndex        =   293
                  Top             =   480
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtPesoNettoTotale 
                  Height          =   315
                  Left            =   13200
                  TabIndex        =   294
                  Top             =   480
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin VB.Label Label4 
                  Caption         =   "Netto a pagare naz."
                  Height          =   255
                  Index           =   4
                  Left            =   8520
                  TabIndex        =   307
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Netto a pagare corr."
                  Height          =   195
                  Index           =   23
                  Left            =   6840
                  TabIndex        =   306
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale arrotondamenti"
                  Height          =   195
                  Index           =   22
                  Left            =   120
                  TabIndex        =   305
                  Top             =   240
                  Width           =   1605
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale documento"
                  Height          =   195
                  Index           =   21
                  Left            =   5160
                  TabIndex        =   304
                  Top             =   240
                  Width           =   1530
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale imposta"
                  Height          =   195
                  Index           =   20
                  Left            =   1800
                  TabIndex        =   303
                  Top             =   240
                  Width           =   1050
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale imponibile"
                  Height          =   195
                  Index           =   19
                  Left            =   3480
                  TabIndex        =   302
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label Label1 
                  Caption         =   "N° pedane"
                  Height          =   255
                  Index           =   12
                  Left            =   16440
                  TabIndex        =   301
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Caption         =   "Peso"
                  Height          =   255
                  Index           =   4
                  Left            =   11160
                  TabIndex        =   300
                  Top             =   240
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Caption         =   "Colli"
                  Height          =   255
                  Index           =   3
                  Left            =   10200
                  TabIndex        =   299
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Caption         =   "N° ped. eff."
                  Height          =   255
                  Index           =   15
                  Left            =   15360
                  TabIndex        =   298
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Pezzi"
                  Height          =   255
                  Index           =   16
                  Left            =   14400
                  TabIndex        =   297
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Caption         =   "Tara"
                  Height          =   255
                  Index           =   17
                  Left            =   12240
                  TabIndex        =   296
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label1 
                  Caption         =   "Peso netto"
                  Height          =   255
                  Index           =   18
                  Left            =   13200
                  TabIndex        =   295
                  Top             =   240
                  Width           =   975
               End
            End
            Begin TabDlg.SSTab SSTab1 
               Height          =   10095
               Left            =   0
               TabIndex        =   88
               Top             =   0
               Width           =   17415
               _ExtentX        =   30718
               _ExtentY        =   17806
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Documento (F8)"
               TabPicture(0)   =   "frmMain.frx":479EA
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "fraAnaDest"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "fraContratto"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "FraTab(1)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "Frame1"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "FraTab(0)"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "cboCambioValuta"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "txtValoreCambioValuta"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "txtDataCambio"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "FraTab(2)"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "FraTab(4)"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "FraTab(5)"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "FraTab(7)"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "Frame4"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).Control(13)=   "Frame3"
               Tab(0).Control(13).Enabled=   0   'False
               Tab(0).ControlCount=   14
               TabCaption(1)   =   "Corpo (F9)"
               TabPicture(1)   =   "frmMain.frx":47A06
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "FraTab(3)"
               Tab(1).Control(1)=   "lblInfoTesta"
               Tab(1).ControlCount=   2
               Begin VB.Frame Frame3 
                  Caption         =   "Commissioni"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   3105
                  Left            =   11180
                  TabIndex        =   192
                  Top             =   4200
                  Width           =   6120
                  Begin VB.CommandButton Command2 
                     Caption         =   "GESTIONE COMMISSIONI"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Left            =   120
                     TabIndex        =   277
                     Top             =   2640
                     Width           =   5895
                  End
                  Begin MSComctlLib.ListView lvwCommissioni 
                     Height          =   2370
                     Left            =   120
                     TabIndex        =   278
                     Top             =   240
                     Width           =   5895
                     _ExtentX        =   10398
                     _ExtentY        =   4180
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   0
                     NumItems        =   0
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "Annotazioni interne"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   1215
                  Left            =   11160
                  TabIndex        =   279
                  Top             =   3120
                  Width           =   6135
                  Begin VB.TextBox txtAnnotazioniInterna 
                     Height          =   885
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   280
                     Top             =   240
                     Width           =   5895
                  End
               End
               Begin VB.Frame FraTab 
                  Caption         =   "Trasporto e spedizione"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   2820
                  Index           =   7
                  Left            =   120
                  TabIndex        =   98
                  Top             =   7200
                  Width           =   17175
                  Begin VB.TextBox txtTargaAutomezzo 
                     Height          =   315
                     Left            =   10920
                     TabIndex        =   31
                     Top             =   480
                     Width           =   2655
                  End
                  Begin VB.TextBox txtIstruzioniMittente 
                     Height          =   420
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   32
                     Top             =   1080
                     Width           =   16935
                  End
                  Begin VB.TextBox txtDescrizioneRigaDoc 
                     Height          =   405
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   34
                     Top             =   2280
                     Width           =   16935
                  End
                  Begin VB.TextBox txtAnnotazioni 
                     Height          =   405
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   33
                     Top             =   1680
                     Width           =   16935
                  End
                  Begin DMTDataCmb.DMTCombo cboAspettoEsteriore 
                     Height          =   315
                     Left            =   4200
                     TabIndex        =   30
                     Top             =   480
                     Width           =   2415
                     _ExtentX        =   4260
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
                  Begin DMTDataCmb.DMTCombo cboVettore 
                     Height          =   315
                     Left            =   6720
                     TabIndex        =   29
                     Top             =   480
                     Width           =   4095
                     _ExtentX        =   7223
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
                  Begin DMTDataCmb.DMTCombo cboTrasporto 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   28
                     Top             =   480
                     Width           =   1935
                     _ExtentX        =   3413
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
                  Begin DMTDataCmb.DMTCombo cboPorto 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   27
                     Top             =   480
                     Width           =   1935
                     _ExtentX        =   3413
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
                  Begin DMTDataCmb.DMTCombo cboVettoreSuccessivo 
                     Height          =   315
                     Left            =   13680
                     TabIndex        =   269
                     Top             =   480
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
                  Begin VB.Label Label4 
                     Caption         =   "Vettore successivo"
                     Height          =   255
                     Index           =   11
                     Left            =   13680
                     TabIndex        =   270
                     Top             =   240
                     Width           =   1455
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Targa automezzo"
                     Height          =   255
                     Index           =   11
                     Left            =   10920
                     TabIndex        =   194
                     Top             =   240
                     Width           =   1935
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Istruzioni del mittente"
                     Height          =   255
                     Index           =   6
                     Left            =   120
                     TabIndex        =   193
                     Top             =   840
                     Width           =   9735
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Annotazioni di fatturazione"
                     Height          =   255
                     Index           =   7
                     Left            =   120
                     TabIndex        =   168
                     Top             =   1485
                     Width           =   6975
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Annotazioni finali del corpo del documento di evasione"
                     Height          =   255
                     Index           =   9
                     Left            =   120
                     TabIndex        =   167
                     Top             =   2080
                     Width           =   6975
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Aspetto esteriore"
                     Height          =   255
                     Index           =   5
                     Left            =   4200
                     TabIndex        =   102
                     Top             =   240
                     Width           =   2415
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Vettore"
                     Height          =   255
                     Index           =   2
                     Left            =   6720
                     TabIndex        =   101
                     Top             =   240
                     Width           =   1815
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Trasporto"
                     Height          =   255
                     Index           =   1
                     Left            =   2160
                     TabIndex        =   100
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Porto"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   99
                     Top             =   240
                     Width           =   1575
                  End
               End
               Begin VB.Frame FraTab 
                  Caption         =   "Spese e Sconti"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   1395
                  Index           =   5
                  Left            =   6720
                  TabIndex        =   130
                  Top             =   5880
                  Width           =   4485
                  Begin DMTEDITNUMLib.dmtCurrency curSpeseIncasso 
                     Height          =   285
                     Left            =   930
                     TabIndex        =   23
                     Top             =   360
                     Width           =   915
                     _Version        =   65536
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Appearance      =   1
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtCurrency curSpeseTrasporto 
                     Height          =   285
                     Left            =   930
                     TabIndex        =   25
                     Top             =   840
                     Width           =   915
                     _Version        =   65536
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Appearance      =   1
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtNumber lngScontoDocPer 
                     Height          =   285
                     Left            =   2550
                     TabIndex        =   24
                     Top             =   390
                     Width           =   915
                     _Version        =   65536
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtCurrency curScontoDocImp 
                     Height          =   285
                     Left            =   2550
                     TabIndex        =   26
                     Top             =   840
                     Width           =   915
                     _Version        =   65536
                     _ExtentX        =   1614
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Appearance      =   1
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sc.Imp"
                     Height          =   195
                     Index           =   18
                     Left            =   1965
                     TabIndex        =   134
                     Top             =   840
                     Width           =   495
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sc.%"
                     Height          =   195
                     Index           =   17
                     Left            =   1965
                     TabIndex        =   133
                     Top             =   450
                     Width           =   390
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Trasporto"
                     Height          =   195
                     Index           =   16
                     Left            =   120
                     TabIndex        =   132
                     Top             =   840
                     Width           =   705
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Incasso"
                     Height          =   195
                     Index           =   15
                     Left            =   120
                     TabIndex        =   131
                     Top             =   360
                     Width           =   555
                  End
               End
               Begin VB.Frame FraTab 
                  Caption         =   "Scadenze"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   1395
                  Index           =   4
                  Left            =   3840
                  TabIndex        =   128
                  Top             =   5880
                  Width           =   2940
                  Begin MSComctlLib.ListView lvwScadenze 
                     Height          =   1050
                     Left            =   120
                     TabIndex        =   129
                     Top             =   240
                     Width           =   2655
                     _ExtentX        =   4683
                     _ExtentY        =   1852
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   0
                     NumItems        =   0
                  End
               End
               Begin VB.Frame FraTab 
                  Caption         =   "Castelletto"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   1395
                  Index           =   2
                  Left            =   120
                  TabIndex        =   126
                  Top             =   5880
                  Width           =   3810
                  Begin MSComctlLib.ListView lvwIVA 
                     Height          =   1050
                     Left            =   120
                     TabIndex        =   127
                     Top             =   240
                     Width           =   3570
                     _ExtentX        =   6297
                     _ExtentY        =   1852
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   0
                     NumItems        =   0
                  End
               End
               Begin VB.Frame FraTab 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   9405
                  Index           =   3
                  Left            =   -74880
                  TabIndex        =   112
                  Top             =   600
                  Width           =   17175
                  Begin VB.CommandButton Command1 
                     Caption         =   "Salva e rettifica"
                     Height          =   285
                     Left            =   15480
                     MaskColor       =   &H00FF0000&
                     TabIndex        =   276
                     Top             =   6720
                     Width           =   1545
                  End
                  Begin VB.CommandButton cmdElencoRett 
                     Caption         =   "Elenco rettifiche"
                     Height          =   285
                     Left            =   15480
                     TabIndex        =   275
                     TabStop         =   0   'False
                     Top             =   8880
                     Width           =   1545
                  End
                  Begin MSComctlLib.ListView lvwArticoli 
                     Height          =   3855
                     Left            =   120
                     TabIndex        =   83
                     TabStop         =   0   'False
                     Top             =   5400
                     Width           =   15285
                     _ExtentX        =   26961
                     _ExtentY        =   6800
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     GridLines       =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   0
                     NumItems        =   0
                  End
                  Begin VB.TextBox txtAndamentoOrdineDett 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H8000000D&
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   177
                     TabStop         =   0   'False
                     Text            =   "100"
                     Top             =   5160
                     Visible         =   0   'False
                     Width           =   15255
                  End
                  Begin VB.Frame Frame2 
                     Caption         =   "Note per riga di lavorazione"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   855
                     Left            =   7560
                     TabIndex        =   261
                     Top             =   4320
                     Width           =   7815
                     Begin VB.TextBox txtAnnotazioniDiRigaLav 
                        Height          =   495
                        Left            =   120
                        MultiLine       =   -1  'True
                        ScrollBars      =   2  'Vertical
                        TabIndex        =   262
                        Top             =   240
                        Width           =   7575
                     End
                  End
                  Begin VB.CommandButton cmdDuplicaRiga 
                     Caption         =   "Duplica riga"
                     Height          =   285
                     Left            =   15480
                     TabIndex        =   248
                     TabStop         =   0   'False
                     Top             =   8160
                     Width           =   1545
                  End
                  Begin VB.Frame fraTotaliPedane 
                     Caption         =   "Totali pedane"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   1575
                     Left            =   15480
                     TabIndex        =   231
                     Top             =   3360
                     Width           =   1575
                     Begin DMTEDITNUMLib.dmtCurrency txtTotalePedane 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   232
                        TabStop         =   0   'False
                        Top             =   480
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   " 0"
                        BackColor       =   12648447
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencySymbol  =   ""
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin DMTEDITNUMLib.dmtCurrency txtTotalePedaneEff 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   233
                        TabStop         =   0   'False
                        Top             =   1080
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   " 0"
                        BackColor       =   12648447
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencySymbol  =   ""
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin VB.Label Label3 
                        Caption         =   "Quantità"
                        Height          =   255
                        Index           =   5
                        Left            =   120
                        TabIndex        =   235
                        Top             =   240
                        Width           =   1215
                     End
                     Begin VB.Label Label3 
                        Caption         =   "Quantità effettiva"
                        Height          =   255
                        Index           =   1
                        Left            =   120
                        TabIndex        =   234
                        Top             =   840
                        Width           =   1335
                     End
                  End
                  Begin VB.Frame fraOrdineRigaOri 
                     Caption         =   "Importo listino"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   1335
                     Left            =   15480
                     TabIndex        =   223
                     Top             =   2160
                     Width           =   1575
                     Begin DMTEDITNUMLib.dmtCurrency txtImportoListinoArticolo 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   224
                        TabStop         =   0   'False
                        Top             =   240
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   " 0"
                        BackColor       =   12648447
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencySymbol  =   ""
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin DMTEDITNUMLib.dmtCurrency txtScontoImpListino 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   225
                        TabStop         =   0   'False
                        Top             =   840
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   " 0"
                        BackColor       =   12648447
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencyDecimalPlaces=   5
                        CurrencySymbol  =   ""
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin VB.Label Label3 
                        Caption         =   "Sconto"
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   226
                        Top             =   600
                        Width           =   1215
                     End
                  End
                  Begin VB.Frame fraImballo 
                     Height          =   1815
                     Left            =   120
                     TabIndex        =   204
                     Top             =   800
                     Width           =   15255
                     Begin VB.CheckBox chkImportoImballoInArticolo 
                        Caption         =   "Prezzo art. compreso imb."
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   375
                        Left            =   9840
                        TabIndex        =   49
                        TabStop         =   0   'False
                        Top             =   360
                        Width           =   2535
                     End
                     Begin VB.TextBox txtDescrizioneImballo 
                        Height          =   315
                        Left            =   1560
                        TabIndex        =   43
                        Top             =   360
                        Width           =   3015
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtImponibileUnitario 
                        Height          =   315
                        Left            =   10920
                        TabIndex        =   67
                        Top             =   1400
                        Width           =   1575
                        _Version        =   65536
                        _ExtentX        =   2778
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   4
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtSconto2 
                        Height          =   315
                        Left            =   10200
                        TabIndex        =   66
                        Top             =   1400
                        Width           =   615
                        _Version        =   65536
                        _ExtentX        =   1085
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtSconto1 
                        Height          =   315
                        Left            =   9480
                        TabIndex        =   65
                        Top             =   1400
                        Width           =   615
                        _Version        =   65536
                        _ExtentX        =   1085
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DmtCodDescCtl.DmtCodDesc CDImballo 
                        Height          =   615
                        Left            =   120
                        TabIndex        =   42
                        Top             =   120
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   1085
                        PropCodice      =   $"frmMain.frx":47A22
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":47A79
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":47AD0
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
                     Begin DMTEDITNUMLib.dmtCurrency txtImportoUnitarioImballo 
                        Height          =   315
                        Left            =   8520
                        TabIndex        =   48
                        Top             =   390
                        Width           =   1215
                        _Version        =   65536
                        _ExtentX        =   2143
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   " 0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencyDecimalPlaces=   5
                        CurrencySymbol  =   ""
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtAliquotaImballo 
                        Height          =   315
                        Left            =   7800
                        TabIndex        =   47
                        TabStop         =   0   'False
                        Top             =   390
                        Width           =   615
                        _Version        =   65536
                        _ExtentX        =   1085
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTDataCmb.DMTCombo CboAliquotaImballo 
                        Height          =   315
                        Left            =   6840
                        TabIndex        =   46
                        TabStop         =   0   'False
                        Top             =   390
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
                     Begin DMTEDITNUMLib.dmtNumber txtTaraUnitaria 
                        Height          =   315
                        Left            =   4680
                        TabIndex        =   44
                        TabStop         =   0   'False
                        Top             =   390
                        Width           =   855
                        _Version        =   65536
                        _ExtentX        =   1508
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   5
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
                        Height          =   315
                        Left            =   8040
                        TabIndex        =   64
                        Top             =   1400
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   5
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtAliquotaArticolo 
                        Height          =   315
                        Left            =   7320
                        TabIndex        =   63
                        TabStop         =   0   'False
                        Top             =   1400
                        Width           =   615
                        _Version        =   65536
                        _ExtentX        =   1085
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTDataCmb.DMTCombo cboAliquotaArticolo 
                        Height          =   315
                        Left            =   6360
                        TabIndex        =   62
                        TabStop         =   0   'False
                        Top             =   1400
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
                     Begin DMTEDITNUMLib.dmtNumber txtQta_UM 
                        Height          =   315
                        Left            =   5160
                        TabIndex        =   61
                        Top             =   1400
                        Width           =   1095
                        _Version        =   65536
                        _ExtentX        =   1931
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPezzi 
                        Height          =   315
                        Left            =   4080
                        TabIndex        =   60
                        Top             =   1400
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   0
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtTara 
                        Height          =   315
                        Left            =   2040
                        TabIndex        =   58
                        Top             =   1400
                        Width           =   855
                        _Version        =   65536
                        _ExtentX        =   1508
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPesoNetto 
                        Height          =   315
                        Left            =   3000
                        TabIndex        =   59
                        Top             =   1400
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPesoLordo 
                        Height          =   315
                        Left            =   960
                        TabIndex        =   57
                        Top             =   1400
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtColli 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   56
                        Top             =   1400
                        Width           =   735
                        _Version        =   65536
                        _ExtentX        =   1296
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtQuantitaPerPedana 
                        Height          =   315
                        Left            =   5640
                        TabIndex        =   45
                        Top             =   390
                        Width           =   1095
                        _Version        =   65536
                        _ExtentX        =   1931
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DmtCodDescCtl.DmtCodDesc CDImballoPrimario 
                        Height          =   615
                        Left            =   120
                        TabIndex        =   50
                        TabStop         =   0   'False
                        Top             =   640
                        Width           =   4530
                        _ExtentX        =   7990
                        _ExtentY        =   1085
                        PropCodice      =   $"frmMain.frx":47B2A
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":47B79
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":47BE1
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtTaraConfImballo 
                        Height          =   315
                        Left            =   4680
                        TabIndex        =   51
                        TabStop         =   0   'False
                        Top             =   900
                        Width           =   855
                        _Version        =   65536
                        _ExtentX        =   1508
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   5
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtNumeroConfImballo 
                        Height          =   315
                        Left            =   5640
                        TabIndex        =   52
                        Top             =   900
                        Width           =   1095
                        _Version        =   65536
                        _ExtentX        =   1931
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   0
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTDataCmb.DMTCombo cboReportNew 
                        Height          =   315
                        Left            =   10800
                        TabIndex        =   254
                        TabStop         =   0   'False
                        Top             =   900
                        Width           =   2055
                        _ExtentX        =   3625
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
                     Begin DMTDataCmb.DMTCombo cboReportPedNew 
                        Height          =   315
                        Left            =   12960
                        TabIndex        =   255
                        TabStop         =   0   'False
                        Top             =   900
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
                     Begin DMTEDITNUMLib.dmtNumber txtQuantitaPerCollo 
                        Height          =   315
                        Left            =   6840
                        TabIndex        =   53
                        Top             =   900
                        Width           =   1455
                        _Version        =   65536
                        _ExtentX        =   2566
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   0
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPesoPerCollo 
                        Height          =   315
                        Left            =   8400
                        TabIndex        =   54
                        Top             =   900
                        Width           =   1095
                        _Version        =   65536
                        _ExtentX        =   1931
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   5
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtMoltiplicatore 
                        Height          =   315
                        Left            =   9600
                        TabIndex        =   55
                        Top             =   900
                        Width           =   1095
                        _Version        =   65536
                        _ExtentX        =   1931
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   5
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtIDRigaContratto 
                        Height          =   255
                        Left            =   14400
                        TabIndex        =   271
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
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà pezzi per collo"
                        Height          =   255
                        Index           =   19
                        Left            =   6840
                        TabIndex        =   260
                        ToolTipText     =   "Numero di confezioni in un imballo"
                        Top             =   700
                        Width           =   1455
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Peso per collo"
                        Height          =   255
                        Index           =   18
                        Left            =   8400
                        TabIndex        =   259
                        ToolTipText     =   "Numero di confezioni in un imballo"
                        Top             =   700
                        Width           =   1095
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Moltiplicatore"
                        Height          =   255
                        Index           =   17
                        Left            =   9600
                        TabIndex        =   258
                        ToolTipText     =   "Numero di confezioni in un imballo"
                        Top             =   700
                        Width           =   1095
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Etichette per pedana"
                        Height          =   255
                        Index           =   16
                        Left            =   12960
                        TabIndex        =   257
                        ToolTipText     =   "Numero di confezioni in un imballo"
                        Top             =   705
                        Width           =   2055
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Etichetta prodotto"
                        Height          =   255
                        Index           =   15
                        Left            =   10800
                        TabIndex        =   256
                        ToolTipText     =   "Numero di confezioni in un imballo"
                        Top             =   705
                        Width           =   1575
                     End
                     Begin VB.Label Label2 
                        Caption         =   "N° confezioni"
                        Height          =   255
                        Index           =   32
                        Left            =   5640
                        TabIndex        =   252
                        ToolTipText     =   "Numero di confezioni in un imballo"
                        Top             =   700
                        Width           =   1095
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Tara"
                        Height          =   255
                        Index           =   13
                        Left            =   4680
                        TabIndex        =   251
                        Top             =   700
                        Width           =   855
                     End
                     Begin VB.Label lblStatoRiga 
                        Alignment       =   2  'Center
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Left            =   12600
                        TabIndex        =   249
                        Top             =   1320
                        Width           =   2535
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Aliq."
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   38
                        Left            =   7800
                        TabIndex        =   222
                        Top             =   180
                        Width           =   315
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Cod. IVA"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   37
                        Left            =   6840
                        TabIndex        =   221
                        Top             =   180
                        Width           =   1005
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Tara"
                        Height          =   255
                        Index           =   6
                        Left            =   4680
                        TabIndex        =   220
                        Top             =   180
                        Width           =   855
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Descrizione imballo"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00800000&
                        Height          =   195
                        Index           =   36
                        Left            =   1560
                        MouseIcon       =   "frmMain.frx":47C3B
                        MousePointer    =   99  'Custom
                        TabIndex        =   219
                        Top             =   180
                        Width           =   1350
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà"
                        Height          =   255
                        Index           =   5
                        Left            =   5160
                        TabIndex        =   218
                        Top             =   1200
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        Caption         =   "P. netto"
                        Height          =   255
                        Index           =   4
                        Left            =   3000
                        TabIndex        =   217
                        Top             =   1200
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Tara "
                        Height          =   255
                        Index           =   3
                        Left            =   2040
                        TabIndex        =   216
                        Top             =   1200
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Pezzi"
                        Height          =   255
                        Index           =   2
                        Left            =   4080
                        TabIndex        =   215
                        Top             =   1200
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        Caption         =   "P. lordo"
                        Height          =   255
                        Index           =   1
                        Left            =   960
                        TabIndex        =   214
                        Top             =   1200
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Colli"
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   213
                        Top             =   1200
                        Width           =   735
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Imp. Unit. Art."
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   31
                        Left            =   8040
                        TabIndex        =   212
                        Top             =   1200
                        Width           =   1365
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Aliq."
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   27
                        Left            =   7320
                        TabIndex        =   211
                        Top             =   1200
                        Width           =   555
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Cod. IVA"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   26
                        Left            =   6360
                        TabIndex        =   210
                        Top             =   1200
                        Width           =   1005
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Imp. Unit. Imb."
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   6
                        Left            =   8520
                        TabIndex        =   209
                        Top             =   180
                        Width           =   1215
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Sc 1"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   4
                        Left            =   9480
                        TabIndex        =   208
                        Top             =   1200
                        Width           =   540
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Sc 2"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   32
                        Left            =   10200
                        TabIndex        =   207
                        Top             =   1200
                        Width           =   540
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Imp. Unit.tot."
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   33
                        Left            =   10920
                        TabIndex        =   206
                        Top             =   1200
                        Width           =   1575
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà X pedana"
                        Height          =   225
                        Index           =   9
                        Left            =   5640
                        TabIndex        =   205
                        Top             =   180
                        Width           =   1095
                     End
                  End
                  Begin VB.Frame fraPedana 
                     Height          =   735
                     Left            =   120
                     TabIndex        =   201
                     Top             =   2520
                     Width           =   15255
                     Begin DMTEDITNUMLib.dmtNumber txtQuantitaPedana 
                        Height          =   315
                        Left            =   9120
                        TabIndex        =   70
                        Top             =   360
                        Width           =   1095
                        _Version        =   65536
                        _ExtentX        =   1931
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTDataCmb.DMTCombo cboUMRigaOrdine 
                        Height          =   315
                        Left            =   12480
                        TabIndex        =   73
                        Top             =   360
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
                        Left            =   120
                        TabIndex        =   68
                        Top             =   120
                        Width           =   5535
                        _ExtentX        =   9763
                        _ExtentY        =   1085
                        PropCodice      =   $"frmMain.frx":47F45
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":47FA0
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":48003
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
                     Begin DmtCodDescCtl.DmtCodDesc CDPedana 
                        Height          =   615
                        Left            =   5760
                        TabIndex        =   69
                        Top             =   120
                        Width           =   3375
                        _ExtentX        =   5953
                        _ExtentY        =   1085
                        PropCodice      =   $"frmMain.frx":4805D
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":480AC
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":48100
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtQuantitaPedanaEff 
                        Height          =   315
                        Left            =   11400
                        TabIndex        =   72
                        Top             =   360
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtColliSfusi 
                        Height          =   315
                        Left            =   10320
                        TabIndex        =   71
                        Top             =   360
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Colli sfusi"
                        Height          =   225
                        Index           =   14
                        Left            =   10320
                        TabIndex        =   253
                        Top             =   150
                        Width           =   855
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà eff."
                        Height          =   225
                        Index           =   12
                        Left            =   11400
                        TabIndex        =   230
                        Top             =   165
                        Width           =   975
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Unità di misura riga di ordine"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   25
                        Left            =   12480
                        TabIndex        =   203
                        Top             =   165
                        Width           =   2625
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà"
                        Height          =   225
                        Index           =   7
                        Left            =   9120
                        TabIndex        =   202
                        Top             =   150
                        Width           =   1095
                     End
                  End
                  Begin VB.Frame fraAnnotazioni 
                     Caption         =   "Annotazioni"
                     BeginProperty Font 
                        Name            =   "Tahoma"
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
                     TabIndex        =   199
                     Top             =   4320
                     Width           =   7455
                     Begin VB.TextBox txtAnnotazioniDiRiga 
                        Height          =   495
                        Left            =   120
                        MultiLine       =   -1  'True
                        ScrollBars      =   2  'Vertical
                        TabIndex        =   200
                        Top             =   240
                        Width           =   7215
                     End
                  End
                  Begin VB.Frame fraVivaio 
                     Height          =   1335
                     Left            =   120
                     TabIndex        =   196
                     Top             =   3120
                     Width           =   15255
                     Begin VB.TextBox txtCollegamentoRigaOrdine 
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
                        Height          =   315
                        Left            =   6840
                        Locked          =   -1  'True
                        TabIndex        =   236
                        Top             =   880
                        Width           =   8295
                     End
                     Begin DmtCodDescCtl.DmtCodDesc CDArticoloPianale 
                        Height          =   615
                        Left            =   120
                        TabIndex        =   74
                        Top             =   90
                        Width           =   5535
                        _ExtentX        =   9763
                        _ExtentY        =   1085
                        PropCodice      =   $"frmMain.frx":4815A
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":481B1
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":48210
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtQtaPianale 
                        Height          =   315
                        Left            =   5760
                        TabIndex        =   75
                        Top             =   340
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   0
                        AllowEmpty      =   0   'False
                     End
                     Begin DmtCodDescCtl.DmtCodDesc CDArticoloProlunga 
                        Height          =   615
                        Left            =   120
                        TabIndex        =   76
                        Top             =   640
                        Width           =   5535
                        _ExtentX        =   9763
                        _ExtentY        =   1085
                        PropCodice      =   $"frmMain.frx":4826A
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":482C2
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":48322
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtQtaProlunga 
                        Height          =   315
                        Left            =   5760
                        TabIndex        =   77
                        Top             =   880
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecimalPlaces   =   0
                        AllowEmpty      =   0   'False
                     End
                     Begin DmtSearchAccount.DmtSearchACS ACSSocio 
                        Height          =   585
                        Left            =   6840
                        TabIndex        =   78
                        Top             =   120
                        Width           =   5625
                        _ExtentX        =   9922
                        _ExtentY        =   1032
                        WidthDescription=   3500
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
                        CaptionDescription=   "Socio/Fornitore"
                        CaptionCode     =   "Codice"
                        OnlyAccounts    =   -1  'True
                     End
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Quantità"
                        Height          =   255
                        Index           =   11
                        Left            =   5760
                        TabIndex        =   198
                        Top             =   670
                        Width           =   975
                     End
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Quantità"
                        Height          =   255
                        Index           =   10
                        Left            =   5760
                        TabIndex        =   197
                        Top             =   150
                        Width           =   975
                     End
                     Begin VB.Label Label5 
                        Caption         =   "Conferimento collegato"
                        Height          =   255
                        Index           =   4
                        Left            =   6840
                        TabIndex        =   237
                        Top             =   675
                        Width           =   2535
                     End
                  End
                  Begin VB.Frame fraArticoli 
                     Height          =   3015
                     Left            =   120
                     TabIndex        =   151
                     Top             =   120
                     Width           =   15255
                     Begin VB.TextBox txtRaggrRigaOrdine 
                        Height          =   315
                        Left            =   12600
                        TabIndex        =   41
                        Top             =   360
                        Width           =   2535
                     End
                     Begin VB.TextBox txtDescrizioneArticolo 
                        Height          =   315
                        Left            =   1560
                        TabIndex        =   36
                        Top             =   375
                        Width           =   3015
                     End
                     Begin DMTDataCmb.DMTCombo cboTipoLavorazione 
                        Height          =   315
                        Left            =   9960
                        TabIndex        =   40
                        TabStop         =   0   'False
                        Top             =   360
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
                     Begin DMTDataCmb.DMTCombo cboTipoCategoria 
                        Height          =   315
                        Left            =   7680
                        TabIndex        =   39
                        TabStop         =   0   'False
                        Top             =   375
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
                     Begin DMTDataCmb.DMTCombo cboCalibro 
                        Height          =   315
                        Left            =   6000
                        TabIndex        =   38
                        TabStop         =   0   'False
                        Top             =   375
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
                     Begin DMTDataCmb.DMTCombo cboUnitaDiMisura 
                        Height          =   315
                        Left            =   4680
                        TabIndex        =   37
                        TabStop         =   0   'False
                        Top             =   375
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
                     Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
                        Height          =   615
                        Left            =   120
                        TabIndex        =   35
                        Top             =   120
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   1085
                        PropCodice      =   $"frmMain.frx":4837C
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":483D4
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":4842B
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
                     Begin VB.Label Label13 
                        Caption         =   "Sub lotto"
                        Height          =   195
                        Index           =   1
                        Left            =   12600
                        TabIndex        =   229
                        Top             =   165
                        Width           =   1935
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Tipo di lavorazione"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   35
                        Left            =   9960
                        TabIndex        =   169
                        Top             =   165
                        Width           =   1335
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Categoria"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   14
                        Left            =   7680
                        TabIndex        =   155
                        Top             =   165
                        Width           =   705
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Calibro"
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   13
                        Left            =   6000
                        TabIndex        =   154
                        Top             =   165
                        Width           =   495
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Descrizione articolo"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00800000&
                        Height          =   195
                        Index           =   24
                        Left            =   1560
                        MouseIcon       =   "frmMain.frx":48485
                        MousePointer    =   99  'Custom
                        TabIndex        =   153
                        Top             =   165
                        Width           =   2700
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "U. M."
                        ForeColor       =   &H00000000&
                        Height          =   195
                        Index           =   28
                        Left            =   4680
                        TabIndex        =   152
                        Top             =   165
                        Width           =   975
                     End
                  End
                  Begin VB.Frame Frame10 
                     Height          =   2175
                     Left            =   15480
                     TabIndex        =   138
                     Top             =   120
                     Width           =   1575
                     Begin DMTEDITNUMLib.dmtCurrency txtImponibileImballo 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   139
                        Top             =   780
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "€ 0"
                        BackColor       =   12648447
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencySymbol  =   "€"
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin DMTEDITNUMLib.dmtCurrency txtTotaleImponibile 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   140
                        TabStop         =   0   'False
                        Top             =   1250
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "€ 0"
                        BackColor       =   12648447
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencySymbol  =   "€"
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin DMTEDITNUMLib.dmtCurrency txtImponibileArticolo 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   141
                        TabStop         =   0   'False
                        Top             =   300
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "€ 0"
                        BackColor       =   12648447
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencySymbol  =   "€"
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin DMTEDITNUMLib.dmtCurrency txtTotaleRiga 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   142
                        TabStop         =   0   'False
                        Top             =   1720
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "€ 0"
                        BackColor       =   12648447
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        CurrencySymbol  =   "€"
                        AllowEmpty      =   0   'False
                        DecFinalZeros   =   -1  'True
                     End
                     Begin VB.Label Label3 
                        Caption         =   "Totale riga"
                        Height          =   255
                        Index           =   4
                        Left            =   120
                        TabIndex        =   146
                        Top             =   1520
                        Width           =   1215
                     End
                     Begin VB.Label Label3 
                        Caption         =   "Imp. Imballo"
                        Height          =   255
                        Index           =   2
                        Left            =   120
                        TabIndex        =   145
                        Top             =   580
                        Width           =   1215
                     End
                     Begin VB.Label Label3 
                        Caption         =   "Tot. imponibile"
                        Height          =   255
                        Index           =   3
                        Left            =   120
                        TabIndex        =   144
                        Top             =   1050
                        Width           =   1215
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Imp. Articolo"
                        Height          =   255
                        Index           =   8
                        Left            =   120
                        TabIndex        =   143
                        Top             =   120
                        Width           =   1215
                     End
                  End
                  Begin VB.CommandButton cmdElimina 
                     Caption         =   "Elimina"
                     Height          =   285
                     Left            =   15480
                     TabIndex        =   82
                     TabStop         =   0   'False
                     Top             =   7440
                     Width           =   1545
                  End
                  Begin VB.CommandButton cmdSalva 
                     Caption         =   "&Salva"
                     Height          =   285
                     Left            =   15480
                     MaskColor       =   &H00FF0000&
                     TabIndex        =   79
                     Top             =   6000
                     Width           =   1545
                  End
                  Begin VB.CommandButton cmdNuovo 
                     Caption         =   "&Nuovo"
                     Height          =   285
                     Left            =   15480
                     TabIndex        =   80
                     Top             =   5280
                     Width           =   1545
                  End
                  Begin VB.Frame Frame8 
                     Height          =   2895
                     Left            =   12960
                     TabIndex        =   113
                     Top             =   6960
                     Visible         =   0   'False
                     Width           =   1575
                     Begin DMTEDITNUMLib.dmtNumber txtQtaConferita 
                        Height          =   255
                        Left            =   120
                        TabIndex        =   114
                        Top             =   360
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   450
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   0
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtColliVenduti_Conferimento 
                        Height          =   255
                        Left            =   120
                        TabIndex        =   115
                        Top             =   1250
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   450
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   0
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtQtaVenduta_conferimento 
                        Height          =   255
                        Left            =   120
                        TabIndex        =   116
                        Top             =   1700
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   450
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   0
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtQtaQuadrata_Conferimento 
                        Height          =   255
                        Left            =   120
                        TabIndex        =   117
                        Top             =   2100
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   450
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   0
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtDifferenzaConferimento 
                        Height          =   255
                        Left            =   120
                        TabIndex        =   118
                        Top             =   2520
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   450
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   0
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtColliConferiti 
                        Height          =   255
                        Left            =   120
                        TabIndex        =   119
                        Top             =   800
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   450
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   0
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                        UseSeparator    =   -1  'True
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin VB.Label Label14 
                        Caption         =   "Q.tà Conferita"
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   125
                        Top             =   180
                        Width           =   1335
                     End
                     Begin VB.Label Label14 
                        Caption         =   "Colli venduti"
                        Height          =   255
                        Index           =   1
                        Left            =   120
                        TabIndex        =   124
                        Top             =   1060
                        Width           =   1335
                     End
                     Begin VB.Label Label14 
                        Caption         =   "Q.tà venduta"
                        Height          =   255
                        Index           =   2
                        Left            =   120
                        TabIndex        =   123
                        Top             =   1500
                        Width           =   1335
                     End
                     Begin VB.Label Label14 
                        Caption         =   "Q.tà quadratura"
                        Height          =   255
                        Index           =   3
                        Left            =   120
                        TabIndex        =   122
                        Top             =   1920
                        Width           =   1335
                     End
                     Begin VB.Label Label14 
                        Caption         =   "Differenza"
                        Height          =   255
                        Index           =   4
                        Left            =   120
                        TabIndex        =   121
                        Top             =   2360
                        Width           =   1335
                     End
                     Begin VB.Label Label14 
                        Caption         =   "Colli conferiti"
                        Height          =   255
                        Index           =   5
                        Left            =   120
                        TabIndex        =   120
                        Top             =   620
                        Width           =   1335
                     End
                  End
                  Begin VB.Frame fraAnnotazioniPerSocio 
                     Caption         =   "Annotazioni per socio"
                     BeginProperty Font 
                        Name            =   "Tahoma"
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
                     TabIndex        =   227
                     Top             =   4320
                     Width           =   15255
                     Begin VB.TextBox txtAnnotazioniPerSocio 
                        Height          =   495
                        Left            =   120
                        MultiLine       =   -1  'True
                        ScrollBars      =   2  'Vertical
                        TabIndex        =   228
                        Top             =   240
                        Width           =   15015
                     End
                  End
               End
               Begin DMTDATETIMELib.dmtDate txtDataCambio 
                  Height          =   255
                  Left            =   14640
                  TabIndex        =   148
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   2535
                  _Version        =   65536
                  _ExtentX        =   4471
                  _ExtentY        =   450
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtValoreCambioValuta 
                  Height          =   255
                  Left            =   12120
                  TabIndex        =   149
                  Top             =   3240
                  Visible         =   0   'False
                  Width           =   2535
                  _Version        =   65536
                  _ExtentX        =   4471
                  _ExtentY        =   450
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboCambioValuta 
                  Height          =   315
                  Left            =   14640
                  TabIndex        =   150
                  Top             =   3480
                  Visible         =   0   'False
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
               Begin VB.Frame FraTab 
                  Caption         =   "Documento"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   5475
                  Index           =   0
                  Left            =   120
                  TabIndex        =   156
                  Top             =   540
                  Width           =   5295
                  Begin VB.TextBox txtCausaleDocumentoEF 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   273
                     TabStop         =   0   'False
                     Top             =   1920
                     Width           =   2415
                  End
                  Begin VB.CheckBox chkOrdineCompletato 
                     Caption         =   "Ordine completato"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   120
                     TabIndex        =   245
                     TabStop         =   0   'False
                     Top             =   4605
                     Width           =   2295
                  End
                  Begin VB.Frame fraOrdinePadre 
                     Caption         =   "Riferimento ordine padre"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   1335
                     Left            =   2760
                     TabIndex        =   238
                     Top             =   3480
                     Width           =   2415
                     Begin DMTDATETIMELib.dmtDate txtDataDocPadre 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   239
                        Top             =   450
                        Width           =   2145
                        _Version        =   65536
                        _ExtentX        =   3784
                        _ExtentY        =   556
                        _StockProps     =   253
                        BackColor       =   16777215
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtNOrdinePadre 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   240
                        Top             =   960
                        Width           =   1290
                        _Version        =   65536
                        _ExtentX        =   2275
                        _ExtentY        =   556
                        _StockProps     =   253
                        BackColor       =   16777215
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtNListaPrelievo 
                        Height          =   315
                        Left            =   1440
                        TabIndex        =   243
                        Top             =   960
                        Width           =   810
                        _Version        =   65536
                        _ExtentX        =   1429
                        _ExtentY        =   556
                        _StockProps     =   253
                        BackColor       =   16777215
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Enabled         =   0   'False
                        Appearance      =   1
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "N° lista"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Index           =   40
                        Left            =   1440
                        TabIndex        =   244
                        Top             =   765
                        Width           =   840
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Numero"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Index           =   29
                        Left            =   120
                        TabIndex        =   242
                        Top             =   765
                        Width           =   1260
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "Data"
                        BeginProperty Font 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Index           =   9
                        Left            =   120
                        TabIndex        =   241
                        Top             =   240
                        Width           =   2085
                     End
                  End
                  Begin VB.TextBox txtAndamentoOrdine 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H8000000D&
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   180
                     Left            =   195
                     Locked          =   -1  'True
                     TabIndex        =   175
                     TabStop         =   0   'False
                     Text            =   "100"
                     Top             =   5075
                     Visible         =   0   'False
                     Width           =   4935
                  End
                  Begin VB.TextBox txtNumeroOrdineCliente 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   1800
                     TabIndex        =   7
                     Top             =   1440
                     Width           =   1575
                  End
                  Begin VB.TextBox txtCausaleDocumento 
                     Height          =   315
                     Left            =   3240
                     TabIndex        =   9
                     TabStop         =   0   'False
                     Top             =   3600
                     Visible         =   0   'False
                     Width           =   1575
                  End
                  Begin VB.CheckBox chkRaggruppaScadenze 
                     Caption         =   "Raggruppa scadenze"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   14
                     TabStop         =   0   'False
                     Top             =   4005
                     Width           =   2295
                  End
                  Begin VB.CheckBox chkRaggruppBolle 
                     Caption         =   "Raggruppa bolle"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   13
                     TabStop         =   0   'False
                     Top             =   3720
                     Width           =   2295
                  End
                  Begin VB.CheckBox chkLordoIVA 
                     Caption         =   "Prezzi lordo IVA"
                     Height          =   240
                     Left            =   120
                     TabIndex        =   12
                     TabStop         =   0   'False
                     Top             =   3420
                     Width           =   2265
                  End
                  Begin VB.CheckBox chkChiuso 
                     Caption         =   "Ordine chiuso"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   120
                     TabIndex        =   15
                     TabStop         =   0   'False
                     Top             =   4305
                     Width           =   2295
                  End
                  Begin DMTDataCmb.DMTCombo cboBancaAzienda 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   4
                     TabStop         =   0   'False
                     Top             =   1920
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
                  Begin DMTDataCmb.DMTCombo cboSezionale 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   0
                     Top             =   405
                     Width           =   2295
                     _ExtentX        =   4048
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
                  Begin DMTDataCmb.DMTCombo cboMagazzino 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   3
                     TabStop         =   0   'False
                     Top             =   900
                     Width           =   1935
                     _ExtentX        =   3413
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
                  Begin DMTDataCmb.DMTCombo cboListino 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   10
                     TabStop         =   0   'False
                     Top             =   2400
                     Width           =   2520
                     _ExtentX        =   4445
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
                  Begin DMTDATETIMELib.dmtDate dtData 
                     Height          =   315
                     Left            =   2520
                     TabIndex        =   1
                     Top             =   405
                     Width           =   1305
                     _Version        =   65536
                     _ExtentX        =   2302
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTEDITNUMLib.dmtNumber lngNumero 
                     Height          =   315
                     Left            =   3960
                     TabIndex        =   2
                     Top             =   405
                     Width           =   1170
                     _Version        =   65536
                     _ExtentX        =   2064
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboListinoAzienda 
                     Height          =   315
                     Left            =   2760
                     TabIndex        =   11
                     TabStop         =   0   'False
                     Top             =   2400
                     Width           =   2400
                     _ExtentX        =   4233
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
                  Begin DMTDATETIMELib.dmtDate txtDataOrdineCliente 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   6
                     Top             =   1440
                     Width           =   1575
                     _Version        =   65536
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataPartenza 
                     Height          =   315
                     Left            =   3480
                     TabIndex        =   8
                     Top             =   1440
                     Width           =   1695
                     _Version        =   65536
                     _ExtentX        =   2990
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboTipoOrdine 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   170
                     TabStop         =   0   'False
                     Top             =   900
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
                  Begin DMTDataCmb.DMTCombo cboAccordoCommerciale 
                     Height          =   315
                     Left            =   2760
                     TabIndex        =   189
                     Top             =   2900
                     Width           =   2415
                     _ExtentX        =   4260
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
                  Begin MSComctlLib.ProgressBar PBOrdine 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   176
                     ToolTipText     =   "Andamento ordine"
                     Top             =   5040
                     Visible         =   0   'False
                     Width           =   5055
                     _ExtentX        =   8916
                     _ExtentY        =   450
                     _Version        =   393216
                     BorderStyle     =   1
                     Appearance      =   0
                     Scrolling       =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboRaggrFatturato 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   267
                     Top             =   2900
                     Width           =   2415
                     _ExtentX        =   4260
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
                  Begin VB.Label Label6 
                     Caption         =   "Causale documento"
                     Height          =   255
                     Index           =   4
                     Left            =   120
                     TabIndex        =   274
                     Top             =   1740
                     Width           =   1935
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Raggruppamento fatturato"
                     Height          =   255
                     Index           =   3
                     Left            =   120
                     TabIndex        =   268
                     Top             =   2700
                     Width           =   2415
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Accordo commerciale"
                     Height          =   255
                     Index           =   2
                     Left            =   2760
                     TabIndex        =   190
                     Top             =   2700
                     Width           =   1935
                  End
                  Begin VB.Label Label13 
                     Caption         =   "100"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   180
                     Index           =   0
                     Left            =   6360
                     TabIndex        =   174
                     Top             =   -120
                     Width           =   375
                  End
                  Begin VB.Line Line1 
                     X1              =   2640
                     X2              =   2640
                     Y1              =   2880
                     Y2              =   4800
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sezionale"
                     Height          =   195
                     Index           =   39
                     Left            =   120
                     TabIndex        =   171
                     Top             =   200
                     Width           =   675
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Data partenza"
                     Height          =   255
                     Index           =   6
                     Left            =   3480
                     TabIndex        =   166
                     Top             =   1240
                     Width           =   1455
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Num. ordine cli."
                     Height          =   255
                     Index           =   3
                     Left            =   1800
                     TabIndex        =   165
                     Top             =   1240
                     Width           =   1455
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Data ordine cli."
                     Height          =   255
                     Index           =   5
                     Left            =   120
                     TabIndex        =   164
                     Top             =   1240
                     Width           =   1455
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Listino azienda"
                     Height          =   195
                     Index           =   30
                     Left            =   2760
                     TabIndex        =   163
                     Top             =   2205
                     Width           =   1650
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Banca azienda"
                     Height          =   255
                     Index           =   0
                     Left            =   2640
                     TabIndex        =   162
                     Top             =   1740
                     Width           =   2535
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Tipo ordine"
                     Height          =   195
                     Index           =   5
                     Left            =   2160
                     TabIndex        =   161
                     Top             =   720
                     Width           =   795
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Magazzino"
                     Height          =   195
                     Index           =   3
                     Left            =   120
                     TabIndex        =   160
                     Top             =   720
                     Width           =   1935
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Data Doc."
                     Height          =   195
                     Index           =   0
                     Left            =   2520
                     TabIndex        =   159
                     Top             =   195
                     Width           =   720
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Numero"
                     Height          =   195
                     Index           =   2
                     Left            =   3960
                     TabIndex        =   158
                     Top             =   195
                     Width           =   1155
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Listino cliente"
                     Height          =   195
                     Index           =   1
                     Left            =   120
                     TabIndex        =   157
                     Top             =   2205
                     Width           =   960
                  End
               End
               Begin VB.Frame Frame1 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3015
                  Left            =   5400
                  TabIndex        =   89
                  Top             =   3000
                  Width           =   5805
                  Begin VB.TextBox txtReferenteAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   94
                     TabStop         =   0   'False
                     Top             =   705
                     Width           =   4695
                  End
                  Begin VB.TextBox txtIndirizzoAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   93
                     TabStop         =   0   'False
                     Top             =   975
                     Width           =   4695
                  End
                  Begin VB.TextBox txtCapAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   92
                     TabStop         =   0   'False
                     Top             =   1245
                     Width           =   855
                  End
                  Begin VB.TextBox txtComuneAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1920
                     TabIndex        =   91
                     TabStop         =   0   'False
                     Top             =   1245
                     Width           =   2775
                  End
                  Begin VB.TextBox txtProvinciaAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   4680
                     TabIndex        =   90
                     TabStop         =   0   'False
                     Top             =   1245
                     Width           =   1095
                  End
                  Begin DMTDataCmb.DMTCombo cboAltroSito 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   20
                     Top             =   360
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
                  Begin DMTDATETIMELib.dmtTime txtOraTrasporto 
                     Height          =   315
                     Left            =   4440
                     TabIndex        =   22
                     Top             =   360
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataTrasporto 
                     Height          =   315
                     Left            =   3120
                     TabIndex        =   21
                     Top             =   360
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboLuogoPresaMerce 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   183
                     Top             =   1920
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
                  Begin DMTDATETIMELib.dmtTime txtOraArrivoLuogo 
                     Height          =   315
                     Left            =   4800
                     TabIndex        =   184
                     Top             =   1920
                     Width           =   855
                     _Version        =   65536
                     _ExtentX        =   1508
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataArrivoLuogo 
                     Height          =   315
                     Left            =   3600
                     TabIndex        =   185
                     Top             =   1920
                     Width           =   1095
                     _Version        =   65536
                     _ExtentX        =   1931
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDAgenteTesta 
                     Height          =   585
                     Left            =   120
                     TabIndex        =   191
                     Top             =   2325
                     Width           =   5625
                     _ExtentX        =   9922
                     _ExtentY        =   1032
                     PropCodice      =   $"frmMain.frx":4878F
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":487E0
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":48831
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin VB.Line Line5 
                     X1              =   120
                     X2              =   5640
                     Y1              =   2280
                     Y2              =   2280
                  End
                  Begin VB.Line Line4 
                     X1              =   120
                     X2              =   5640
                     Y1              =   1605
                     Y2              =   1605
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Luogo di presa merce"
                     Height          =   255
                     Index           =   10
                     Left            =   120
                     TabIndex        =   188
                     Top             =   1680
                     Width           =   1815
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Data arrivo"
                     Height          =   255
                     Index           =   9
                     Left            =   3600
                     TabIndex        =   187
                     Top             =   1680
                     Width           =   975
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Ora arrivo"
                     Height          =   255
                     Index           =   10
                     Left            =   4800
                     TabIndex        =   186
                     Top             =   1680
                     Width           =   735
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Ora arrivo"
                     Height          =   255
                     Index           =   7
                     Left            =   4440
                     TabIndex        =   173
                     Top             =   120
                     Width           =   735
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Data arrivo"
                     Height          =   255
                     Index           =   8
                     Left            =   3120
                     TabIndex        =   172
                     Top             =   120
                     Width           =   975
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Referente"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   97
                     Top             =   720
                     Width           =   975
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Indirizzo"
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   96
                     Top             =   960
                     Width           =   975
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Altra destinazione"
                     Height          =   255
                     Index           =   2
                     Left            =   120
                     TabIndex        =   95
                     Top             =   120
                     Width           =   2895
                  End
               End
               Begin VB.Frame FraTab 
                  Caption         =   "Cliente"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   2535
                  Index           =   1
                  Left            =   5400
                  TabIndex        =   103
                  Top             =   540
                  Width           =   5805
                  Begin VB.TextBox txtNLetteraIntento 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   840
                     TabIndex        =   309
                     Top             =   2040
                     Width           =   495
                  End
                  Begin DmtCodDescCtl.DmtCodDesc cdAnagrafica 
                     Height          =   585
                     Left            =   120
                     TabIndex        =   5
                     Top             =   840
                     Visible         =   0   'False
                     Width           =   5625
                     _ExtentX        =   9922
                     _ExtentY        =   1032
                     PropCodice      =   $"frmMain.frx":4888B
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":488EF
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":48940
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Enabled         =   0   'False
                  End
                  Begin VB.CommandButton cmdLetteraIntento 
                     Height          =   315
                     Left            =   480
                     Picture         =   "frmMain.frx":4899A
                     Style           =   1  'Graphical
                     TabIndex        =   181
                     ToolTipText     =   "Lettere di intento del cliente"
                     Top             =   2040
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdEliminaRifLetInt 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":48F24
                     Style           =   1  'Graphical
                     TabIndex        =   178
                     ToolTipText     =   "Elimina riferimento lettera intento"
                     Top             =   2040
                     Width           =   375
                  End
                  Begin VB.TextBox txtProvincia 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   3870
                     Locked          =   -1  'True
                     TabIndex        =   108
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   1855
                  End
                  Begin VB.TextBox txtComune 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1125
                     Locked          =   -1  'True
                     TabIndex        =   107
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   2730
                  End
                  Begin VB.TextBox txtCAP 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   106
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   1005
                  End
                  Begin VB.TextBox txtIndirizzo 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1560
                     Locked          =   -1  'True
                     TabIndex        =   105
                     TabStop         =   0   'False
                     Top             =   800
                     Width           =   4170
                  End
                  Begin VB.TextBox txtPartitaIva 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   104
                     TabStop         =   0   'False
                     Top             =   800
                     Width           =   1425
                  End
                  Begin DMTDataCmb.DMTCombo cboIvaCliente 
                     Height          =   315
                     Left            =   2520
                     TabIndex        =   18
                     TabStop         =   0   'False
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
                  Begin DMTDataCmb.DMTCombo cboBancaCliente 
                     Height          =   315
                     Left            =   3120
                     TabIndex        =   17
                     Top             =   1550
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
                  Begin DMTDataCmb.DMTCombo cboPagamento 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   16
                     Top             =   1550
                     Width           =   2880
                     _ExtentX        =   5080
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
                  Begin DMTDataCmb.DMTCombo cboValuta 
                     Height          =   315
                     Left            =   4320
                     TabIndex        =   19
                     Top             =   2040
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
                  Begin DMTEDITNUMLib.dmtNumber txtIDLetteraIntento 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   179
                     Top             =   1845
                     Visible         =   0   'False
                     Width           =   375
                     _Version        =   65536
                     _ExtentX        =   661
                     _ExtentY        =   450
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataLetteraIntento 
                     Height          =   315
                     Left            =   1320
                     TabIndex        =   180
                     Top             =   2040
                     Width           =   1095
                     _Version        =   65536
                     _ExtentX        =   1931
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                  End
                  Begin DmtSearchAccount.DmtSearchACS ACSCliente 
                     Height          =   585
                     Left            =   120
                     TabIndex        =   195
                     Top             =   240
                     Width           =   5565
                     _ExtentX        =   9816
                     _ExtentY        =   1032
                     WidthCode       =   900
                     WidthDescription=   3150
                     WidthSecondDescription=   1400
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
                     IDSearchTypeConto=   6
                     OnlyAccounts    =   -1  'True
                  End
                  Begin VB.Label lblLetteraIntento 
                     Caption         =   "Lettera d'intento"
                     Height          =   255
                     Left            =   840
                     TabIndex        =   182
                     Top             =   1860
                     Width           =   1575
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Valuta"
                     Height          =   255
                     Index           =   3
                     Left            =   4320
                     TabIndex        =   147
                     Top             =   1860
                     Width           =   1455
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Pagamento"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   111
                     Top             =   1350
                     Width           =   2775
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Banca"
                     Height          =   255
                     Index           =   1
                     Left            =   3120
                     TabIndex        =   110
                     Top             =   1350
                     Width           =   2655
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Esenzione I.V.A."
                     Height          =   255
                     Index           =   2
                     Left            =   2520
                     TabIndex        =   109
                     Top             =   1860
                     Width           =   1695
                  End
               End
               Begin VB.Frame fraContratto 
                  Caption         =   "Riferimento contratto"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   1215
                  Left            =   11160
                  TabIndex        =   263
                  Top             =   2040
                  Width           =   6135
                  Begin VB.CheckBox chkConfDaContratto 
                     Caption         =   "Conferma della presa in visione delle note del contratto e le note delle righe del documento che provengono da un contratto"
                     Height          =   495
                     Left            =   480
                     TabIndex        =   272
                     Top             =   600
                     Width           =   5535
                  End
                  Begin VB.CommandButton cmdEliminaRifContratto 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":494AE
                     Style           =   1  'Graphical
                     TabIndex        =   266
                     ToolTipText     =   "Elimina riferimento contratto"
                     Top             =   240
                     Width           =   375
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIDContratto 
                     Height          =   255
                     Left            =   4680
                     TabIndex        =   265
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   450
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.TextBox txtDescrContratto 
                     Height          =   315
                     Left            =   480
                     Locked          =   -1  'True
                     TabIndex        =   264
                     Top             =   240
                     Width           =   5535
                  End
               End
               Begin VB.Frame fraAnaDest 
                  Caption         =   "Anagrafica di destinazione"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   1695
                  Left            =   11180
                  TabIndex        =   246
                  Top             =   540
                  Width           =   6135
                  Begin VB.CheckBox chkStampaFattProForma 
                     Caption         =   "Fattura Proforma inviata"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   120
                     TabIndex        =   308
                     TabStop         =   0   'False
                     Top             =   960
                     Width           =   5865
                  End
                  Begin DmtSearchAccount.DmtSearchACS ACSAnaDest 
                     Height          =   585
                     Left            =   120
                     TabIndex        =   247
                     Top             =   240
                     Width           =   5925
                     _ExtentX        =   10451
                     _ExtentY        =   1032
                     WidthCode       =   700
                     WidthDescription=   3600
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
                     CaptionDescription=   "Anagrafica di destinazione"
                     CaptionCode     =   "Codice"
                     IDSearchTypeConto=   6
                     OnlyAccounts    =   -1  'True
                  End
               End
               Begin VB.Label lblInfoTesta 
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   135
                  Top             =   360
                  Width           =   14535
               End
            End
         End
         Begin DmtGridCtl.DmtGrid BrwMain 
            Height          =   2415
            Left            =   0
            TabIndex        =   250
            Top             =   0
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   4260
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnsHeaderHeight=   20
         End
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   9075
         Left            =   0
         TabIndex        =   137
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   16007
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
      Begin DMTWheelCtrl.SpareWheel SpareWheel 
         Left            =   120
         Top             =   1380
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   480
         MousePointer    =   9  'Size W E
         Top             =   180
         Width           =   60
      End
   End
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   81
      Top             =   11610
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Oggetto report
Private oReport As dmtReportLib.dmtReport
'L'applicazione corrente
Private WithEvents m_App As DMTRunAppLib.Application
Attribute m_App.VB_VarHelpID = -1
'Il processo corrente
Private m_Process As DMTRunAppLib.Process
'Il tipo di documento corrente
Private m_DocType As DmtDocManLib.DBFormDocType
'Il documento corrente
Private WithEvents m_Document As DmtDocManLib.DBFormDocument
Attribute m_Document.VB_VarHelpID = -1
'La vista tabellare attiva
Private m_ActiveTableView As DmtDocManLib.TableView
'Il filtro attivo
Private m_ActiveFilter As DmtDocManLib.Filter
'Il report da stampare
Private m_Report As DmtDocManLib.Report
'La variabile  m_Semaphore mantiene un riferimento all'oggetto
'Semaphore che gestisce i conflitti di multiutenza
Private m_Semaphore As Semaforo.dmtSemaphore
'Indica se all'evento KeyPress del Form il tasto deve essere annullato
Private m_EatKey As Boolean
'Indica se l'utente ha modificato uno dei campi del documento
Private m_Changed As Boolean
'Indica se i valori dei campi del documento sono stati salvati
Private m_Saved As Boolean
'Indica se è in corso la definizione di una ricerca
Private m_Search As Boolean
'Indica se uno dei filtri è stato selezionato
Private m_FilterSelected As Boolean
'Indica lo stato di visibilità della vista tabellare
'prima dell'inizio della fase di esecuzione della
'anteprima di stampa
Private m_TabMode As Boolean
'Indica se si sta muovendo lo splitter
Private m_SplitterMoving As Boolean
'Nome dell'eventuale database esteso
Private m_ExtendedDatabase As String
'Processo "Shell su evento OnSave" nome del campo collegato
Private m_LinkedField As String
'Handle della finestra della anteprima di stampa
Private m_PreviewWindowHandle As Long
'Flag che permette l'esecuzione di Form_Activate soltanto all'avvio del programma
Private m_bOnFirstTime As Boolean
'Impedisce il Reposition della browse.
Private m_bAvoidReposition As Boolean
'Consente l'esecuzione del codice contenuto in BrwMain_OnChangeGuiMode()
Private bEnableGuiEvent As Boolean
'Indica se è stato attivato un link
Private m_LinkActive As Boolean

''Oggetto adibito alla gestione del processo On_Extend
'Private m_ExtendApplication As DmtExtendAppLib.ExtendApplication

'Costanti che rappresentano le modalità di visualizzazione
Private Enum neVisualModality
    Insert          'Modalità INSERIMENTO
    Modify          'Modalità VARIAZIONE
    Find            'Modalità TROVA
    Browse          'Modalità ELENCO
    Preview         'Modalità ANTEPRIMA
End Enum

'Costanti usate da SetStatus4Modality per l'apertura/chiusura dell'anteprima di stampa
Private Enum nePreviewModality
    OpenPrw
    ClosePrw
End Enum

Private m_iNumeroCopieDefault As Integer
Private m_OrientamentoDefault As OrientationConsts


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

Public bNotReturnValue As Boolean


'Versione del controllo ActiveBar
Private Const BARMENUVERSION = "3.0"
'Variabile per la gestione degli shortcut del Menu
Private aryShortCut(1) As New ActiveBar3LibraryCtl.ShortCut


'Oggetto utilizzato per ottenere la notifica dei cambiamenti di valore
'all'interno dei singoli campi del documento
Private WithEvents oDocChangeNotify As DocChangeNotify
Attribute oDocChangeNotify.VB_VarHelpID = -1


'Indica se stiamo inserendo un nuovo dettaglio (si è premuto il pulsante Nuovo)
Private bNuovoDettaglio As Boolean
'Indica se si sta prendendo in variazione un dettaglio (si è cliccato sulla listview dei dettagli)
Private bVariazioneDettaglio As Boolean
'Indica se si è in fase di caricamento di un documento esistente
Private bloading As Boolean
Private bLoadingRiga As Boolean

Private A_Riga(1) As Long
'Variabili recordset per le visualizzazioni delle griglie
Private rsGriglia As DmtOleDbLib.adoResultset
Private NuovoRecordComm As Long


Private NumeroRigaSelezionata As Long
Private NumeroRecordPerModifica As Long
Private NumeroRecordLista As Long


Private Mov As DmtMovim.cMovimentazione


Private Const IDDocumento As Long = 11
Private rsTmpPed As ADODB.Recordset
Private rsRettifica As ADODB.Recordset


Public Property Set Application(ByVal NewValue As DMTRunAppLib.Application)
    Set m_App = NewValue
End Property

Public Property Get Application() As DMTRunAppLib.Application
    Set Application = m_App
End Property



'**+
'Nome: ChangeStringsLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge le stringe dal file di risorse per gestire l'opzione multilingue.
'Qui vanno inserite tutte le stringhe aggiunte in frmMain solo se si vuole
'gestire l'opzione multilingua
'**/
Public Sub ChangeStringsLanguage()
    '//////////////////////////////////////////////////////////////////////////////
    'ATTENZIONE
    'Inserire qui il codice per la lettura dal file di risorse di tutte le stringhe
    'per le quali si vuole gestire l'opzione multilingue.
    '//////////////////////////////////////////////////////////////////////////////
End Sub

'**+
'Nome: ChangeToolBarLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge dal file di risorse le stringe delle ToolTipText e dei suggerimenti da visualizzare
'sulla Statusbar per gestire l'opzione multilingue
'**/
Public Sub ChangeToolBarLanguage()

    'New
    BarMenu.Bands("Standard").Tools("New").ToolTipText = GetToolTipText4ToolBar("New")
    BarMenu.Bands("Standard").Tools("New").Description = GetDescription4StatusBar("New")
    
    'Save
    BarMenu.Bands("Standard").Tools("Save").ToolTipText = GetToolTipText4ToolBar("Save")
    BarMenu.Bands("Standard").Tools("Save").Description = GetDescription4StatusBar("Save")

    'Print
    BarMenu.Bands("Standard").Tools("Print").ToolTipText = GetToolTipText4ToolBar("Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")

    'PrePrint
    BarMenu.Bands("Standard").Tools("PrePrint").ToolTipText = GetToolTipText4ToolBar("PrePrint")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")

    'Cut
    BarMenu.Bands("Standard").Tools("Cut").ToolTipText = GetToolTipText4ToolBar("Cut")
    BarMenu.Bands("Standard").Tools("Cut").Description = GetDescription4StatusBar("Cut")

    'Copy
    BarMenu.Bands("Standard").Tools("Copy").ToolTipText = GetToolTipText4ToolBar("Copy")
    BarMenu.Bands("Standard").Tools("Copy").Description = GetDescription4StatusBar("Copy")

    'Paste
    BarMenu.Bands("Standard").Tools("Paste").ToolTipText = GetToolTipText4ToolBar("Paste")
    BarMenu.Bands("Standard").Tools("Paste").Description = GetDescription4StatusBar("Paste")

    'Delete
    BarMenu.Bands("Standard").Tools("Delete").ToolTipText = GetToolTipText4ToolBar("Delete")
    BarMenu.Bands("Standard").Tools("Delete").Description = GetDescription4StatusBar("Delete")

    'Clear
    BarMenu.Bands("Standard").Tools("Clear").ToolTipText = GetToolTipText4ToolBar("Clear")
    BarMenu.Bands("Standard").Tools("Clear").Description = GetDescription4StatusBar("Clear")

    'NewSearch
    BarMenu.Bands("Standard").Tools("NewSearch").ToolTipText = GetToolTipText4ToolBar("NewSearch")
    BarMenu.Bands("Standard").Tools("NewSearch").Description = GetDescription4StatusBar("NewSearch")

    'ExecuteSearch
    BarMenu.Bands("Standard").Tools("ExecuteSearch").ToolTipText = GetToolTipText4ToolBar("ExecuteSearch")
    BarMenu.Bands("Standard").Tools("ExecuteSearch").Description = GetDescription4StatusBar("ExecuteSearch")

    'ChangeView
    BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").ToolTipText = GetToolTipText4ToolBar("ChangeView")

    'SearchPrevious
    BarMenu.Bands("Standard").Tools("SearchPrevious").ToolTipText = GetToolTipText4ToolBar("SearchPrevious")
    BarMenu.Bands("Standard").Tools("SearchPrevious").Description = GetDescription4StatusBar("SearchPrevious")

    'SearchNext
    BarMenu.Bands("Standard").Tools("SearchNext").ToolTipText = GetToolTipText4ToolBar("SearchNext")
    BarMenu.Bands("Standard").Tools("SearchNext").Description = GetDescription4StatusBar("SearchNext")

    'ExportWord
    BarMenu.Bands("Band_Export").Tools("ExportWord").ToolTipText = GetToolTipText4ToolBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")

    'ExportExcel
    BarMenu.Bands("Band_Export").Tools("ExportExcel").ToolTipText = GetToolTipText4ToolBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")

    'ExportHtml
    BarMenu.Bands("Band_Export").Tools("ExportHtml").ToolTipText = GetToolTipText4ToolBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")

    'ExportPDF
    BarMenu.Bands("Band_Export").Tools("ExportPDF").ToolTipText = GetToolTipText4ToolBar("ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")

    BarMenu.RecalcLayout
End Sub


'**+
'Nome: ChangeMenuLanguage
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge dal file di risorse le stringe delle Caption e dei suggerimenti da visualizzare
'sulla Statusbar per gestire l'opzione multilingue
'**/
Public Sub ChangeMenuLanguage()

    '--- Menu PopUp del pulsante "ChangeView" della Toolbar ---
    'ChangeView - Form
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'ChangeView - Tabella
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'ChangeView - Filtro
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    '---                           ---                      ---
    

    'File
    BarMenu.Bands("Band_Menu").Tools("File").Caption = GetCaption4MenuBar("File")
    BarMenu.Bands("Band_Menu").Tools("File").Description = GetDescription4StatusBar("File")

    'File-New
    BarMenu.Bands("Band_File").Tools("Mnu_New").Caption = GetCaption4MenuBar("Mnu_New")
    BarMenu.Bands("Band_File").Tools("Mnu_New").Description = GetDescription4StatusBar("Mnu_New")
    
    'File-Save
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Description = GetDescription4StatusBar("Mnu_Save")
    
    'File-PrePrint
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Caption = GetCaption4MenuBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    
    'File-Print
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    
    'File-Exit
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Caption = GetCaption4MenuBar("Mnu_Exit")
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Description = GetDescription4StatusBar("Mnu_Exit")
    
    'Edit
    BarMenu.Bands("Band_Menu").Tools("Edit").Caption = GetCaption4MenuBar("Edit")
    BarMenu.Bands("Band_Menu").Tools("Edit").Description = GetDescription4StatusBar("Edit")
    
    'Edit-Delete
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = GetCaption4MenuBar("Mnu_Delete")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Description = GetDescription4StatusBar("Mnu_Delete")
    
    'Edit-Clear
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Caption = GetCaption4MenuBar("Mnu_Clear")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Description = GetDescription4StatusBar("Mnu_Clear")
    
    'Edit-Cut
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Caption = GetCaption4MenuBar("Mnu_Cut")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Description = GetDescription4StatusBar("Mnu_Cut")
    
    'Edit-Copy
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Caption = GetCaption4MenuBar("Mnu_Copy")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Description = GetDescription4StatusBar("Mnu_Copy")
    
    'Edit-Paste
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Caption = GetCaption4MenuBar("Mnu_Paste")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Description = GetDescription4StatusBar("Mnu_Paste")
    
    'Edit-NewSearch
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Caption = GetCaption4MenuBar("Mnu_NewSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Description = GetDescription4StatusBar("Mnu_NewSearch")
    
    'Edit-ExecuteSearch
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Caption = GetCaption4MenuBar("Mnu_ExecuteSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Description = GetDescription4StatusBar("Mnu_ExecuteSearch")
    
    'Edit-SearchPrevious
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Caption = GetCaption4MenuBar("Mnu_SearchPrevious")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Description = GetDescription4StatusBar("Mnu_SearchPrevious")
    
    'Edit-SearchNext
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Caption = GetCaption4MenuBar("Mnu_SearchNext")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Description = GetDescription4StatusBar("Mnu_SearchNext")
    
    'View
    BarMenu.Bands("Band_Menu").Tools("View").Caption = GetCaption4MenuBar("View")
    BarMenu.Bands("Band_Menu").Tools("View").Description = GetDescription4StatusBar("View")
    
    'View-FormView
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'View-TableView
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'View-SearchFilter
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    'View-Folders
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Caption = GetCaption4MenuBar("Mnu_Folders")
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Description = GetDescription4StatusBar("Mnu_Folders")
    
    'View-ToolBar
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Caption = GetCaption4MenuBar("Mnu_ToolBar")
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Description = GetDescription4StatusBar("Mnu_ToolBar")
    
    'Tools
    BarMenu.Bands("Band_Menu").Tools("Tools").Caption = GetCaption4MenuBar("Tools")
    BarMenu.Bands("Band_Menu").Tools("Tools").Description = GetDescription4StatusBar("Tools")
    
    'Tools-Export
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Caption = GetCaption4MenuBar("Mnu_Export")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Description = GetDescription4StatusBar("Mnu_Export")
    
    'Tools-Options
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Caption = GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Description = GetDescription4StatusBar("Mnu_Options")
    
    'Tools-Export-ExportWord
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("Mnu_ExportWord")
    
    'Tools-Export-ExportExcel
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("Mnu_ExportExcel")
    
    'Tools-Export-ExportHtml
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("Mnu_ExportHtml")

    'Tools-Export-ExportPDF
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("Mnu_ExportPDF")

    'Help-HelpOnLine
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Caption = GetCaption4MenuBar("Mnu_HelpOnLine")
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Description = GetDescription4StatusBar("Mnu_HelpOnLine")
    
    'Help-Arg
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Caption = GetCaption4MenuBar("Mnu_Arg")
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Description = GetDescription4StatusBar("Mnu_Arg")
    
    'Help-Web
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Description = GetDescription4StatusBar("Mnu_Web")
    
    
    'Help-Info
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'Help-Agg_Web
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Description = GetDescription4StatusBar("Mnu_Agg_Web")
    
    'Help-Info
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'PopUp-RunApplication
    BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").Caption = GetCaption4MenuBar("Mnu_RunApplication")
    
    'PopUp-SearchObject
    BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").Caption = GetCaption4MenuBar("Mnu_SearchObject")
    
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: SetStatusBarVisibility
'
'Parametri:Boolean che valorizzerà la proprietà Visible della StatusBar
'
'Valori di ritorno:
'
'Funzionalità:
'Su richiesta di frmOption, Mostra/Nasconde la Statusbar
'**/
Public Sub SetStatusBarVisibility(ByVal bVisible As Boolean)
    stbStatusbar.Visible = bVisible
End Sub

'**+
'Nome: SetToolBarIcons
'
'Parametri:
'LargeIcons - Il tipo di icona da usare per i bottoni,
'grandi o piccole
'
'Valori di ritorno:
'
'Funzionalità:
'Cambia il tipo di icona della ToolBar standard
'**/
Public Sub SetToolBarIcons(ByVal LargeIcons As Boolean)
    Dim iPicture As Integer

    BarMenu.LargeIcons = LargeIcons
    If LargeIcons Then
        BarMenu.Bands("Standard").Tools("New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Export").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML32), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB32), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB32), &HC0C0C0
        
        'cbc - L'icona del pulsante "ChangeView" dipende dalla modalità attuale
        iPicture = IIf(BrwMain.Visible, IDB_STD_FORM32, IDB_STD_GRID32)
        BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
        
        BarMenu.LargeIcons = False
    Else
        BarMenu.Bands("Standard").Tools("New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Export").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML16), &HC0C0C0
        BarMenu.Bands("Band_Export").Tools("ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB16), &HC0C0C0
        BarMenu.Bands("Standard").Tools("Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB16), &HC0C0C0
    
        'L'icona del pulsante "ChangeView" dipende dalla modalità attuale
        iPicture = IIf(BrwMain.Visible, IDB_STD_FORM16, IDB_STD_GRID16)
        BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
    End If
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: SetVisibilityIDFields
'
'Parametri: Optional IDVisible As Variant (Boolean)
'           Se IDVisible è presente (chiamata da frmOption) viene usato il suo valore
'           per settare la visibilità dei campi ID, altrimenti viene letta l'impostazione
'           del registry
'
'Valori di ritorno:
'
'Funzionalità: Mostra/Nasconde i campi ID della Browse
'**/
Public Sub SetVisibilityIDFields(Optional ByVal IDVisible As Variant)
    Dim Col As DmtGridCtl.dgColumnHeader
    Dim bValue As Boolean

    'Legge le impostazioni dal registry
    bValue = IIf(IsMissing(IDVisible), AppOptions.IDFieldsVisibility, IDVisible)

    For Each Col In BrwMain.ColumnsHeader
        If Left(Col.FieldName, 2) = "ID" Then
            Col.Visible = bValue
        End If
    Next Col
    BrwMain.LoadUserSettings
    'L'aspetto della browse viene ridisegnato.
    BrwMain.Refresh
End Sub


'ATTENZIONE: Nella funzione GetDescription4StatusBar vanno impostati tutti i
'            suggerimenti dei pulsanti della Toolbar e delle voci di menu
'            da visualizzare sulla Statusbar.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.

'**+
'Nome:   GetDescription4StatusBar
'
'Parametri: sToolName è il nome del pulsante o della voce di menu per i quali
'           si vuole ottenere il messaggio sulla Statusbar
'
'Valori di ritorno: La stringa da visualizzare sulla StatusBar
'
'Funzionalità: Restituisce la stringa del suggerimento associato ad un bottone
'              della toolbar o ad una voce di menu
'**/
Private Function GetDescription4StatusBar(ByVal sToolName As String) As String
    Dim sApplicationName As String
    Dim sTipoOggetto As String
    Dim sStr As String
    Dim sTemp As String

    
    sApplicationName = m_App.FunctionName
    sTipoOggetto = m_DocType.Name
    
    Select Case sToolName
    
        Case "File"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_App.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_FILE)
        
        Case "New", "Mnu_New"
            sStr = "Crea un nuovo " & sTipoOggetto
            
        Case "Save", "Mnu_Save"
            sStr = "Memorizza il " & sTipoOggetto & " corrente"
        
        Case "Print", "Mnu_Print"
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sStr = "Stampa i " & sTipoOggetto & " correnti"
            Else
                'Si è in modalità form
                sStr = "Stampa il " & sTipoOggetto & " corrente"
            End If
            
        Case "PrePrint", "Mnu_PrePrint"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                sTemp = sTipoOggetto & " correnti"
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_SETPREVIEW)
            Else
                'Si è in modalità form
                sTemp = sTipoOggetto & " corrente"
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_SETPREVIEW)
            End If
    
        Case "Mnu_Exit"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_EXIT)
            
        Case "Edit"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_App.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_MODIFY)
    
        Case "Cut", "Mnu_Cut"
            sStr = gResource.GetCustomizedMessage(IDS_SB_CUT)
            
        Case "Copy", "Mnu_Copy"
            sStr = gResource.GetCustomizedMessage(IDS_SB_COPY)
            
        Case "Paste", "Mnu_Paste"
            sStr = gResource.GetCustomizedMessage(IDS_SB_PASTE)
            
        Case "Delete", "Mnu_Delete"
            sStr = "Elimina il " & sTipoOggetto & " corrente"
            
        Case "Clear", "Mnu_Clear"
            sStr = gResource.GetCustomizedMessage(IDS_SB_CLEAR)
            
        Case "NewSearch", "Mnu_NewSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHWINDOW)
            
        Case "ExecuteSearch", "Mnu_ExecuteSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHEXECUTE)
            
        Case "Mnu_FormView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_FORM)
            
        Case "Mnu_TableView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_TABLE)
            
        Case "Mnu_SearchFilter"
            sStr = "Espone " & m_DocType.Name & " in modo <filtri>."
            
        Case "ChangeView"
            If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
                'Si è in modalità tabellare
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add m_DocType.Name, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_FORM)
            Else
                'Si è in modalità form
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add m_DocType.Name, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_TABLE)
            End If
            
        Case "View"
            sStr = gResource.GetMessage(IDS_SB_DISPLAY)
            
        Case "SearchPrevious", "Mnu_SearchPrevious"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHPREVIOUS)
            
        Case "SearchNext", "Mnu_SearchNext"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_SEARCHNEXT)
            
        Case "Mnu_Folders"
            sStr = "Riquadro attività"
            
        Case "Mnu_ToolBar"
            sStr = gResource.GetMessage(IDS_SB_TOOLBAR)
            
        Case "Tools"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_TOOLS)
            
        Case "Mnu_Export"
            sStr = gResource.GetMessage(IDS_SB_EXPORT)
            
        Case "Mnu_Options"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add TheApp.FunctionName, 1
            sStr = gResource.GetCustomizedMessage(IDS_SB_OPTION)

            
        Case "ExportWord", "Mnu_ExportWord"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTWORD)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTWORD)
            End If
        
        Case "ExportExcel", "Mnu_ExportExcel"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTEXCEL)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTEXCEL)
            End If
        
        Case "ExportHtml", "Mnu_ExportHtml"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTHTML)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTHTML)
            End If
        
        Case "ExportPDF", "Mnu_ExportPDF"
            gResource.CustomStrings.Clear
            If BrwMain.Visible Then
                'Si è in modalità tabellare
                'Qui va inserito il plurale di TipoOggetto
                sTemp = " i " & sTipoOggetto & " correnti "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTACROBAT)
            Else
                'Si è in modalità form
                sTemp = " il " & sTipoOggetto & " corrente "
                gResource.CustomStrings.Add sTemp, 1
                sStr = gResource.GetCustomizedMessage(IDS_SB_EXPORTACROBAT)
            End If
        
        Case "Mnu_HelpOnLine"
            sStr = gResource.GetMessage(IDS_SB_SUMMARY)
            
        Case "Mnu_Arg"
            sStr = gResource.GetMessage(IDS_SB_ARG)
            
        Case "Mnu_Web"
            sStr = gResource.GetMessage(IDS_SB_WEB)
            
        Case "Mnu_Info"
            sStr = gResource.GetMessage(IDS_SB_INFO)
            
        Case "Mnu_Web", "Web"
            sStr = gResource.GetMessage(IDS_SB_WEB)
        
        Case "Mnu_Agg_Web", "Agg_Web"
            sStr = gResource.GetMessage(IDS_SB_AGG_WEB)
    End Select
    
    GetDescription4StatusBar = sStr
End Function
'//////////////////////////////////////////////////////////////////////////////////
'ATTENZIONE: Nella funzione GetToolTipText4ToolBar vanno impostate tutte le
'            stringhe dei ToolTipText dei pulsanti della Toolbar.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.
'//////////////////////////////////////////////////////////////////////////////////
'**+
'Nome:   GetToolTipText4ToolBar
'
'Parametri: sToolName è il nome del pulsante per il quale
'           si vuole ottenere la stringa per la proprietà ToolTipText
'
'Valori di ritorno: La stringa ToolTipText
'
'Funzionalità: Restituisce la stringa del suggerimento associato ad un bottone
'              della toolbar (ToolTipext)
'**/
Private Function GetToolTipText4ToolBar(ByVal sToolName As String) As String
    Dim sStr As String
    
    gResource.CustomStrings.Clear
    
    Select Case sToolName
    
        Case "New"
            sStr = gResource.GetMessage(TT_NEW)
            
        Case "Save"
            sStr = gResource.GetMessage(TT_SAVE)
        
        Case "Print"
            sStr = gResource.GetMessage(TT_PRINT)
            
        Case "PrePrint"
            sStr = gResource.GetMessage(TT_PREVIEW)
    
        Case "Cut"
            sStr = gResource.GetMessage(TT_CUT)
            
        Case "Copy"
            sStr = gResource.GetMessage(TT_COPY)
            
        Case "Paste"
            sStr = gResource.GetMessage(TT_PASTE)
            
        Case "Delete"
            sStr = gResource.GetMessage(TT_DELETE)
            
        Case "Clear"
            sStr = gResource.GetMessage(TT_CLEAR)
            
        Case "NewSearch"
            sStr = gResource.GetMessage(TT_SEARCH)
            
        Case "ExecuteSearch"
            sStr = gResource.GetMessage(TT_SEARCHEXECUTE)
            
        Case "ChangeView"
            If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
                'Si è in modalità tabellare
                sStr = gResource.GetMessage(TT_FORM)
            Else
                'Si è in modalità form
                sStr = gResource.GetMessage(TT_SEARCHRESULT)
            End If
            
        Case "SearchPrevious"
            sStr = gResource.GetMessage(TT_SEARCHPREVIOUS)
            
        Case "SearchNext"
            sStr = gResource.GetMessage(TT_SEARCHNEXT)
            
        Case "ExportWord"
            sStr = gResource.GetMessage(TT_WORD)
        
        Case "ExportExcel"
            sStr = gResource.GetMessage(TT_EXCEL)
        
        Case "ExportHtml"
            sStr = gResource.GetMessage(TT_HTML)
        
        Case "ViewAssistant" 'toolbar
            sStr = gResource.GetMessage(TT_SHOW_ASSISTANT)
            
        Case "Help" 'toolbar e menu
            sStr = gResource.GetMessage(TT_HELP)

    End Select
    
    GetToolTipText4ToolBar = sStr
End Function

'//////////////////////////////////////////////////////////////////////////////////
'ATTENZIONE: Nella funzione GetCaption4MenuBar vanno impostate tutte le
'            stringhe delle Caption delle voci di menu.
'            Per gestire l'opzione multilingua occorre inserire nel file di risorse
'            tutte le stringhe occorrenti.
'//////////////////////////////////////////////////////////////////////////////////
'**+
'Nome:   GetCaption4MenuBar
'
'Parametri: sToolName è il nome della voce di menu per la quale
'           si vuole ottenere la stringa per la Caption
'
'Valori di ritorno: La stringa da visualizzare nella Caption del menu
'
'Funzionalità: Restituisce la stringa della Caption di una voce di menu
'**/
Private Function GetCaption4MenuBar(ByVal sToolName As String) As String
    Dim sStr As String
    
    gResource.CustomStrings.Clear
    
    Select Case sToolName
    
        Case "File"
            sStr = gResource.GetMessage(MNU_FILE)
        
        Case "Mnu_New"
            sStr = gResource.GetMessage(MNU_NEW)
            aryShortCut(1).Value = "Control+N"
            BarMenu.Bands("Band_File").Tools("Mnu_New").ShortCuts = aryShortCut
            
        Case "Mnu_Save"
            If m_App.Language <> 1 Then
                sStr = gResource.GetMessage(MNU_SAVE)
                aryShortCut(1).Value = "Control+S"
                BarMenu.Bands("Band_File").Tools("Mnu_Save").ShortCuts = aryShortCut
            Else
                sStr = gResource.GetMessage(MNU_SAVE)
                aryShortCut(1).Value = "Shift+F12"
                BarMenu.Bands("Band_File").Tools("Mnu_Save").ShortCuts = aryShortCut
            End If
        
        Case "Mnu_PrePrint"
            sStr = gResource.GetMessage(MNU_PREVIEW)
        
        Case "Mnu_Print"
            If m_App.Language <> 1 Then
                sStr = gResource.GetMessage(MNU_PRINT) & "..."
                aryShortCut(1).Value = "Control+P"
                BarMenu.Bands("Band_File").Tools("Mnu_Print").ShortCuts = aryShortCut
            Else
                sStr = gResource.GetMessage(MNU_PRINT) & "..."
                aryShortCut(1).Value = "Control+Shift+F12"
                BarMenu.Bands("Band_File").Tools("Mnu_Print").ShortCuts = aryShortCut
            End If
    
        Case "Mnu_Exit"
            sStr = gResource.GetMessage(MNU_EXIT)
            
        Case "Edit"
            sStr = gResource.GetMessage(MNU_MODIFY)
    
        Case "Mnu_Delete"
            sStr = gResource.GetMessage(MNU_DELETE)
            aryShortCut(1).Value = "Delete"
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
            
        Case "Mnu_Clear"
            sStr = gResource.GetMessage(MNU_CLEAR)
    
        Case "Mnu_Cut"
            sStr = gResource.GetMessage(MNU_CUT)
            aryShortCut(1).Value = "Control+X"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").ShortCuts = aryShortCut
            
        Case "Mnu_Copy"
            sStr = gResource.GetMessage(MNU_COPY)
            aryShortCut(1).Value = "Control+C"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").ShortCuts = aryShortCut
            
        Case "Mnu_Paste"
            sStr = gResource.GetMessage(MNU_PASTE)
            aryShortCut(1).Value = "Control+V"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").ShortCuts = aryShortCut
            
        Case "Mnu_NewSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_FIND)
            aryShortCut(1).Value = "Control+T"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").ShortCuts = aryShortCut
            
        Case "Mnu_ExecuteSearch"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_EXECUTE_SEARCH)
            aryShortCut(1).Value = "Control+E"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").ShortCuts = aryShortCut
            
        Case "Mnu_SearchPrevious"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_PREVIOUS_SEARCH)
            aryShortCut(1).Value = "Control+P"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").ShortCuts = aryShortCut
            
        Case "Mnu_SearchNext"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_NEXT_SEARCH)
            aryShortCut(1).Value = "Control+S"
            frmMain.BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").ShortCuts = aryShortCut
            
        Case "View"
            sStr = gResource.GetMessage(MNU_DISPLAY)
            
        Case "Mnu_FormView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_FORM)
            aryShortCut(1).Value = "Control+F"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_FormView").ShortCuts = aryShortCut
            
        Case "Mnu_TableView"
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add m_DocType.Name, 1
            sStr = gResource.GetMessage(MNU_TABLE)
            aryShortCut(1).Value = "Control+M"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_TableView").ShortCuts = aryShortCut
            
        Case "Mnu_SearchFilter"
            sStr = "Mo&dalità filtri"
            aryShortCut(1).Value = "Control+Shift+T"
            frmMain.BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").ShortCuts = aryShortCut
            
        Case "Mnu_Folders"
            sStr = "&Riquadro attività"
            
        Case "Mnu_ToolBar"
            sStr = gResource.GetMessage(MNU_TOOLBAR)
            
        Case "Tools"
            sStr = gResource.GetMessage(MNU_TOOL)
            
        Case "Mnu_Export"
            sStr = gResource.GetMessage(MNU_EXPORT)
            
        Case "Mnu_Options"
            sStr = gResource.GetMessage(MNU_OPTION)
            
        Case "Mnu_ExportWord"
                sStr = gResource.GetMessage(MNU_EXPORT_WORD)
        
        Case "Mnu_ExportExcel"
                sStr = gResource.GetMessage(MNU_EXPORT_EXCEL)
        
        Case "Mnu_ExportHtml"
                sStr = gResource.GetMessage(MNU_EXPORT_HTML)
        
        Case "Mnu_ExportPDF"
                sStr = gResource.GetMessage(MNU_EXPORT_ACROBAT)
        
        Case "Help" 'toolbar e menu
            sStr = "&?"

        Case "Mnu_HelpOnLine"
            sStr = gResource.GetMessage(MNU_HELP)
            aryShortCut(1).Value = "F1"
            frmMain.BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").ShortCuts = aryShortCut
            
        Case "Mnu_Arg"
            sStr = gResource.GetMessage(MNU_ARG)
            aryShortCut(1).Value = "Shift+F1"
            frmMain.BarMenu.Bands("Band_Help").Tools("Mnu_Arg").ShortCuts = aryShortCut
            
        Case "Mnu_Web"
            sStr = gResource.GetMessage(MNU_WEB)
            
        Case "Mnu_Agg_Web"
            sStr = gResource.GetMessage(MNU_AGG_WEB)
            
        Case "Mnu_Info"
            sStr = gResource.GetMessage(MNU_INFO)
            
        Case "Mnu_RunApplication"
            sStr = gResource.GetMessage(MNU_EXE_GEST)
            aryShortCut(1).Value = "Control+G"
            frmMain.BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").ShortCuts = aryShortCut
        
        Case "Mnu_SearchObject"
            sStr = gResource.GetMessage(MNU_SEARCH)
            aryShortCut(1).Value = "Control+R"
            frmMain.BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").ShortCuts = aryShortCut
            
    End Select
    
    GetCaption4MenuBar = sStr
End Function


'**+
'Nome:   RefreshDescriptions4StatusBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Reimposta i messaggi da visualizzare sulla StatusBar per quelle
'              voci che dipendono dalla modalità di visualizzazione (Form/Tabella).
'
'**/
Private Sub RefreshDescriptions4StatusBar()
    'ATTENZIONE:
    'Inserire qui tutte le voci di menu ed i pulsanti della toolbar per i quali si
    'vuole cambiare il suggerimento sulla StatusBar in funzione della modalità di
    'visualizzazione. Ad esempio è possibile avere dei messaggi al SINGOLARE per
    'la modalità form e PLURALE per la modalità tabellare.
    'La funzione GetDescription4StatusBar si occupa di determinare la frase esatta.
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("ExportPDF")

End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 25/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: Caption2Display
'
'Parametri:
'  Boolean ReadFromGrid - determina se le stringhe per la costruzione della caption devono essere lette direttamente
'  dai campi del documento o dalla collection AllColumns della BrwMain.
'
'Valori di ritorno: String
'
'Funzionalità:
'                  ///////////////////////////////////////////////////////////////////////////////////////////////////////
'                  In questa funzione va inserito il codice per la determinazione della caption del form principale
'                  per le modalità Modify e Browse in base alle esigenze specifiche.
'                  Di default viene usato esclusivamente il campo del documento individuato dalla
'                  costante CAMPO_PER_CAPTION.
'                  ///////////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Function Caption2Display(Optional ByVal ReadFromGrid As Boolean) As String
    If Not m_Document.EOF And Not m_Document.BOF Then
        If Not ReadFromGrid Then
            Caption2Display = m_App.Caption & ": Documento n° " & lngNumero.Value & " del " & dtData.Text
        Else
            Caption2Display = m_App.Caption & ": Documento n° " & fnNotNull(BrwMain.AllColumns("Doc_numero").Value) & " del " & fnNotNull(BrwMain.AllColumns("Doc_data").Value)
        End If
    Else
        Caption2Display = m_App.FunctionName
    End If
End Function



'**+
'Nome :SetStatus4Modality
'
'Parametri:NewModality rappresenta la modalità di visualizzazione
'          che si vuole ottenere.
'          ModePreview è uno switch per apertura o chiusura anteprima di stampa.
'
'Valori di ritorno:
'
'Funzionalità: Abilita i pulsanti della Toolbar e le voci di menu in funzione
'              di una determinata modalità di visualizzazione.
'              (disabilita tutti i rimanenti pulsanti e voci di menu)
'              Imposta la Caption del form in funzione della modalità di visualizzazione
'**/
Private Sub SetStatus4Modality(ByVal NewModality As neVisualModality, _
                                Optional ByVal ModePreview As nePreviewModality)
    Dim KeyON As Currency
    Dim KeyOFF As Currency
    Dim iPicture As Integer
   
    
    'Indica lo stato di visibilità della ToolBar standard
    'prima della visualizzazione della ToolBar della anteprima
    'di stampa
    Static bToolBarStandardVisible As Boolean
    
    'Indica lo stato di attivazione dei bottoni della ToolBar
    'standard prima della visualizzazione della ToolBar della
    'anteprima di stampa
    Static curToolBarStandardStatus As Currency
    
    
    'Elimina l'acceleratore CUT
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = gResource.GetMessage(MNU_DELETE)
    'Rimuove lo shortcut "Delete"
    aryShortCut(1).Clear
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
    BarMenu.Bands("Band_Tools").Tools("Mnu_stampa_doc_sel").Enabled = False
    'Imposta i pulsanti e le voci di menu
    Select Case NewModality
    
        Case Insert
            KeyOFF = BTN_SAVE + BTN_PRINT + BTN_PREVIEW + BTN_DELETE + BTN_SEARCH
            KeyOFF = KeyOFF + BTN_SEARCHTABLE + BTN_SEARCHFORM + BTN_VIEWMODE
            KeyOFF = KeyOFF + BTN_FILTER
            KeyOFF = KeyOFF + BTN_PREVIOUS + BTN_NEXT
            KeyOFF = KeyOFF + BTN_WORD + BTN_EXCEL + BTN_HTML + BTN_PDF
            KeyON = BTN_ALL - KeyOFF
            Me.Caption = m_App.Caption
            oFiltersActivity.AbortNewFilter
            
            If BrwMain.GuiMode = dgFilterDefinition Then
                bEnableGuiEvent = False
                BrwMain.GuiMode = dgNormal
                bEnableGuiEvent = True
            End If
            
            m_Search = False
            
        Case Modify
            KeyOFF = BTN_SAVE + BTN_CLEAR + BTN_SEARCH + BTN_SEARCHFORM
            KeyON = BTN_ALL - KeyOFF
            'in modalità variazione si è necessariamente in modalità form
            'pertanto il pulsante ChangeView della toolbar deve visualizzare
            'l'icona della griglia
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_GRID32, IDB_STD_GRID16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
                'Monta la caption del form principale
                Me.Caption = Caption2Display(False)
            End If
            oFiltersActivity.AbortNewFilter
                        
            m_Search = False
            
        Case Find
        
            'Solo se esiste almeno un elemento nel data manager.
            If Not (m_Document.EOF = True And m_Document.BOF = True) Then
                KeyON = BTN_VIEWMODE + BTN_SEARCHTABLE + BTN_SEARCHFORM
            End If
            KeyON = KeyON + BTN_NEW + BTN_CUT + BTN_COPY + BTN_PASTE
            KeyON = KeyON + BTN_CLEAR + BTN_SEARCH
            KeyOFF = BTN_ALL - KeyON
            'In modalità Find verrà proposto il pulsante per andare in modalità tabella
            'pertanto il pulsante ChangeView della toolbar deve visualizzare
            'l'icona della griglia
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_GRID32, IDB_STD_GRID16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
            Me.Caption = gResource.GetMessage(TT_SEARCH) & " - " & m_App.Caption
            
            oFiltersActivity.AbortNewFilter
                
            'Cancella eventuali blocchi su qualsiasi azione.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            
        Case Browse
            KeyOFF = BTN_SAVE + BTN_CLEAR + BTN_SEARCH + BTN_PREVIOUS + BTN_NEXT
            KeyOFF = KeyOFF + BTN_SEARCHTABLE + BTN_CUT + BTN_COPY + BTN_PASTE
            KeyON = BTN_ALL - KeyOFF
            'Seleziona l'icona grande o piccola in base alle impostazioni correnti
            iPicture = IIf(GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "LargeIcon", False), IDB_STD_FORM32, IDB_STD_FORM16)
            BarMenu.Bands("Standard").Tools("ChangeView").SetPicture 0, gResource.GetBitmap(iPicture), &HC0C0C0
            BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
            If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
                'Monta la caption del form principale
                Me.Caption = Caption2Display(False)
            End If
            'Inserisce l'acceleratore CUT
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = gResource.GetMessage(MNU_DELETE)
            'Inserisce lo shortcut "Delete"
            aryShortCut(1).Value = "Delete"
            BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").ShortCuts = aryShortCut
            If BrwMain.GuiMode = dgNormal Then
                BarMenu.Bands("Band_Tools").Tools("Mnu_stampa_doc_sel").Enabled = True
            End If
            
            'Questo controllo si è reso necessario per evitare un loop infinito
            'con la gestione dell'evento BrwMain_OnChangeGuiMode() quando dal
            'Menu della browse si va in modalità tabellare.
            If BrwMain.GuiMode <> dgNormal Then
                BrwMain.GuiMode = dgNormal
            End If
            
            'Se il filtro attivo è un filtro temporaneo viene abilitato il pulsante
            'Salva Filtro del DocTypeExplorer per poterlo rendere permanente.
            If m_ActiveFilter.ID = -1 Then
                oFiltersActivity.NewFilterBegin   'Abilita il pulsante Salva Filtro
            Else
                oFiltersActivity.AbortNewFilter   'Disabilita il pulsante Salva Filtro
            End If
            ActivityBox.Redraw = True
            
            m_Search = False
            
            'Cancella eventuali blocchi su qualsiasi azione.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions

        Case Preview
            If ModePreview = OpenPrw Then
                bToolBarStandardVisible = BarMenu.Bands("Standard").Visible
                curToolBarStandardStatus = GetStatusToolBar(True)
                KeyON = BTN_PRINT + BTN_EXCEL + BTN_WORD + BTN_HTML + BTN_PDF
                KeyOFF = BTN_ALL - KeyON
                BarMenu.Bands("Band_View").Tools("Mnu_Folders").Enabled = False
                BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = False
                BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Enabled = False
                BarMenu.Bands("Standard").Visible = False
                BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = True
                BarMenu.RecalcLayout
            Else
                BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = False
                BarMenu.Bands("Standard").Visible = bToolBarStandardVisible
                ActivateBarButtons curToolBarStandardStatus, True
                ActivateBarButtons BTN_ALL - curToolBarStandardStatus, False
                BarMenu.Bands("Band_View").Tools("Mnu_Folders").Enabled = True
                BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = True
                BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Enabled = True
                BarMenu.RecalcLayout
            End If
            
    End Select
    
    'Attiva/disattiva i pulsanti e le voci di menu
    ActivateBarButtons KeyON, True
    ActivateBarButtons KeyOFF, False
End Sub

'**+
'Autore                 : Diamante S.p.a
'
'Nome                   : PermissionToSave
'
'Parametri:
'
'Valori di ritorno: True se il documento può essere salvato, False altrimenti.
'
'Funzionalità: Controlli da effettuare PRIMA di salvare il documento corrente
'
'**/
Private Function PermissionToSave() As Boolean
On Error GoTo ERR_PermissionToSave
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Testo As String

    PermissionToSave = True
    
    If dtData.Value = 0 Then
        sbMsgInfo "Impossibile salvare il documento corrente senza aver specificato la data documento", m_App.FunctionName
        PermissionToSave = False
        'dtData.SetFocus
        
        Exit Function
    End If
    
    If cboPagamento.CurrentID = 0 Then
        PermissionToSave = False
        sbMsgInfo "Impossibile salvare il documento corrente senza aver specificato la modalità di pagamento", m_App.FunctionName
        'cboPagamento.SetFocus
        
        Exit Function
    End If
    If cdAnagrafica.KeyFieldID = 0 Then
        PermissionToSave = False
        sbMsgInfo "Impossibile salvare il documento corrente senza aver specificato il cliente", m_App.FunctionName
        'cdAnagrafica.SetFocus
        
        Exit Function
    End If
    If LINK_BLOCCO_CLIENTE = 1 Then
        PermissionToSave = False
        sbMsgInfo "Impossibile salvare il documento corrente poichè il cliente risulta bloccato", m_App.FunctionName
        'cdAnagrafica.SetFocus
       
        Exit Function
    End If
    If (ATTIVA_FIDO_ALY = 1) Then
        If (CONTROLLO_FIDO_ALYANTE(oDoc.Field("Link_Nom_anagrafica", , sTabellaTestata)) = True) Then
            ALY_CONFERMA_SALVA_DOC = False
            frmFidoAly.Show vbModal
            If (ALY_CONFERMA_SALVA_DOC = False) Then
                Screen.MousePointer = 0
                PermissionToSave = False
                Exit Function
            End If
        End If
    Else
        If GET_CONTROLLO_FIDO_CLIENTE = False Then
            AVVIA_FIDO_DOPO_CONTROLLO = False
            frmFido.Show vbModal
            If AVVIA_FIDO_DOPO_CONTROLLO = False Then
                'cdAnagrafica.SetFocus
                PermissionToSave = False
                Exit Function
            End If
        End If
    End If
    If oDoc.IDOggetto > 0 Then
        '''''''''''''''''''''''CONTROLLO SE ORDINE CONFERMATO'''''''''''''''''''''
        sSQL = "SELECT IDOggetto FROM RV_POTMPEvasioneOrdini "
        sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
        
        Set rs = Cn.OpenResultset(sSQL)
        If Not rs.EOF Then
            PermissionToSave = False
            MsgBox "Impossibile aggiornare l'ordine poichè risulta confermato per la vendita", vbCritical, TheApp.FunctionName
            
            Exit Function
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
    
    If Me.CDAgenteTesta.KeyFieldID > 0 Then
        If GET_LINK_REGOLA_PROVV_AGE(Me.CDAgenteTesta.KeyFieldID, TheApp.IDFirm) = 0 Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "L'agente selezionato non ha impostata una regola predefinita di calcolo della provvigione." & vbCrLf
            Testo = Testo & "Vuoi continuare a salvare il documento?"
            If MsgBox(Testo, vbYesNo + vbQuestion, TheApp.FunctionName) = vbNo Then
                PermissionToSave = False
                Exit Function
            End If
        End If
    End If
    
    If oDoc.IDOggetto > 0 Then
        If GET_CONTROLLO_DATI_TOUR(oDoc.IDOggetto, Me.cboVettore.CurrentID, Trim(Me.txtTargaAutomezzo.Text)) = True Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Il vettore o/e la targa automezzo non sono congruenti ai dati del tour collegato." & vbCrLf
            Testo = Testo & "Vuoi continuare con il salvataggio del documento?"
            If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then
                PermissionToSave = False
                Exit Function
            End If
        End If
    End If
    
    If Me.txtIDContratto.Value > 0 Then
        If Me.chkConfDaContratto.Value = vbUnchecked Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Per continuare con il salvataggio del documento bisogna confermare la presa in visione delle note sulla testa e le note di ogni singola riga del documento che provengono dal contratto selezionato"
            sbMsgInfo Testo, m_App.FunctionName
            PermissionToSave = False
            Exit Function
        End If
    End If
    
Exit Function

ERR_PermissionToSave:
    MsgBox Err.Description, vbCritical, "PermissionToSave"
End Function


'**+
'Nome: SearchNext
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Posizionamento al record successivo
'**/
Private Sub SearchNext()
    
    m_Document.MoveNext
    
    If m_Document.EOF Then
        'Si era già sull'ultimo record (prima di MoveNext).
        
        'Si annulla l'operazione
        m_Document.MovePrevious
        sbMsgInfo gResource.GetMessage(MESS_NO_NEXT_ELEMENTS), m_App.FunctionName
        Exit Sub
    Else
        'Controlla la presenza di eventuali conflitti nel caso di multiutenza.
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions) Then
            m_Document.MovePrevious
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
        End If
    End If
    
End Sub

'**+
'Nome: SearchPrevious
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Posizionamento al record precedente
'**/
Private Sub SearchPrevious()
    
    m_Document.MovePrevious
    
    If m_Document.BOF Then
        'Si era già sul primo record (prima di MovePrevious).
        
        'Si annulla l'operazione
        m_Document.MoveNext
        sbMsgInfo gResource.GetMessage(MESS_NO_PREVIOUS_ELEMENTS), m_App.FunctionName
        Exit Sub
    Else
        'Controlla la presenza di eventuali conflitti nel caso di multiutenza.
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions) Then
            m_Document.MoveNext
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
        End If
    End If
End Sub

'**+
'Nome: BrowseReposition
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni da compiere al riposizionamento del record corrente
'**/
Private Sub BrowseReposition()

    'Dopo un Save del documento avviene un Refresh della Browse ma in tal caso
    'è inutile effettuare il refresh del form.
    If Not m_bAvoidReposition Then
    
        'Refresh dei campi del form
        RefreshFormFields
        
        'Refresh della caption del Form
        If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
            'Monta la caption del form principale
            Me.Caption = Caption2Display(False)
        End If
        
    End If
 
    'Refresh delle variabili di stato
    m_Changed = False
    m_Saved = False
    m_Search = False
    
    'Annullamento di un eventuale inizio di inserimento di un nuovo record
    m_Document.AbortNew
    
End Sub



'**+
'Nome: NewRecord
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni su richiesta nuovo record
'**/
Private Sub NewRecord()
Dim IDListinoDefault As Long


'--------------------------------------------------------------------------------------------
'NOTA:
'Il gruppo di istruzioni sottostanti e la riga  'Imposta il blocco su inserimento'
'sono state commentate per far si che la manutenzione NON imposti alcun blocco per
'l'azione Inserimento.
'Pertanto 2 o più utenti potranno effettuare contemporaneamente la suddetta azione.
'Se si intende impedire questa possibilità sarà sufficiente ripristinare le righe commentate.
'--------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
'    'Controllo se ho il permesso di salvare ( nel caso di conflitti di multiutenza )
'    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, SemAllObjects, SemInsertAction) Then
'        'C'è un altro utente in modalità inserimento che blocca la medesima azione per
'        'tutti gli altri utenti. Pertanto annullo l'operazione di inserimento ed esco.
'        Exit Sub
'    End If




    'Ho il permesso per l'azione inserimento.
    '
    'Cancella il blocco precedente
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    
    '--------------------------------
    'Imposta il blocco su inserimento
'    m_Semaphore.SetObjectAction m_DocType.ID, SemAllObjects, SemInsertAction
    '
    'A questo punto nessun altro utente potrà effettuare una operazione di inserimento
    'finchè non verrà cancellato il blocco su inserimento.

    'Pulisce i controlli presenti nella form
    If m_bOnFirstTime = False Then
        OnClear
    End If
    
    'Pulisce i valori presenti all'interno della struttura di tabelle della DmtDocs
    oDoc.ClearValues
    
   
    'Imposta l'IDAzienda per cui si genererà il documento
    oDoc.IDAzienda = TheApp.IDFirm
    'Imposta l'IDFiliale per cui si genererà il documento
    oDoc.IDFiliale = TheApp.Branch
    'Imposta il tipo di anagrafica da utilizzare per il documento corrente
    'Nel caso di documenti che hanno come soggetto un fornitore bisogna indicare il valore 3
    oDoc.IDTipoAnagrafica = 2 'Cliente
    'Imposta l'identificativo dell'utente corrente
    oDoc.IDUtente = TheApp.IDUser
    'Imposta il sezionale da utilizzare per la numerazione del documento
    
    Me.dtData.Enabled = True
    Me.lngNumero.Enabled = True
    Me.cboSezionale.Enabled = True
    
    Me.dtData.Value = Date
    'Imposta la data con la data di sistema
    oDoc.Field "Doc_data", Date, sTabellaTestata
    
    fnSetSezionale
    fnSetCausaleDocumento
    
    Me.cboMagazzino.WriteOn fnGetParametriMagazzino("IDMagazzino_Vendita")
    oDoc.Field "Link_Val_valuta", oDoc.DBDefaults.Link_Val_valuta_nazionale, sTabellaTestata
    oDoc.Field "Link_Val_cambio", Null, sTabellaTestata
    Me.cboListinoAzienda.WriteOn oDoc.DBDefaults.Link_Doc_listino_base
    Me.cboPagamento.WriteOn oDoc.DBDefaults.IDPagamentoDocDefault

    NumeroRiga = 0
    
    'Inizializza il documento
    oDoc.InitializeDocument
    
    
    'Imposta la riga attiva per la tabella di testata (sempre la prima ed unica riga)
    'oDoc.Tables(sTabellaTestata).SetActiveRetail 1
    
    'Imposta il magazzino di default
    oDoc.ReadDataFromStore fnGetParametriMagazzino("IDMagazzino_Vendita"), MainStore
    oDoc.Field "Link_Doc_contratto_bancario_az", oDoc.DBDefaults.Link_Doc_contratto_bancario_az, sTabellaTestata
    oDoc.Field "Link_Val_valuta", oDoc.DBDefaults.Link_Val_valuta_nazionale, sTabellaTestata
    
    'Imposta la valuta con la valuta nazionale
    'oDoc.Field "Link_Val_valuta", oDoc.DBDefaults.Link_Val_valuta_nazionale, sTabellaTestata
    'oDoc.Field "Link_Val_cambio", Null, sTabellaTestata
    
    
    'IDListinoDefault = GET_LISTINO_DEFAULT(Me.cdAnagrafica.KeyFieldID)
    'If IDListinoDefault = 0 Then
    '    IDListinoDefault = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
    'End If
    
    'Me.cboListino.WriteOn IDListinoDefault
    'oDoc.Field "Link_Doc_Listino", IDListinoDefault, sTabellaTestata
    
    NuovoDocumento = 0
    'Aggiorna il contenuto delle listview
    fnEliminaDatiTemporanei
    
    
    Me.txtNOrdinePadre.Value = Me.lngNumero.Value
    Me.txtDataDocPadre.Value = Me.dtData.Value
    Me.txtNListaPrelievo.Value = 1
    
    sbPopalaListaArticoli True
    sbPopalaListaIva
    sbPopalaListaScadenze
    sbPopalaListaCommissioni
    'Si predispone per l'inserimento di un nuovo articolo
    'cmdNuovo_Click
    
    'Refresh delle variabili di stato
    m_Search = False
    m_Changed = False
    m_Saved = False
    
    'Refresh della toolbar in modalità inserimento
    SetStatus4Modality 0 'Insert
    
    'Ripristina la vista del Form
    BrwMain.Visible = False
    
    
    'Il primo campo del Form riceve l'input focus
    'CONTROLLA_BLOCCHI_INSERIMENTI
    'SSTab1.TabEnabled(2) = False
    SSTab1.Tab = 0

    
    If VISUALIZZA_ANDAMENTO_ORDINE = 1 Then
        GET_ANDAMENTO_ORDINE fnNotNullN(oDoc.Field("RV_POIDOrdinePadre", , sTabellaTestata))
    End If
    
    
    
    On Error Resume Next
    dtData.SetFocus
    
    
    ABILITA_CONTROLLI
    
End Sub

'**+
'Nome                   : ClearControl
'
'Parametri              : ctrControl As Control - controllo da pulire
'
'Valori di ritorno      :
'
'Funzionalità           : Pulisce un controllo sulla base del tipo del controllo stesso
'
'**/
Private Sub ClearControl(ByVal ctrControl As Control)
    Dim sType As String

    sType = TypeName(ctrControl)
    
    If sType = "fpDateTime" Or sType = "TextBox" Or sType = "fpText" Or sType = "fpLongInteger" Or sType = "fpCurrency" Or sType = "fpDoubleSingle" Or sType = "dmtDate" Or sType = "dmtTime" Then
        ctrControl.Text = ""
    ElseIf sType = "CheckBox" Then
        ctrControl.Value = 0
    ElseIf sType = "fpBoolean" Then
        ctrControl.Value = 0
    ElseIf sType = "ComboBox" Then
        ctrControl.ListIndex = -1
    ElseIf sType = "DMTCombo" Then
        ctrControl.ListIndex = -1
    ElseIf sType = "ListBox" Then
        ctrControl.ListIndex = -1
    ElseIf sType = "ListView" Then
        ctrControl.ListItems.Clear
    ElseIf sType = "TreeView" Then
        ctrControl.Nodes.Clear
    ElseIf sType = "Town" Then
        ctrControl.Reset
    ElseIf sType = "dmtCurrency" Or sType = "dmtNumber" Then
        ctrControl.Value = 0
        'ctrControl.Text = ""
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDNode = 0
    ElseIf sType = "DmtFirmGerarchy" Then
        ctrControl.LoadActivity 0
    ElseIf sType = "DMTProgControl" Then
        'Queste istruzioni forzano il refresh
        'e il reset del componente
        ctrControl.IDArticolo = 0
        ctrControl.Show
    ElseIf sType = "DmtCodDesc" Then
        ctrControl.Load 0
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDAnagrafica = 0
        ctrControl.Description = ""
        ctrControl.Code = ""
        ctrControl.SecondDescription = ""

    End If
End Sub


'**+
'Nome: ClearFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Pulisce il contenuto dei campi di input del Form
'**/
Private Sub ClearFormFields()
    ClearControl dtData
    ClearControl lngNumero
    ClearControl chkLordoIVA
    ClearControl cboListino
    ClearControl cboPagamento
    ClearControl cdAnagrafica
    ClearControl txtPartitaIva
    ClearControl txtIndirizzo
    ClearControl txtCAP
    ClearControl txtComune
    ClearControl txtProvincia
    ClearControl lvwIVA
    ClearControl lvwScadenze
    'ClearControl curSpeseImballo
    ClearControl curSpeseIncasso
    ClearControl curSpeseTrasporto
    ClearControl lngScontoDocPer
    ClearControl curScontoDocImp
    ClearControl curTotImponibile
    ClearControl curTotImposta
    ClearControl curTotDocumento
    ClearControl curTotArrotondamenti
    ClearControl curNettoAPagare
    ClearControl txtAnnotazioni
    ClearControl txtDataOrdineCliente
    ClearControl txtNumeroOrdineCliente
    ClearControl txtDataPartenza
    ClearControl txtCausaleDocumento
    ClearControl txtDataTrasporto
    ClearControl txtOraTrasporto
    ClearControl cboListinoAzienda
    ClearControl cboLuogoPresaMerce
    ClearControl cboVettoreSuccessivo
    ClearControl txtAnnotazioniInterna
    ClearControl txtDescrizioneRigaDoc
    ClearControl chkRaggruppBolle
    ClearControl chkRaggruppaScadenze
    ClearControl chkChiuso
    ClearControl cboAspettoEsteriore
    ClearControl cboAltroSito
    ClearControl cboTipoOrdine
    ClearControl txtDataTrasporto
    ClearControl txtOraTrasporto
    ClearControl txtDataArrivoLuogo
    ClearControl txtOraArrivoLuogo
    ClearControl CDAgenteTesta
    ClearControl txtTargaAutomezzo
    ClearControl txtIstruzioniMittente
    ClearControl txtNOrdinePadre
    ClearControl txtDataDocPadre
    ClearControl txtNListaPrelievo
    ClearControl chkOrdineCompletato
    ClearControl txtNPedaneTesta
    ClearControl txtColliTotali
    ClearControl txtPesoTotale
    ClearControl ACSAnaDest
    ClearControl txtIDContratto
    ClearControl cboRaggrFatturato
    ClearControl chkConfDaContratto
    ClearControl chkStampaFattProForma
End Sub

'**+
'Nome: ExecuteMenuCommand
'
'Parametri:
'sToolName - Nome del comando selezionato
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione dei comandi generati dal controllo ActiveBar
'**/
Private Sub ExecuteMenuCommand(ByVal sToolName As String)
    Dim iAnswer As Integer

    'cbcxn
    'Notifica alla (eventuale) applicazione che gestisce il processo On_Extend la
    'pressione di un Tool. Se l'applicazione chiamata restituisce True viene annullata l'operazione.
    'If m_ExtendApplication.BeforeCommandClick(sToolName) Then Exit Sub

    Select Case sToolName
        Case "Cut", "Mnu_Cut"
            SendKeys ("+{DEL}")
            
        Case "Copy", "Mnu_Copy"
            SendKeys ("^{INSERT}")
            
        Case "Paste", "Mnu_Paste"
            SendKeys ("+{INSERT}")
            
        Case "Mnu_Folders"
            OnFolders
            
        Case "ClosePreview"
            ClosePreview
            
        Case "Save", "Mnu_Save"
            OnSave
            
        Case "Mnu_Exit"
            Unload frmMain
            
        Case "Delete", "Mnu_Delete"
            OnDelete
            
        Case "Clear", "Mnu_Clear"
            OnClear
            
        Case "ExecuteSearch", "Mnu_ExecuteSearch"
            OnExecuteSearch
        
        Case "SearchNext", "Mnu_SearchNext"
            
            OnMoveCurrentRecord SRCNEXT, sToolName
        
        Case "SearchPrevious", "Mnu_SearchPrevious"
            
            OnMoveCurrentRecord SRCPREVIOUS, sToolName
        
        Case "ChangeView", "Mnu_FormView", "Mnu_TableView"
            OnChangeView sToolName
            
        Case "Mnu_ToolBar"
            OnToolBarOptions
            
        Case "Mnu_Options"
            OnOptions
            
        Case "Mnu_Info"
            OnInfo
            
        Case "PrePrint", "Mnu_PrePrint", "Print", "Mnu_Print", "ExportPDF", "Mnu_ExportPDF", "MailPDF", "ExportWord", "Mnu_ExportWord", "MailWord", "ExportExcel", "Mnu_ExportExcel", "MailExcel", "ExportHtml", "Mnu_ExportHtml", "MailMHTL"
            OnPrint sToolName
            
        Case "NewSearch", "Mnu_NewSearch", "Mnu_SearchFilter"
            OnNewSearch
            
        Case "New", "Mnu_New"
            
            OnNew sToolName
           
        Case "Mnu_RunApplication", "Mnu_SearchObject"
            OnRunApplication sToolName
        Case "Mnu_Summary"
            OnSummary
        Case "Mnu_FastHelp", "Help"
            OnFastHelp
        Case "Mnu_HelpOnLine"
            OnHelpOnLine
        Case "Mnu_Arg"
             OnArg
        Case "Mnu_Web1"
             sbOpenURL hwnd, URL_DIAMANTE
        Case "Mnu_stampa_doc_sel"
            GET_STAMPA_DOCUMENTI_SEL
        Case "cmdLavorazioni"
            cmdLavorazioni_Click
        Case "cmdAnalizzaOrdine"
            cmdAnalizzaOrdine_Click
        Case "cmdTour"
            cmdTour_Click
        Case "cmdConferimento"
            cmdConferimento_Click
        Case "cmdSalvaComeNuovo"
            cmdSalvaComeNuovo_Click
        Case "cmdListaPrelievo"
            cmdListaPrelievo_Click
        Case "cmdContratto"
            GET_CONTRATTO
        Case "cmdAvviaListaQualita"
             cmdGestioneQualitaLista
        Case "cmdAvviaQualita"
             cmdGestioneQualita
    End Select
    

    'cbcxn
    'Notifica alla (eventuale) applicazione che gestisce il processo On_Extend la
    'pressione di un Tool DOPO avere eseguito l'operazione ad esso associata
    'm_ExtendApplication.AfterCommandClick sToolName
    
End Sub

'**+
'Nome: RefreshFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Riempie i valori dei campi del Form con i valori
'del documento.
'**/
Private Sub RefreshFormFields()
    
    'In questi casi non si deve far nulla
    If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
        'Viene impostata la variabile per indicare che stiamo procedendo alla lettura di un documento
        bloading = True
        
        oDoc.ClearValues
        'Leggiamo il documento
        oDoc.ReadWithTO m_Document("IDOggetto").Value, m_Document("IDTipoOggetto").Value
        
        'Aggiorniamo il contenuto delle listview
        'fnEliminaDatiTemporanei
        'Viene impostata la variabile per indicare la lettura del documento è terminata
        bloading = False
    End If
    
End Sub

'**+
'Nome: ClearFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Azzera i valori dei campi del documento
'**/
Private Sub ClearFields()
    Dim Field As DmtDocManLib.Field
    
    For Each Field In m_Document.Fields
        Field.Value = Empty
    Next
End Sub

'**+
'Nome: Change
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni su variazione di un campo del Form
'**/
Private Sub Change()
    'Se si è in modalità tabellare non deve essere eseguita perchè
    'altrimenti al Click della Browse si attiverebbe il pulsante Salva
    If Not m_Search And Not BrwMain.Visible Then
        ActivateBarButtons BTN_SAVE, True
    
        m_Changed = True
        m_Saved = False
        m_Search = False
    End If
End Sub

'**+
'Nome: ClosePreview
'**+
'Nome: ClosePreview
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Chiude la finestra della anteprima di stampa
'**/
Private Sub ClosePreview()
    Dim myDate
    
    On Error GoTo errHandler
        
    If m_Report.ClosePreview Then
        m_PreviewWindowHandle = 0
        PicForm.Visible = True
        BrwMain.Visible = m_TabMode
        ActivityBox.Visible = BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked
        FormRecalcLayout
        Set m_Report = Nothing
        SetStatus4Modality Preview, ClosePrw
    End If
    Exit Sub
errHandler:
    'Se si verifica un errore "SQL server in use"
    'la subroutine entra in un ciclo di attesa per
    '3 secondi prima di tentare nuovamente la chiusura
    myDate = Now
    If Err.Description = "SQL server in use" Then
        While Not (Now = DateAdd("s", 3, myDate))
        Wend
        Resume
    End If
    Err.Raise Err.Number, , Err.Description
End Sub




'**+
'Nome: ShortCut
'
'Parametri:
'KeyCode - Codice del tasto
'Shift - Stato del tasto Shift
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione degli accelleratori da tastiera
'**/
'**+
'Nome: ShortCut
'
'Parametri:
'KeyCode - Codice del tasto
'Shift - Stato del tasto Shift
'
'Valori di ritorno:
'
'Funzionalità:
'Gestione degli accelleratori da tastiera
'**/
Private Function ShortCut(KeyCode As Integer, Shift As Integer) As Boolean
    Dim bCtrlDown As Boolean
    Dim bShiftDown As Boolean
    Dim bAltDown As Boolean
    
    bShiftDown = (Shift And vbShiftMask) > 0
    bCtrlDown = (Shift And vbCtrlMask) > 0
    bAltDown = (Shift And vbAltMask) > 0
    
    Select Case KeyCode
         Case vbKeyF12
            If bShiftDown Then
                If bCtrlDown Then
                    If BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled Then
                        ExecuteMenuCommand ("Mnu_Print")
                        ShortCut = True
                    End If
                Else
                    If BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled Then
                    
                        'Forza il lostfocus ed attende l'esecuzione di eventuali eventi associati
                        AutoLostFocus
                        
                        ExecuteMenuCommand ("Mnu_Save")
                        ShortCut = True
                    End If
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyF1
            SendMessage hwnd, WM_SETREDRAW, 0, 0
            'SendKeys ("{ESC}")
            DoEvents
            SendMessage hwnd, WM_SETREDRAW, 1, 0
            If bShiftDown Then
                'case shift F1
                ExecuteMenuCommand ("Mnu_Arg")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            Else
                ExecuteMenuCommand ("Mnu_HelpOnLine")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyN
            If bCtrlDown Then
                If BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled Then
                    ExecuteMenuCommand ("Mnu_New")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyX
'            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Cut")
'                    ShortCut = True
'                End If
''                KeyCode = 0
''                Shift = 0
'            End If
            
        Case vbKeyC
            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Copy")
'                    ShortCut = True
'                End If
                KeyCode = 0
                Shift = 0
            End If
            If bAltDown And BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible Then
                ClosePreview
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyV
            If bCtrlDown Then
'                If BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled Then
'                    ExecuteMenuCommand ("Mnu_Paste")
'                    ShortCut = True
'                End If
                'KeyCode = 0
                'Shift = 0
            End If
            
        Case vbKeyT
            If bCtrlDown And bShiftDown = False Then   'CTRL + T
                If BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled Then
                    ExecuteMenuCommand ("Mnu_NewSearch")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            If bCtrlDown And bShiftDown = True Then     'CTRL + MAIUSC + T
                If BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchFilter")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyE
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled Then
                    ExecuteMenuCommand "Mnu_ExecuteSearch"
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyP
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchPrevious")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyS
            If bCtrlDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled Then
                    ExecuteMenuCommand ("Mnu_SearchNext")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyM
            If bCtrlDown Then
                If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled Then
                    ExecuteMenuCommand ("Mnu_TableView")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyF
            If bCtrlDown Then
                If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled Then
                    ExecuteMenuCommand ("Mnu_FormView")
                    ShortCut = True
                End If
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyDelete
            'Il tasto Canc ha effetto solo se il controllo attivo è la browse principale.
            If ActiveControl.Name = "BrwMain" And Not bShiftDown Then
                If BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled Then
                    If BrwMain.Visible Then
                        ExecuteMenuCommand ("Mnu_Delete")
                        ShortCut = True
                        KeyCode = 0
                        Shift = 0
                    End If
                End If
            End If
            
        Case vbKeyR
            If bCtrlDown Then
                ExecuteMenuCommand ("Mnu_SearchObject")
                ShortCut = True
                'La condizione sottostante è necessaria per attivare l'acceleratore CTRL+R dalla modalità
                'filtri della DmrGrid
                If Not BrwMain.Visible Or (BrwMain.Visible And BrwMain.GuiMode = dgNormal) Then
                    KeyCode = 0
                    Shift = 0
                End If
            End If
            
        Case vbKeyG
            If bCtrlDown Then
                ExecuteMenuCommand ("Mnu_RunApplication")
                ShortCut = True
                KeyCode = 0
                Shift = 0
            End If
    
        Case vbKeyEscape
             If Not ActiveControl Is Nothing Then
                    If TypeName(ActiveControl) = "DmtGrid" Then
                        If BrwMain.GuiMode = dgFilterDefinition Then
                            If Not (m_Document.EOF = True And m_Document.BOF = True) Then
                                BrwMain.GuiMode = dgNormal
                                ExecuteMenuCommand "Mnu_TableView"
                                ShortCut = True
                            Else
                                'Ripulisce il contenuto delle condizioni.
                                BrwMain.Conditions.ClearValues
                                'Imposta la modalità FilterDefinition
                                BrwMain.GuiMode = dgFilterDefinition
                                ShortCut = True
                            End If
                        End If
                    End If
                KeyCode = 0
                Shift = 0
            End If
    
    End Select

End Function


'**+
'Nome: ShowErrorLog
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Mostra il dialogo di informazioni su l'ultimo errore
'bloccante verificatosi durante l'esecuzione del programma
'**/
Private Sub ShowErrorLog()
    Load frmErrorLog
    frmErrorLog.DMTErrorContol.MainProgram.Comments = App.Comments
    frmErrorLog.DMTErrorContol.MainProgram.Company = App.CompanyName
    frmErrorLog.DMTErrorContol.MainProgram.Copyright = App.LegalCopyright
    frmErrorLog.DMTErrorContol.MainProgram.Description = App.FileDescription
    frmErrorLog.DMTErrorContol.MainProgram.FileName = App.EXEName
    frmErrorLog.DMTErrorContol.MainProgram.Version = App.Major & "." & App.Minor & "." & App.Revision
    frmErrorLog.DMTErrorContol.ErrorNumber = Err.Number
    frmErrorLog.DMTErrorContol.ErrorDescription = Err.Description
    frmErrorLog.DMTErrorContol.Show
    frmErrorLog.Show vbModal
    End
End Sub


'**+
'Nome: OnBeforeOpenDoc
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazioni da effettuare prima dell'apertura del documento.
'**/
Private Sub OnBeforeOpenDoc()
Dim cl As DmtGridCtl.dgColumnHeader
    'Inserire qui le
    'inizializzazioni da effettuare prima dell'apertura del documento.
    
    'Inizializza la DmtDocs
    If oDoc Is Nothing Then
        'Crea una istanza dell'oggetto cDocument della DmtDocs
        Set oDoc = New DmtDocs.cDocument
        'Crea una istanza dell'oggetto DocChangeNotify per intercettare le variazioni
        'di valori che avvengono all'interno della struttura di tabelle della DmtDocs
        Set oDocChangeNotify = New DocChangeNotify
        
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set oDoc.Connection = TheApp.Database.Connection
        'Imposta l'oggetto a cui notificare le variazioni di valori
        Set oDoc.ChangeNotifyObj = oDocChangeNotify
        
        'Imposta il tipo di documento da gestire (Ved. tabella TipoOggetto)
        oDoc.SetTipoOggetto 15 'Ordine da cliente
        'Imposta la funzione di DMT che gestisce il tipo di documento interessato (Ved. tabella Funzione)
        oDoc.IDFunzione = 128 'Ordine da cliente
        'Preleva il nome delle tabelle in cui sono memorizzati i dati del documento
        oDoc.TablesNames oDoc.IDTipoOggetto, sTabellaTestata, sTabellaDettaglio, sTabellaIVA, sTabellaScadenze
        
        'Imposta l'IDAzienda per cui si genererà il documento
        oDoc.IDAzienda = TheApp.IDFirm
        'Imposta l'IDFiliale per cui si genererà il documento
        oDoc.IDFiliale = TheApp.Branch
        'Imposta il tipo di anagrafica da utilizzare per il documento corrente
        'Nel caso di documenti che hanno come soggetto un fornitore bisogna indicare il valore 3
        oDoc.IDAttivitaAzienda = GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch)
        oDoc.IDTipoAnagrafica = 2 'Cliente
        
        'Imposta l'identificativo dell'utente corrente
        oDoc.IDUtente = TheApp.IDUser
        oDoc.Descrizione = GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto)
        oDoc.UpdateOnlyModified = False
        
    End If

    'Inizializza la combo dei listini
    With cboListino
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT IDListino, Listino FROM Listino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino = 0"
        .SQL = .SQL & " ORDER BY Listino"
    End With
    'Inizializza la combo dei listini
    With Me.cboListinoAzienda
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT IDListino, Listino FROM Listino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino = 0"
        .SQL = .SQL & " ORDER BY Listino"
    End With
    With Me.cboBancaAzienda
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDBancaPerAnagrafica"
        .DisplayField = "BancaPerAnagrafica"
        .SQL = "SELECT BancaPerAnagrafica.IDBancaPerAnagrafica, BancaPerAnagrafica.BancaPerAnagrafica "
        .SQL = .SQL & "FROM Anagrafica INNER JOIN "
        .SQL = .SQL & "Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica INNER JOIN "
        .SQL = .SQL & "BancaPerAnagrafica ON Azienda.IDAnagrafica = BancaPerAnagrafica.IDAnagrafica "
        .SQL = .SQL & " WHERE ((BancaPerAnagrafica.IDAzienda = " & TheApp.IDFirm & "))"
        .SQL = .SQL & " ORDER BY BancaPerAnagrafica.BancaPerAnagrafica"
    End With

    With Me.cboMagazzino
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .SQL = "SELECT IDMagazzino, Magazzino FROM Magazzino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " ORDER BY Magazzino"
    End With
    With Me.cboSezionale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT  Sezionale.IDSezionale, Sezionale.Sezionale, RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "FROM Sezionale INNER JOIN "
        .SQL = .SQL & "RegistroIvaPerTipoOggetto ON Sezionale.IDRegistroIva = RegistroIvaPerTipoOggetto.IDRegistroIva AND "
        .SQL = .SQL & "Sezionale.IDFiliale = RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & oDoc.IDTipoOggetto
        .SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & oDoc.IDFiliale
        '.SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDTipoModulo=1 "
        '.SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDRegistroIVA=1 "
    End With
    
    'Inizializza la combo dei pagamenti
    With cboPagamento
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDPagamento"
        .DisplayField = "Pagamento"
        .SQL = "SELECT IDPagamento, Pagamento FROM Pagamento"
        .SQL = .SQL & " ORDER BY Pagamento"
    End With
    
    With Me.cboPorto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDPorto"
        .DisplayField = "Porto"
        .SQL = "SELECT IDPorto, Porto FROM Porto"
        .SQL = .SQL & " ORDER BY Porto"
    End With
    With Me.cboRaggrFatturato
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRaggruppamentoFatturato"
        .DisplayField = "RaggruppamentoFatturato"
        .SQL = "SELECT * FROM RaggruppamentoFatturato"
        .SQL = .SQL & " ORDER BY RaggruppamentoFatturato"
    End With
    With Me.cboTrasporto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDTipoSpedizione"
        .DisplayField = "TipoSpedizione"
        .SQL = "SELECT IDTipoSpedizione, TipoSpedizione FROM TipoSpedizione"
        .SQL = .SQL & " ORDER BY TipoSpedizione"
    End With
    With Me.cboVettore
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .SQL = "SELECT IDVettore, Vettore FROM Vettore"
        .SQL = .SQL & " ORDER BY Vettore"
    End With
    With Me.cboAspettoEsteriore
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDAspettoEsterioreArticolo"
        .DisplayField = "AspettoEsterioreArticolo"
        .SQL = "SELECT IDAspettoEsterioreArticolo, AspettoEsterioreArticolo FROM AspettoEsterioreArticolo"
        .SQL = .SQL & " ORDER BY AspettoEsterioreArticolo"
    End With

    With Me.cboAliquotaArticolo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Codice"
        .SQL = "SELECT IDIva, Codice FROM Iva"
        .SQL = .SQL & " ORDER BY Codice"
    End With

    With Me.cboIvaCliente
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT IDIva, Iva FROM Iva"
        .SQL = .SQL & " ORDER BY AliquotaIva"
    End With

    With Me.CboAliquotaImballo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Codice"
        .SQL = "SELECT IDIva, Codice FROM Iva"
        .SQL = .SQL & " ORDER BY Codice"
    End With

    With Me.cboUnitaDiMisura
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .SQL = "SELECT IDUnitaDiMisura, UnitaDiMisura FROM UnitaDiMisura"
        .SQL = .SQL & " ORDER BY UnitaDiMisura"
    End With

    'Inizializza il controllo Codice-Descrizione per la ricerca dei clienti
    With cdAnagrafica
        Set .Application = TheApp
        Set .Database = TheApp.Database
        .HwndContainer = Me.hwnd
        .CodeCaption4Find = "Cognome / Ragione sociale"
        .CodeField = "Anagrafica"
        .CodeIsNumeric = False

        .DescriptionCaption4Find = "Nome"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .IDExecuteFunction = 29 'Anagrafica
    End With
   
    
    'Imballo
    With Me.CDImballo
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
   
    'Imballo primario
    With Me.CDImballoPrimario
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
   
    
    'Pedana
    With Me.CDPedana
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
     
    'Pianale
    With Me.CDArticoloPianale
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
    
    'Prolunga
    With Me.CDArticoloProlunga
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
     
    'Tipo di Pedana
    With Me.CDTipoPedana
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceTipoPedana"
        .DescriptionField = "TipoPedana"
        .KeyField = "IDRV_POTipoPedana"
        .TableName = "RV_POIETipoPedana"
        .Filter = "IDAzienda = " & m_App.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice tipo pedana"
        .DescriptionCaption4Find = "Descrizione tipo pedana"
        .IDExecuteFunction = GET_FUNZIONE(fnGetTipoOggetto("RV_POTipoPedana")) 'Articoli
        .CodeIsNumeric = False
    End With
    
    Set Me.ACSSocio.Connection = TheApp.Database.Connection
    ACSSocio.ApplicationName = App.Title
    ACSSocio.Client = App.EXEName
    ACSSocio.IDFirm = TheApp.IDFirm
    ACSSocio.IDUser = TheApp.IDUser
    ACSSocio.UserName = TheApp.User
    ACSSocio.SearchType = DmtSearchSuppliers
    ACSSocio.HwndContainer = Me.hwnd

    'Articolo
    With Me.CDArticolo
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
   

    With Me.cboCalibro
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POCalibro"
        .DisplayField = "Calibro"
        .SQL = "SELECT * FROM RV_POCalibro"
        .SQL = .SQL & " ORDER BY Calibro"
    End With
  
    With Me.cboTipoCategoria
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoCategoria"
        .DisplayField = "TipoCategoria"
        .SQL = "SELECT * FROM RV_POTipoCategoria"
        .SQL = .SQL & " ORDER BY TipoCategoria"
    End With
  
    With Me.cboTipoLavorazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .SQL = "SELECT * FROM RV_POTipoLavorazione"
        .SQL = .SQL & " ORDER BY TipoLavorazione"
    End With


    With Me.cboUMRigaOrdine
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoUMOrdine"
        .DisplayField = "TipoUMOrdine"
        .SQL = "SELECT * FROM RV_POTipoUMOrdine"
        .SQL = .SQL & " ORDER BY TipoUMOrdine"
    End With


 
    'Inizializza la ListView contenente l'elenco degli articoli presenti in un documento
    With lvwArticoli
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        
        .ColumnHeaders.Add , , "NR", 300                            '00
        .ColumnHeaders.Add , , "Tipo", 300                          '01
        .ColumnHeaders.Add , , "Rett.", 300                         '02
        .ColumnHeaders.Add , , "ID Art.", 100                       '03
        .ColumnHeaders.Add , , "Cod. Articolo", 1000                '04
        .ColumnHeaders.Add , , "Articolo", 2000                     '05
        .ColumnHeaders.Add , , "Q.tà", 1000, lvwColumnRight         '06
        .ColumnHeaders.Add , , "Imp. Unit.", 1000, lvwColumnRight   '07
        .ColumnHeaders.Add , , "Sc. 1", 700, lvwColumnRight         '08
        .ColumnHeaders.Add , , "Sc. 2", 700, lvwColumnRight         '09
        .ColumnHeaders.Add , , "Imp.Un.Net.", 1000, lvwColumnRight  '10
        
        .ColumnHeaders.Add , , "IDImballoPrimario", 100, lvwColumnLeft   '11
        .ColumnHeaders.Add , , "Codice imballo primario", 1300, lvwColumnLeft   '12
        .ColumnHeaders.Add , , "Descrizione imballo primario", 2000, lvwColumnLeft   '13
        
        .ColumnHeaders.Add , , "ID Art. Ped.", 100                  '14
        .ColumnHeaders.Add , , "Codice Pedanda", 1000               '15
        .ColumnHeaders.Add , , "Pedana", 2000                       '16
        .ColumnHeaders.Add , , "Q.tà pedana", 1000, lvwColumnRight  '17
        .ColumnHeaders.Add , , "ID Tipo U.M. riga.", 100            '18
        .ColumnHeaders.Add , , "U.M. riga", 1000                    '19
        
        
        
        .ColumnHeaders.Add , , "Colli", 800, lvwColumnRight         '20
        .ColumnHeaders.Add , , "Pezzi", 1000, lvwColumnRight        '21
        .ColumnHeaders.Add , , "Peso lordo", 1000, lvwColumnRight   '22
        .ColumnHeaders.Add , , "Tara", 1000, lvwColumnRight         '23
        '.ColumnHeaders.Add , , "Cod. IVA", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "% IVA", 500, lvwColumnRight         '24
        .ColumnHeaders.Add , , "Netto riga", 1300, lvwColumnRight   '25
        .ColumnHeaders.Add , , "Lordo riga", 1300, lvwColumnRight   '26
        
        .ColumnHeaders.Add , , "Raggrup. ord.", 1300, lvwColumnLeft   '27
        .ColumnHeaders.Add , , "Categoria", 1300, lvwColumnLeft   '28
        .ColumnHeaders.Add , , "Calibro", 1300, lvwColumnLeft   '29
        .ColumnHeaders.Add , , "Tipo lavorazione", 1300, lvwColumnLeft   '30




    End With
    
    'Inizializza la ListView contenente il castelletto IVA del documento
    With lvwIVA
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "Iva", 600
        .ColumnHeaders.Add , , "Descr. Iva", 1000
        .ColumnHeaders.Add , , "Imponibile", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Imposta", 1000, lvwColumnRight
    End With
    
    'Inizializza la ListView contenente l'elenco delle scadenze presenti in un documento
    With lvwScadenze
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "Data", 1000
        .ColumnHeaders.Add , , "Importo", 1000, lvwColumnRight
    End With



    With Me.cboValuta
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDValuta"
        .DisplayField = "Valuta"
        .SQL = "SELECT * FROM Valuta"
        .SQL = .SQL & " ORDER BY Valuta"
    End With
    
    With Me.cboCambioValuta
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDCambio"
        .DisplayField = "DataCambio"
        .SQL = "SELECT * FROM Cambio"
    End With

    With Me.cboVettoreSuccessivo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .SQL = "SELECT * FROM Vettore ORDER BY Vettore"
        .Fill
    End With

    With Me.cboLuogoPresaMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT * FROM SitoPerAnagrafica  "
        .SQL = .SQL & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .SQL = .SQL & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With

    With Me.cboTipoOrdine
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoOrdine"
        .DisplayField = "TipoOrdine"
        .SQL = "SELECT * FROM RV_POTipoOrdine  "
        .SQL = .SQL & " ORDER BY TipoOrdine "
        .Fill
    End With


    'Inizializza il controllo Codice-Descrizione per la ricerca dei clienti
    With Me.CDAgenteTesta
        Set .Application = TheApp
        Set .Database = TheApp.Database  '<-- Notare la connessione DMTDBLib.Database
        .HwndContainer = Me.hwnd
        .CodeCaption4Find = "Cognome / Ragione sociale"
        .CodeField = "Anagrafica"
        .CodeIsNumeric = False
        .DescriptionCaption4Find = "Nome"
        .DescriptionField = "Nome"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepAgente"
        .Filter = "IDAzienda = " & TheApp.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .IDExecuteFunction = 29 'Anagrafica
    End With

    Set Me.ACSCliente.Connection = TheApp.Database.Connection
    ACSCliente.ApplicationName = App.Title
    ACSCliente.Client = App.EXEName
    ACSCliente.IDFirm = TheApp.IDFirm
    ACSCliente.IDUser = TheApp.IDUser
    ACSCliente.UserName = TheApp.User
    ACSCliente.SearchType = DmtSearchCustomers
    ACSCliente.HwndContainer = Me.hwnd

    Set Me.ACSAnaDest.Connection = TheApp.Database.Connection
    ACSAnaDest.ApplicationName = App.Title
    ACSAnaDest.Client = App.EXEName
    ACSAnaDest.IDFirm = TheApp.IDFirm
    ACSAnaDest.IDUser = TheApp.IDUser
    ACSAnaDest.UserName = TheApp.User
    ACSAnaDest.SearchType = DmtSearchCustomers
    ACSAnaDest.HwndContainer = Me.hwnd
    
    
    GET_TABELLE_ARTICOLO
    
    With Me.cboReportPedNew
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POConfigurazioneEtichetta"
        .DisplayField = "ConfigurazioneEtichetta"
        .SQL = "SELECT * FROM RV_POConfigurazioneEtichetta "
        .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
        .SQL = .SQL & " AND IDRV_POTipoEtichetta=2 "
        .SQL = .SQL & " ORDER BY ConfigurazioneEtichetta "
        .Fill
    End With

    With Me.cboReportNew
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POConfigurazioneEtichetta"
        .DisplayField = "ConfigurazioneEtichetta"
        .SQL = "SELECT * FROM RV_POConfigurazioneEtichetta "
        .SQL = .SQL & "WHERE IDAzienda=" & TheApp.IDFirm
        .SQL = .SQL & " AND IDRV_POTipoEtichetta=1 "
        .SQL = .SQL & " ORDER BY ConfigurazioneEtichetta "
        .Fill
    End With
    
    'Inizializza la ListView contenente l'elenco delle commissioni presenti nel documento
    With lvwCommissioni
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False
        
        .ColumnHeaders.Add , , "Commissione", 3300
        .ColumnHeaders.Add , , "%", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Q.tà", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Importo", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Tipo pedana", 1000, lvwColumnLeft
        .ColumnHeaders.Add , , "% Liq. (su merce netta)", 1500, lvwColumnRight
        .ColumnHeaders.Add , , "Tipo calcolo", 1000, lvwColumnLeft
    End With
End Sub

'**+
'Nome: Start
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione e procedura di avvio
'**/
Private Sub Start()
    Dim OLDCursor As Integer
    Dim ToolID As Integer
    Dim Field As DmtDocManLib.Field
    Dim oActivity As IActivity
    Dim o As Activity
    Dim oFilter As Filter

        
    
    'Apertura del documento
    If Len(m_ExtendedDatabase) > 0 Then
        'Apre un nuovo documento usando il database esteso
        Set m_Document = m_App.OpenDocument(m_DocType, m_ExtendedDatabase)
    Else
        'Apre un nuovo documento usando il database diamante
        Set m_Document = m_App.OpenDocument(m_DocType)
    End If
    
    
    'NOTA: Con la sottostante proprietà settata a TRUE i metodi OnXXXDocumentsLink()
    'non sono più necessari in quanto il modello ad oggetti si occupa della gestione
    'dei sottodocumenti.
    '
    'Abilita la gestione automatica degli eventuali DocumentsLink
    m_Document.EnableRefreshDocumentsLinks = True
    
    
    'Clessidra
    OLDCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    Caption = m_App.Caption
    
    'Inizializzazione del controllo ActiveBar
    InitMenuBar ToolID
    InitToolBar ToolID
    ActivateBarButtons BTN_ALL, True
    
    'Inizializzazione del riquadro attività
    With ActivityBox
        .Activities.Clear
        
        'Aggiunge l'attività dei reports
        Set oActivity = .Activities.Add("DmtActBoxLib.ReportsActivity", "Reports")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID, TheApp.IDFirm
        Set o = oActivity
        Set oReportsActivity = o.InternalClass
        
        'Aggiunge l'attività dei filtri
        Set oActivity = .Activities.Add("DmtActBoxLib.FiltersActivity", "Filters")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType
        Set o = oActivity
        Set oFiltersActivity = o.InternalClass
        
        'Aggiunge l'attività delle viste tabellari
        Set oActivity = .Activities.Add("DmtActBoxLib.TableViewsActivity", "TableViews")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID
        Set o = oActivity
        Set oTableViewsActivity = o.InternalClass

        'Aggiunge l'attività delle esportazioni
        Set oActivity = .Activities.Add("DmtActBoxLib.ExportActivity", "Export")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load m_DocType.ID
        Set o = oActivity
        Set oExportActivity = o.InternalClass
        
        'Aggiunge l'attività del supporto tecnico
        Set oActivity = .Activities.Add("DmtActBoxLib.SupportActivity", "Support")
        Set oActivity.Connection = TheApp.Database.InternalConnection
        oActivity.Load
        Set o = oActivity
        Set oSupportActivity = o.InternalClass
        
        'attiva/disattiva la visualizzazione delle attività
        EnableDOMActivitiesItems
        
        'imposta quale attività deve essere attivata per default
        If m_DefaultActivity <> "" Then
            Set .CurrentActivity = .Activities(m_DefaultActivity)
        End If
        
        'ridisegna il controllo
        .Redraw = True
    End With


    'Lettura impostazioni dal registry
    ReadRegistrySettings
        
    'Aggiunge due filtri temporanei, uno per le ricerche temporanee
    'e uno per la stampa in modalità form
    m_DocType.AddFilter "Temp"
    m_DocType.AddFilter "Form"

    
    'Imposta dei filtri fissi per l'estrazione dei documenti
    'da visualizzare all'interno della manutenzione
    
    'Imposta il filtro sull'IDAzienda
    m_DocType.AddFixedCondition "IDAzienda = " & TheApp.IDFirm
    'Imposta il filtro sull'IDFiliale
    m_DocType.AddFixedCondition "IDFiliale = " & TheApp.Branch
    
    
    
    
    'Inizializzazioni da fare prima dell'apertura del documento
    
    ConnessioneDiamanteADO
    
    ParametroUMRigaOrdine
    ParametroImballo
    ParametroLavorato
    ParametroGrezzo
    ParametroSocio
    ParametroObbligatorio
    ParametroTipoArrotondamento
    ParametroTipoScarto
    ParametroTipoCaloPeso
    ParametroTipoAumentoPeso
    ParametroNuovoCalcolo
    ParametroAndamentoOrdine
    ParametroGestioneOrdineVivaio
    ParametroSezionalePrelievo
    ParametroPesoArticolo
    
    GET_PARAMETRI_IVA_IMBALLO
    DISEGNA_FORM
    GET_MODULO_ATTIVATO MODULO_CODICE, 80
    GET_PAR_CONN_ALY
    GET_PARAMETRI_FIDO_ALYANTE
    AltriParametriCooperativa
    'OnStart
    
    OnBeforeOpenDoc
    
    LINK_OGGETTO_ORDINE_PADRE_REGISTRY = 0
     
    If Len(m_App.Caller) > 0 And (m_App.CallerFieldValue) > 0 Then
        '-------------------------------------------------
        '     Il programma è stato chiamato da un link.
        '-------------------------------------------------
        
        'In tal caso occorre mostrare in modalità variazione il record richiesto dal programma client.
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next

        'Imposta una condizione di ricerca basata sull'ID del record richiesto dal programma client.
        m_DocType.Fields("IDOggetto").Value = m_App.CallerFieldValue '171600
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"

        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")

        'Inidica, nel caso di esegui gestione, se riportare il valore corrente al chiamante
        
        bNotReturnValue = CBool(Val(GetSetting(REGISTRY_KEY, App.EXEName, "NoReturnValue", "0")))
        
        
        LINK_OGGETTO_ORDINE_PADRE_REGISTRY = Val(GetSetting(REGISTRY_KEY, App.EXEName, "IDOggettoOrdinePadre"))
        
        'Si comunica al documento quale filtro eseguire all'avvio.
        Set m_Document.ActiveFilter = m_ActiveFilter
        
    Else
        '---------------------------------------------------
        '     Il programma non è stato chiamato da un link.
        '---------------------------------------------------
        
        'Il filtro attivo alla partenza è quello predefinito
        For Each oFilter In m_DocType.Filters
            If oFilter.ID = oFiltersActivity.DefaultFilterID Then
                Set m_ActiveFilter = m_DocType.Filters(oFilter.Name)
                Exit For
            End If
        Next
       'Si comunica al documento quale filtro eseguire all'avvio.
        Set m_Document.ActiveFilter = m_ActiveFilter
        m_Document.Dataset.Recordset.Sort = "RV_PODataOrdinePadre Desc, RV_PONumeroOrdinePadre Desc, RV_PONumeroListaPrelievo Desc"
        Set BrwMain.Recordset = m_Document.Dataset.Recordset
        
        bNotReturnValue = 1
    End If
    
    'Prima di aprire il documento occorre comunicargli qual'è il campo chiave primaria.
    m_Document.PrimaryKey = "IDOggetto" '& m_Document.TableName

    'Apertura del documento.
    m_Document.OpenDoc
    
    
    'Questa impostazione serve per conservare le impostazioni grafiche
    BrwMain.IDUser = m_App.IDUser
    'Permette di gestire l'evento BrwMain_OnApplyFilter
    BrwMain.AutoFiltering = False
    'Con questa impostazione la dmtGrid NON effettua mai il Move sul documento.
    'Questo pertanto andrà forzato in BrwMain_DblClick e BrwMain_KeyDown.
    BrwMain.EnableMove = False
    'Inizializza le colonne da visualizzare nella griglia
    If m_DocType.DefaultTableView Is Nothing Then
        Err.Raise ERR_NO_DEFAULT_TABLEVIEW, , "Default TableView not found"
    Else
        BrwMain.LoadColumns m_DocType.DefaultTableView
        SetVisibilityIDFields
    End If
    
    'Crea i campi per la ricerca.
    CreateBrowserConditions
    'Assegnazione del riferimento alla fonte dati (binding sul recordset del documento)
    
    'rif14

    
    'Set BrwMain.Recordset = m_Document.Dataset.Recordset
    Set BrwMain.Recordset = m_Document.Data
    
    
            
     'Viene inizializzato il dialogo di stampa
    'With DmtPrnDlg
    '    Set .Application = m_App
    '    Set .DocType = m_DocType
    'End With
    
    
    
    'Ripulisco la tabella semaforo.
    'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    
    'Evita il blocco della toolbar
    'BarMenu.ResetHooks
    
    Screen.MousePointer = OLDCursor
    
    
End Sub




'**+
'Nome: ConditionType
'
'Parametri: DBType è il valore di DMTDocManLib.Field.DBType e rappresenta
'           il tipo di dato corrispondente all'oggetto Field in base dati.
'
'Valori di ritorno: una costante di tipo ConditionTypeConstants usata dalla Browse
'                   per costruire una condizione di ricerca.
'
'Funzionalità: Trasforma una costante DBType in una costante compatibile ConditionTypeConstants
'**/
Private Function ConditionType(ByVal DBType As Integer) As DmtGridCtl.ConditionTypeConstants
    Select Case DBType

        'dbTypeCHAR, dbTypeVARCHAR, dbTypeWCHAR, dbTypeWVARCHAR
        Case 1, 12, -8, -9
            ConditionType = dgCondTypeText
       
        'dbTypeNUMERIC, dbTypeDECIMAL, dbTypeINTEGER, dbTypeSMALLINT, dbTypeFLOAT
        'dbTypeREAL, dbTypeDOUBLE, dbTypeBIGINT, dbTypeTINYINT
        Case 2, 3, 4, 5, 6, 7, 8, -5, -6
            ConditionType = dgCondTypeNumber
            
        'dbTypeTIMESTAMP  ////NOTA: Se si desidera un campo dmCondTypeTime occorre gestirlo ad Hoc.
        Case 135
            ConditionType = dgCondTypeDate
    
        'dbTypeBIT
        Case -7, 11
            ConditionType = dgCondTypeBoolean
            
    End Select
End Function

'**+
'Nome: CreateBrowserConditions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Crea automaticamente i campi per la ricerca (modalità DefineFilter)
'              a partire dai campi non ID del documento.
'**/
Private Sub CreateBrowserConditions()
    Dim Field As DmtDocManLib.Field
    Dim Cond As DmtGridCtl.dgCondition
    
    
    If Right(Trim(MenuOptions.ConnectionString), 1) = ";" Then
        Me.BrwMain.ConnectionString = MenuOptions.ConnectionString & "User Id=" & m_App.User & ";Password=" & m_App.Password
    Else
        Me.BrwMain.ConnectionString = MenuOptions.ConnectionString & ";" & "User Id=" & m_App.User & ";Password=" & m_App.Password
    End If
    'Vengono creati automaticamente i campi per la ricerca.
    'In una applicazione specifica questo codice andrà sostituito integralmente per definire
    'dei campi di ricerca ad hoc.
    
    'Non viene visualizzata la Check Intervallo perchè attualmente
    'il modello ad oggetti non prevede la gestione di filtri con
    'clausole BETWEEN.


    BrwMain.Conditions.Clear

    BrwMain.Conditions.WidthConditions = 250
    BrwMain.Conditions.WidthFields = 300
    BrwMain.Conditions.WidthIntervals = 100
    
    BrwMain.Title.BackColor = vb3DFace
    BrwMain.Title.ForeColor = vbBlack
    BrwMain.Title.Font.Bold = True

    BrwMain.Conditions.Add "Group1", "Dati generali documento", ""
    BrwMain.Conditions("Group1").IsHeader = True
        Set Cond = BrwMain.Conditions.Add("RV_PODataOrdinePadre", "Data documento padre", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
            Cond.RangeChecked = True
            Cond.FromValue = DateAdd("m", -3, Date)
            Cond.ToValue = Date
        Set Cond = BrwMain.Conditions.Add("RV_PONumeroOrdinePadre", "Numero documento padre", m_DocType.TableName, False, True, , dgCondTypeNumber)
            Cond.Indentation = 20
        
        Set Cond = BrwMain.Conditions.Add("Link_doc_sezionale", "Sezionale", m_DocType.TableName, False, False, , dgCondTypeComboDB)
            Cond.Indentation = 20
            Cond.RecordSource = "SELECT * FROM Sezionale WHERE IDFiliale=" & TheApp.Branch & "  ORDER BY Sezionale"
            Cond.DisplayField = "Sezionale"
            Cond.KeyField = "IDSezionale"
        Set Cond = BrwMain.Conditions.Add("Doc_data", "Data documento", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
            Cond.RangeChecked = True
        Set Cond = BrwMain.Conditions.Add("Doc_numero", "Numero documento", m_DocType.TableName, False, True, , dgCondTypeNumber)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("Nom_ragione_sociale_o_cognome", "Anagrafica", m_DocType.TableName, True, False, , dgCondTypeText)
           Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("Link_Vet_vettore", "Vettore", m_DocType.TableName, False, False, , dgCondTypeComboDB)
            Cond.Indentation = 20
            Cond.RecordSource = "SELECT * FROM Vettore ORDER BY Vettore "
            Cond.DisplayField = "Vettore"
            Cond.KeyField = "IDVettore"
            
        Set Cond = BrwMain.Conditions.Add("RV_POTargaAutomezzo", "Targa automezzo", m_DocType.TableName, False, False, , dgCondTypeText)
           Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("Doc_ordine_chiuso", "Ordine chiuso", m_DocType.TableName, False, False, , dgCondTypeBoolean)
           Cond.FromValue = "NO"
           Cond.Indentation = 20

           
        Set Cond = BrwMain.Conditions.Add("RV_POOrdineCompletato", "Ordine completato", m_DocType.TableName, False, False, , dgCondTypeBoolean)
           Cond.Indentation = 20
           
        Set Cond = BrwMain.Conditions.Add("RV_POIDTipoOrdine", "Tipo ordine", m_DocType.TableName, False, False, , dgCondTypeComboDB)
            Cond.Indentation = 20
            Cond.RecordSource = "SELECT * FROM RV_POTipoOrdine ORDER BY TipoOrdine  "
            Cond.DisplayField = "TipoOrdine"
            Cond.KeyField = "IDRV_POTipoOrdine"
            
        Set Cond = BrwMain.Conditions.Add("Doc_age_ragione_sociale", "Agente", m_DocType.TableName, False, False, , dgCondTypeText)
           Cond.Indentation = 20
            
            
    BrwMain.Conditions.Add "Group3", "Altra destinazione", ""
    BrwMain.Conditions("Group3").IsHeader = True
        Set Cond = BrwMain.Conditions.Add("SitoPerAnagrafica", "Destinazione diversa", m_DocType.TableName, True, False, , dgCondTypeText)
           Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("RV_PODataArrivoMerce", "Data arrivo merce", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("RV_POOraArrivoMerce", "Ora arrivo merce", m_DocType.TableName, False, True, , dgCondTypeNumber)
            Cond.Indentation = 20

    BrwMain.Conditions.Add "Group4", "Presa di luogo merce", ""
    BrwMain.Conditions("Group4").IsHeader = True
        Set Cond = BrwMain.Conditions.Add("RV_POIDLuogoPresaMerce", "Luogo di presa merce", m_DocType.TableName, False, False, , dgCondTypeComboDB)
            Cond.Indentation = 20
            Cond.RecordSource = "SELECT * FROM SitoPerAnagrafica  "
            Cond.RecordSource = Cond.RecordSource & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
            Cond.RecordSource = Cond.RecordSource & " ORDER BY SitoPerAnagrafica "
            Cond.DisplayField = "SitoPerAnagrafica"
            Cond.KeyField = "IDSitoPerAnagrafica"
        Set Cond = BrwMain.Conditions.Add("RV_PODataArrivoMerceLuogo", "Data arrivo merce", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("RV_POOraArrivoMerceLuogo", "Ora arrivo merce", m_DocType.TableName, False, True, , dgCondTypeNumber)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("RV_POIDTrasportatoreSuccessivo", "Vettore successivo", m_DocType.TableName, False, False, , dgCondTypeComboDB)
            Cond.Indentation = 20
            Cond.RecordSource = "SELECT * FROM Vettore ORDER BY Vettore  "
            Cond.DisplayField = "Vettore"
            Cond.KeyField = "IDVettore"
    
    BrwMain.Conditions.Add "Group2", "Altri dati", ""
    BrwMain.Conditions("Group2").IsHeader = True
        Set Cond = BrwMain.Conditions.Add("Doc_data_presso_nom", "Data ordine cliente", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("NumeroPressoNom", "Numero ordine cliente", m_DocType.TableName, False, True, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("Doc_data_prevista_evasione", "Data partenza", m_DocType.TableName, False, True, , dgCondTypeDate)
            Cond.Indentation = 20
        
    BrwMain.Conditions.Add "Group5", "Annotazioni", ""
    BrwMain.Conditions("Group5").IsHeader = True
        Set Cond = BrwMain.Conditions.Add("Doc_annotazioni_variazio", "Annotazioni di fatturazione", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("RV_POAnnotazioniInterna", "Annotazioni interne", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
        Set Cond = BrwMain.Conditions.Add("RV_PODescrizioneCorpoDocEv", "Descrizioni corpo documento", m_DocType.TableName, False, False, , dgCondTypeText)
            Cond.Indentation = 20
'    BrwMain.Conditions.Add "Group6", "Altro", ""
'    BrwMain.Conditions("Group6").IsHeader = True
'        Set Cond = BrwMain.Conditions.Add("IDOggetto", "IDOggetto", m_DocType.TableName, False, False, , dgCondTypeText)
'            Cond.Indentation = 20

End Sub

'**+
'Nome: Export
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue l'esportazione del documento con controllo di errore
'**/
Private Sub ExportDocument(ByVal Appl As Long)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    m_Document.Export m_Report, Appl
    Screen.MousePointer = OLDCursor
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    If Err.Number = 20507 Then
        'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
        sbMsgInfo "File di report non trovato", m_App.FunctionName
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
End Sub

'**+
'Nome: PrintDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue la stampa del documento con controllo di errore per nessuna stampante
'definita
'**/
Private Sub PrintDocument(ByVal ToolName As String)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    '**+ Riferimento al cursore corrente
    OLDCursor = Screen.MousePointer
    
    '**+ Inizializzazione selezioni di stampa
    'm_Report.Copies = GET_NUMERO_COPIE
    'm_Report.Orientation = GET_ORIENTAMENTO
    'm_Report.PrinterName = ""
    
    If ToolName = "Mnu_Print" Then
        '**+ stampa con dialogo
        Set DmtPrnDlg.Connection = TheApp.Database.Connection
        DmtPrnDlg.IDFiliale = TheApp.Branch
        DmtPrnDlg.IDTipoOggetto = fnGetTipoOggetto
        DmtPrnDlg.Copies = m_Report.Copies
        DmtPrnDlg.Refresh
        
        DmtPrnDlg.Show
        If Not DmtPrnDlg.Cancel Then
            Screen.MousePointer = vbHourglass
            m_Report.Copies = DmtPrnDlg.Copies
            
            m_Report.DoPrint DmtPrnDlg.PrinterName
            'm_Document.DoPrint m_Report
        End If

    Else
        'Stampa diretta
        Screen.MousePointer = vbHourglass
        m_Document.DoPrint m_Report
    End If
    
    Screen.MousePointer = OLDCursor
    Exit Sub

errHandler:
    Screen.MousePointer = OLDCursor
    If Err.Number = vbObjectError + 36 Then
        ' errore generato all'interno della DMTDocManLib per nessuna stampante
        sbMsgInfo "Non è possibile ottenere informazioni sulla stampante." & Chr(13) & "Controllare che sia installata correttamente", m_App.FunctionName
    ElseIf Err.Number = vbObjectError + 4 Then
        'Si è annullata la stampa.
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
    
End Sub

'**+
'Nome: DoNewDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Procedure per la richiesta di un nuovo documento
'**/
Private Function DoNewDocument() As Integer
    
    '------------------------------------------------
    'Inserire qui se occorre del codice specifico
    'per la manutenzione.
    '------------------------------------------------
    
    DoNewDocument = ChooseAboutSaving
End Function

'**+
'Nome: WriteStatusBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Scrive una stringa di testo nella StatusBar
'**/
Private Sub WriteStatusBar(ByVal sTesto As String)
    If stbStatusbar.Style = sbrSimple Then
        stbStatusbar.SimpleText = sTesto
    Else
        stbStatusbar.Panels(1).Text = sTesto
    End If
End Sub

'**+
'Nome: FormUnload
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue i controlli alla richiesta di abbandono del form
'**/
Private Function FormUnload() As Integer
    Dim sMessage As String
    Dim sMessage1 As String
    Dim lIDField As Long

    
    If m_Changed Then
        Select Case ChooseAboutSaving
            Case vbCancel
                FormUnload = 1
                Exit Function
            Case vbYes
                OnSave
                'Se la registrazione non è andata a buon fine
                'esce e non chiude il programma
                If Not m_Saved Then
                    FormUnload = 1
                    Exit Function
                End If
        End Select
    End If
        
    
    If m_PreviewWindowHandle > 0 Then
        ClosePreview
    End If
    
    SaveRegistrySettings
    
    'Se il programma è stato chiamato da un link occorre restituire l'ID del record attivo
    'all'applicazione chiamante.
    If Len(m_App.Caller) > 0 Then
        'Il programma è stato chiamato da un link.
       'MsgBox m_App.Caller
        If m_App.Caller = "RV_POMenuGreenTop" Then Exit Function
        'Se non verrà correttamente selezionato un elemento sarà restituito il valore -1 all'applicazione client.
        lIDField = -1
        
        'Se il documento è vuoto non si deve far nulla.
        'Se la browse è in modalità Filter Definition non formula la domanda di riporto dei dati nel programma chiamante.
        'If (Not (m_Document.EOF And m_Document.BOF)) And (BrwMain.GuiMode <> dgFilterDefinition) Then
        If (oDoc.IDOggetto > 0) And (BrwMain.GuiMode <> dgFilterDefinition) Then
            'MsgBox bNotReturnValue
            If Abs(bNotReturnValue) > 0 Then
                'MsgBox oDoc.IDOggetto
                'ATTENZIONE: La stringa sMessage1 deve essere personalizzata a seconda dei casi!!!
                sMessage1 = " il " & m_DocType.Name
                sMessage = sMessage1 & " """ & Caption2Display(False) & """"
                
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add sMessage, 1
                   
                'Viene chiesto se si intende riportare il record corrente al programma chiamante.
                If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYPASTE), m_App.FunctionName) = vbYes Then
                    'Legge l'ID del record corrente affinchè venga riportato all'applicazione chiamante.
                    
                    lIDField = oDoc.IDOggetto
                    
                    
                End If
            End If
        End If
        
        'Scrive sul registry l'ID da passare all'aplicazione chiamante.
        SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), m_App.Caller, "IDField", lIDField
                                
    End If
    
End Function

'**+
'Nome: FormRecalcLayout
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Ricalcolo del layout del form
'**/
Private Sub FormRecalcLayout()
    Dim Height As Single
    Dim Width As Single

    'Se il form è minimizzato non serve il ricalcolo del layout
    If WindowState <> vbMinimized Then
        ActivityBox.Top = BarMenu.ClientAreaTop
        ActivityBox.Left = BarMenu.ClientAreaLeft
        ActivityBox.Height = IIf(BarMenu.ClientAreaHeight > 0, BarMenu.ClientAreaHeight, 0)
        
        imgSplitter.Visible = ActivityBox.Visible
        imgSplitter.Top = ActivityBox.Top
        imgSplitter.Height = ActivityBox.Height
        
        If ActivityBox.Visible Then
            imgSplitter.Left = ActivityBox.Width + ActivityBox.Left
            picSplitter.Left = imgSplitter.Left
        End If
        
        PicForm.Top = BarMenu.ClientAreaTop
        
        If ActivityBox.Visible Then
            PicForm.Left = imgSplitter.Left + imgSplitter.Width
        Else
            PicForm.Left = BarMenu.ClientAreaLeft
        End If
        


        Width = BarMenu.ClientAreaWidth - IIf(ActivityBox.Visible, ActivityBox.Width + imgSplitter.Width, 0)
        Height = BarMenu.ClientAreaHeight
        
        PicForm.Width = IIf(Width < 100, 100, Width)
        PicForm.Height = IIf(Height < 100, 100, Height)
        
        'RIDIMENSIONA LA SPLIT BAR IN BASE ALLA DIMENSIONE DEL FORM
        DMTSplitBar1.Move PicForm.Left, PicForm.Top, PicForm.Width, PicForm.Height
        'INIZIALIZZA LA SPLIT BAR
        DMTSplitBar1.SetSplitBar Height, Width, Me.PicForm2.Height, Me.PicForm2.Width
        
        'PicForm.Top = BarMenu.ClientAreaTop
        
        'If ActivityBox.Visible Then
        '    PicForm.Left = imgSplitter.Left + imgSplitter.Width
        'Else
        '    PicForm.Left = BarMenu.ClientAreaLeft
        'End If
        
        BrwMain.Top = PicForm.ScaleTop
        BrwMain.Left = PicForm.ScaleLeft
        BrwMain.Width = PicForm.ScaleWidth
        BrwMain.Height = PicForm.ScaleHeight
        

    End If
End Sub

'**+
'Nome: GetStatusToolBar
'
'Parametri:
'Enabled - Stato di abilitazione da controllare
'
'Valori di ritorno:
'
'Funzionalità:
'Calcola lo stato dei bottoni della ToolBar standard
'**/
Private Function GetStatusToolBar(ByVal Enabled As Boolean) As Currency
    Dim Valore As Currency

    Valore = 0
    If BarMenu.Bands("Standard").Tools("New").Enabled = Enabled Then Valore = Valore Or BTN_NEW
    If BarMenu.Bands("Standard").Tools("Save").Enabled = Enabled Then Valore = Valore Or BTN_SAVE
    If BarMenu.Bands("Standard").Tools("Print").Enabled = Enabled Then Valore = Valore Or BTN_PRINT
    If BarMenu.Bands("Standard").Tools("PrePrint").Enabled = Enabled Then Valore = Valore Or BTN_PREVIEW
    If BarMenu.Bands("Standard").Tools("Cut").Enabled = Enabled Then Valore = Valore Or BTN_CUT
    If BarMenu.Bands("Standard").Tools("Copy").Enabled = Enabled Then Valore = Valore Or BTN_COPY
    If BarMenu.Bands("Standard").Tools("Paste").Enabled = Enabled Then Valore = Valore Or BTN_PASTE
    If BarMenu.Bands("Standard").Tools("Delete").Enabled = Enabled Then Valore = Valore Or BTN_DELETE
    If BarMenu.Bands("Standard").Tools("Clear").Enabled = Enabled Then Valore = Valore Or BTN_CLEAR
    If BarMenu.Bands("Standard").Tools("NewSearch").Enabled = Enabled Then Valore = Valore Or BTN_FIND
    If BarMenu.Bands("Standard").Tools("ExecuteSearch").Enabled = Enabled Then Valore = Valore Or BTN_SEARCH
    If BarMenu.Bands("Standard").Tools("ChangeView").Enabled = Enabled Then Valore = Valore Or BTN_VIEWMODE
    If BarMenu.Bands("Standard").Tools("SearchPrevious").Enabled = Enabled Then Valore = Valore Or BTN_PREVIOUS
    If BarMenu.Bands("Standard").Tools("SearchNext").Enabled = Enabled Then Valore = Valore Or BTN_NEXT
    If BarMenu.Bands("Standard").Tools("Export").Enabled = Enabled Then Valore = Valore Or BTN_EXPORT
    If BarMenu.Bands("Band_Export").Tools("ExportWord").Enabled = Enabled Then Valore = Valore Or BTN_WORD
    If BarMenu.Bands("Band_Export").Tools("ExportExcel").Enabled = Enabled Then Valore = Valore Or BTN_EXCEL
    If BarMenu.Bands("Band_Export").Tools("ExportHtml").Enabled = Enabled Then Valore = Valore Or BTN_HTML
    If BarMenu.Bands("Band_Export").Tools("ExportPDF").Enabled = Enabled Then Valore = Valore Or BTN_PDF
    
    If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHFORM
    If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHTABLE
    If BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled = Enabled Then Valore = Valore Or BTN_FILTER

    If BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled = Enabled Then Valore = Valore Or BTN_NEW
    If BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled = Enabled Then Valore = Valore Or BTN_SAVE
    If BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled = Enabled Then Valore = Valore Or BTN_PRINT
    If BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Enabled = Enabled Then Valore = Valore Or BTN_PREVIEW

    If BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled = Enabled Then Valore = Valore Or BTN_CUT
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled = Enabled Then Valore = Valore Or BTN_COPY
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled = Enabled Then Valore = Valore Or BTN_PASTE
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled = Enabled Then Valore = Valore Or BTN_DELETE
    If BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Enabled = Enabled Then Valore = Valore Or BTN_CLEAR
    If BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled = Enabled Then Valore = Valore Or BTN_FIND
    If BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled = Enabled Then Valore = Valore Or BTN_SEARCH
    If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled = Enabled Then Valore = Valore Or BTN_PREVIOUS
    If BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled = Enabled Then Valore = Valore Or BTN_NEXT

    If BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Enabled = Enabled Then Valore = Valore Or BTN_TOOLS
    If BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHFORM
    If BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = Enabled Then Valore = Valore Or BTN_SEARCHTABLE

    If BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Enabled = Enabled Then Valore = Valore Or BTN_EXPORT
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Enabled = Enabled Then Valore = Valore Or BTN_WORD
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Enabled = Enabled Then Valore = Valore Or BTN_EXCEL
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Enabled = Enabled Then Valore = Valore Or BTN_HTML
    If BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Enabled = Enabled Then Valore = Valore Or BTN_PDF

    GetStatusToolBar = Valore
End Function

'**+
'Nome: ReadRegistrySettings
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Legge i valori registrati nel registry relativi allo stato
'dei controlli del Form
'
'
'**/
Private Sub ReadRegistrySettings()
    Dim Index As Integer
    Dim FormHeight As Single
    Dim FormWidth As Single
    Dim NomeBanda As String
    Dim lngIDLanguage As Long
    Dim bFoldersVisible As Boolean
    Dim lValue As Long
           
           
    'Lettura file di help
    App.HelpFile = MenuOptions.ProgramsPath & "\Diamante.chm"
           
    ' Legge dal Registry le impostazioni sulla lingua
    lngIDLanguage = AppOptions.IDLanguage
           
    ' Modifica tutte le stringhe nel linguaggio corrente ( se <> da default )
    If lngIDLanguage <> NATIVE_LANGUAGE Then
        gResource.IDCurrentLanguage = lngIDLanguage
        'Setta i nuovi ToolTipText della Toolbar
        'e le Caption dei menu
        ChangeMenuLanguage
        ChangeToolBarLanguage
        'Traduce tutte le stringhe presenti sul form
        '(Solo se ChangeStringsLanguage è gestita dal programmatore !!!)
        ChangeStringsLanguage
    End If
    
    'Settaggio per la statusbar
    stbStatusbar.Visible = AppOptions.StatusBarVisibility
        
        
    '**+ settaggi per la barra degli strumenti
    With BarMenu
    
        '**+ E' necessario verificare la versione dell'activebar xchè nella nuova vesione 3.0
        'sono stati cambiati i valori di impostazione della proprietà DockingArea
        If AppOptions.BARMENUVERSION = BARMENUVERSION Then
    
            For Index = 0 To .Bands.Count - 1
            
                'Settaggi sulle toolbar (ancoraggio e dimensioni)
                If .Bands(Index).Type <> ddBTPopup Then
                    With .Bands(Index)
                        If AppOptions.ToolbarDockingArea(Index) > -1 Then
                            .DockingArea = AppOptions.ToolbarDockingArea(Index)
                            .DockLine = AppOptions.ToolbarDockLine(Index)
                            lValue = AppOptions.ToolbarHeight(Index)
                            If lValue > 0 Then .Height = lValue
                            lValue = AppOptions.ToolbarWidth(Index)
                            If lValue > 0 Then .Width = lValue
                            '**+ Attenzione le impostazioni del Left e Top devono essere effettuate dopo
                            'quelle dell'Height e del Width xchè se siamo in presenza di valori superiori
                            'a quelli della ClientArea azzera il left e top impostati in precedenza **/
                            lValue = AppOptions.ToolbarLeft(Index)
                            If lValue > 0 Then .Left = lValue
                            lValue = AppOptions.ToolbarTop(Index)
                            If lValue > 0 Then .Top = lValue
                            .DockingOffset = AppOptions.ToolbarDockingOffset(Index)
                        End If
                    End With
                End If
            
                'Settaggi sulla visibilità delle toolbar.
                If .Bands(Index).Type = ddBTNormal And .Bands(Index).Name <> BAND_CLOSE_PREVIEW Then
                     NomeBanda = .Bands(Index).Name
                     .Bands(NomeBanda).Visible = AppOptions.ToolbarVisibility(NomeBanda)
                End If
        
            Next Index
            
        End If
        
        'Settaggio sulla visualizzazione dei tooltip.
        .DisplayToolTips = AppOptions.DisplayTooltip
    End With
        
    
    'Dimensione delle icone della ToolBar
    SetToolBarIcons AppOptions.LargeIcon
    
    BarMenu.RecalcLayout
   
    bFoldersVisible = AppOptions.FoldersVisibility
   
    'Settaggi del riquadro attività
    ActivityBox.Visible = bFoldersVisible
    ActivityBox.Width = AppOptions.FoldersWidth
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked = bFoldersVisible
    m_DefaultActivity = AppOptions.DefaultActivity
    
    '**+ settaggi per la finestra principale del programma
    WindowState = AppOptions.WindowState
    If WindowState = 0 Then
        FormHeight = AppOptions.FormHeight
        If FormHeight <> -1 Then
            Height = FormHeight
        End If
        FormWidth = AppOptions.FormWidth
        If FormWidth <> -1 Then
            Width = FormWidth
        End If
    End If
End Sub

'**+
'Nome: SaveRegistrySettings
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Salva i valori relativi allo stato dei controlli del Form
'nel registry
'
'**/
Private Sub SaveRegistrySettings()
    Dim I As Integer

    '**+ Salva le impostazioni relative alle toolbar
    With AppOptions
        
        For I = 0 To BarMenu.Bands.Count - 1
            If BarMenu.Bands(I).Type <> ddBTPopup Then
                    .ToolbarDockingArea(I) = BarMenu.Bands(I).DockingArea
                    .ToolbarDockLine(I) = BarMenu.Bands(I).DockLine
                    .ToolbarLeft(I) = BarMenu.Bands(I).Left
                    .ToolbarTop(I) = BarMenu.Bands(I).Top
                    .ToolbarHeight(I) = BarMenu.Bands(I).Height
                    .ToolbarWidth(I) = BarMenu.Bands(I).Width
                    .ToolbarDockingOffset(I) = BarMenu.Bands(I).DockingOffset
            End If
        Next I
        .BARMENUVERSION = BARMENUVERSION
        
        'Salva le impostazioni relative alla finestra principale.
        If WindowState <> vbMinimized Then
            .FormHeight = Height
            .FormWidth = Width
            .WindowState = WindowState
        End If
        
        'Salva le impostazioni del riquadro attività
        .FoldersWidth = ActivityBox.Width
        .FoldersVisibility = ActivityBox.Visible
        .DefaultActivity = ActivityBox.CurrentActivityKey
    End With
End Sub

'**+
'Nome: ChangeView
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Cambia modalità di visualizzazione dei dati tra Form e vista tabellare
'
'**/
Private Sub ChangeView(Optional ByVal sToolName As Variant)

    'Se non vi sono record presenti nel browser
    'la modalità di visualizzazione non cambia e si esce.
    If (m_Document.EOF = True And m_Document.BOF = True) Then Exit Sub

    'Se si proviene dalla modalità tabellare
    '( o dalla modalità filtro provenendo dalla modalità tabellare )
    'potrebbe essere necessario allineare il documento con l'ultima selezione fatta nella browse.
    If BrwMain.Visible = True Then
        If BrwMain.ListIndex > 0 Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    

    If IsMissing(sToolName) Then sToolName = "ChangeView"

    'Cambia la visibiltà del browser
    If sToolName = "ChangeView" Then
        BrwMain.Visible = IIf(BrwMain.Visible And BrwMain.GuiMode = dgNormal, False, True)
    Else
        BrwMain.Visible = IIf((sToolName = "Mnu_FormView"), False, True)
    End If
    
    'Se si va in modalità form ed il record è bloccato si torna in modalità tabellare
    'impedendo di effettuare modifiche su quel record.
    'Quando si va in modalità tabellare il controllo non è necessario.
    If Not BrwMain.Visible Then

        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions) Then
            'Il record è bloccato - si va in modalità tabellare
            
            BrwMain.Visible = True

            'Input Focus al browser
            BrwMain.SetFocus

            'Refresh dello stato dei bottoni della ToolBar standard e dei menu
            SetStatus4Modality Browse

            Exit Sub
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
        End If

    End If
    
    
    
    'Se si era in fase di immissione di un nuovo record viene annullata
    m_Document.AbortNew
    
    If BrwMain.Visible Then 'Modalità tabellare
        
        'Input Focus al browser
        BrwMain.SetFocus
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality 3 'Browse
        
    Else 'Modalità form
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality 1 'Modify
        
        'Input Focus al primo campo del form
        If dtData.Enabled = True Then
            dtData.SetFocus
        Else
            Me.cdAnagrafica.SetFocus
        End If
    End If
       
    'Imposta i suggerimenti da visualizzare sulla Statusbar in funzione
    'della modalità di visualizzazione corrente.
    'Ad esempio in alcuni casi le frasi sono al Singolare/Plurare.
    'La funzione GetDescription4StatusBar si occupa di determinare la frase esatta.
    'La Sub RefreshDescriptions4StatusBar deve essere chiamata anche in Execute_Search()--> Vedi.
    RefreshDescriptions4StatusBar
End Sub

'**+
'Nome: InitMenuBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione della MenuBar
'
'**/
Private Sub InitMenuBar(ByRef ToolID As Integer)
    BarMenu.Bands.Add "Band_Menu"
    BarMenu.Bands("Band_Menu").WrapTools = True
    BarMenu.Bands("Band_Menu").Type = ddBTMenuBar
    BarMenu.Bands("Band_Menu").DockLine = 1
    BarMenu.Bands("Band_Menu").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("Band_Menu").GrabHandleStyle = ddGSNormal

    'File
    BarMenu.Bands.Add "Band_File"
    BarMenu.Bands("Band_File").Type = ddBTPopup
    BarMenu.Bands("Band_File").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "File"
    BarMenu.Bands("Band_Menu").Tools("File").SubBand = "Band_File"
    BarMenu.Bands("Band_Menu").Tools("File").Caption = GetCaption4MenuBar("File")
    BarMenu.Bands("Band_Menu").Tools("File").Description = GetDescription4StatusBar("File")

    'File-New
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_New"
    BarMenu.Bands("Band_File").Tools("Mnu_New").SetPicture 0, gResource.GetBitmap(IDB_STD_NEW16), &HC0C0C0
    BarMenu.Bands("Band_File").Tools("Mnu_New").Caption = GetCaption4MenuBar("Mnu_New")
    BarMenu.Bands("Band_File").Tools("Mnu_New").Description = GetDescription4StatusBar("Mnu_New")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_Save"
    BarMenu.Bands("Band_File").Tools("SepMnu_Save").ControlType = ddTTSeparator
    
    'File-Save
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Save"
    BarMenu.Bands("Band_File").Tools("Mnu_Save").SetPicture 0, gResource.GetBitmap(IDB_STD_SAVE16), &HC0C0C0
    If m_App.Language <> 1 Then
        BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    Else
        BarMenu.Bands("Band_File").Tools("Mnu_Save").Caption = GetCaption4MenuBar("Mnu_Save")
    End If
    BarMenu.Bands("Band_File").Tools("Mnu_Save").Description = GetDescription4StatusBar("Mnu_Save")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_PrePrint"
    BarMenu.Bands("Band_File").Tools("SepMnu_PrePrint").ControlType = ddTTSeparator
    
    'File-PrePrint
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_PrePrint"
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIEW16), &HC0C0C0
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Caption = GetCaption4MenuBar("Mnu_PrePrint")
    BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Description = GetDescription4StatusBar("Mnu_PrePrint")
    
    'File-Print
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Print"
    BarMenu.Bands("Band_File").Tools("Mnu_Print").SetPicture 0, gResource.GetBitmap(IDB_STD_PRINT16), &HC0C0C0
    If m_App.Language <> 1 Then
        BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    Else
        BarMenu.Bands("Band_File").Tools("Mnu_Print").Caption = GetCaption4MenuBar("Mnu_Print")
    End If
    BarMenu.Bands("Band_File").Tools("Mnu_Print").Description = GetDescription4StatusBar("Mnu_Print")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "SepMnu_Exit"
    BarMenu.Bands("Band_File").Tools("SepMnu_Exit").ControlType = ddTTSeparator
    
    'File-Exit
    ToolID = ToolID + 1
    BarMenu.Bands("Band_File").Tools.Add ToolID, "Mnu_Exit"
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Caption = GetCaption4MenuBar("Mnu_Exit")
    BarMenu.Bands("Band_File").Tools("Mnu_Exit").Description = GetDescription4StatusBar("Mnu_Exit")
    
    'Edit
    BarMenu.Bands.Add "Band_Edit"
    BarMenu.Bands("Band_Edit").Type = ddBTPopup
    BarMenu.Bands("Band_Edit").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Edit"
    BarMenu.Bands("Band_Menu").Tools("Edit").SubBand = "Band_Edit"
    BarMenu.Bands("Band_Menu").Tools("Edit").Caption = GetCaption4MenuBar("Edit")
    BarMenu.Bands("Band_Menu").Tools("Edit").Description = GetDescription4StatusBar("Edit")
    
    'Edit-Delete
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Delete"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").SetPicture 0, gResource.GetBitmap(IDB_STD_DELETE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Caption = GetCaption4MenuBar("Mnu_Delete")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Description = GetDescription4StatusBar("Mnu_Delete")
    
    'Edit-Clear
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Clear"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").SetPicture 0, gResource.GetBitmap(IDB_STD_CLEAR16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Caption = GetCaption4MenuBar("Mnu_Clear")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Description = GetDescription4StatusBar("Mnu_Clear")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "SepMnu_Cut"
    BarMenu.Bands("Band_Edit").Tools("SepMnu_Cut").ControlType = ddTTSeparator
    
    'Edit-Cut
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Cut"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").SetPicture 0, gResource.GetBitmap(IDB_STD_CUT16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Caption = GetCaption4MenuBar("Mnu_Cut")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Description = GetDescription4StatusBar("Mnu_Cut")
    
    'Edit-Copy
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Copy"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").SetPicture 0, gResource.GetBitmap(IDB_STD_COPY16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Caption = GetCaption4MenuBar("Mnu_Copy")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Description = GetDescription4StatusBar("Mnu_Copy")
    
    'Edit-Paste
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_Paste"
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").SetPicture 0, gResource.GetBitmap(IDB_STD_PASTE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Caption = GetCaption4MenuBar("Mnu_Paste")
    BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Description = GetDescription4StatusBar("Mnu_Paste")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "SepMnu_NewSearch"
    BarMenu.Bands("Band_Edit").Tools("SepMnu_NewSearch").ControlType = ddTTSeparator
    
    'Edit-NewSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_NewSearch"
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_FIND16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Caption = GetCaption4MenuBar("Mnu_NewSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Description = GetDescription4StatusBar("Mnu_NewSearch")
    
    'Edit-ExecuteSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_ExecuteSearch"
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").SetPicture 0, gResource.GetBitmap(IDB_STD_EXECUTE16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Caption = GetCaption4MenuBar("Mnu_ExecuteSearch")
    BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Description = GetDescription4StatusBar("Mnu_ExecuteSearch")
    
    'Edit-SearchPrevious
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_SearchPrevious"
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").SetPicture 0, gResource.GetBitmap(IDB_STD_PREVIOUS16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Caption = GetCaption4MenuBar("Mnu_SearchPrevious")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Description = GetDescription4StatusBar("Mnu_SearchPrevious")
    
    'Edit-SearchNext
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Edit").Tools.Add ToolID, "Mnu_SearchNext"
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").SetPicture 0, gResource.GetBitmap(IDB_STD_NEXT16), &HC0C0C0
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Caption = GetCaption4MenuBar("Mnu_SearchNext")
    BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Description = GetDescription4StatusBar("Mnu_SearchNext")
    
    'View
    BarMenu.Bands.Add "Band_View"
    BarMenu.Bands("Band_View").Type = ddBTPopup
    BarMenu.Bands("Band_View").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "View"
    BarMenu.Bands("Band_Menu").Tools("View").SubBand = "Band_View"
    BarMenu.Bands("Band_Menu").Tools("View").Caption = GetCaption4MenuBar("View")
    BarMenu.Bands("Band_Menu").Tools("View").Description = GetDescription4StatusBar("View")
    
    'View-FormView
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_FormView"
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").SetPicture 0, gResource.GetBitmap(IDB_STD_FORM16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_View").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'View-TableView
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_TableView"
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").SetPicture 0, gResource.GetBitmap(IDB_STD_GRID16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_View").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
    'View - SearchFilter
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_SearchFilter"
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").SetPicture 0, gResource.GetBitmap(IDB_FILTRO16), &HC0C0C0
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "SepMnu_Folders"
    BarMenu.Bands("Band_View").Tools("SepMnu_Folders").ControlType = ddTTSeparator
    
    'View-Folders
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_Folders"
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Caption = GetCaption4MenuBar("Mnu_Folders")
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Description = GetDescription4StatusBar("Mnu_Folders")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "SepMnu_ToolBar"
    BarMenu.Bands("Band_View").Tools("SepMnu_ToolBar").ControlType = ddTTSeparator
    
    'View-ToolBar
    ToolID = ToolID + 1
    BarMenu.Bands("Band_View").Tools.Add ToolID, "Mnu_ToolBar"
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Caption = GetCaption4MenuBar("Mnu_ToolBar")
    BarMenu.Bands("Band_View").Tools("Mnu_ToolBar").Description = GetDescription4StatusBar("Mnu_ToolBar")
    
    'Tools
    BarMenu.Bands.Add "Band_Tools"
    BarMenu.Bands("Band_Tools").Type = ddBTPopup
    BarMenu.Bands("Band_Tools").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Tools"
    BarMenu.Bands("Band_Menu").Tools("Tools").SubBand = "Band_Tools"
    BarMenu.Bands("Band_Menu").Tools("Tools").Caption = GetCaption4MenuBar("Tools")
    BarMenu.Bands("Band_Menu").Tools("Tools").Description = GetDescription4StatusBar("Tools")
    
    'Tools-Export
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_Export"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").ControlType = ddTTLabel
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").SubBand = "Mnu_Band_Export"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Caption = GetCaption4MenuBar("Mnu_Export")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Description = GetDescription4StatusBar("Mnu_Export")
    
    'Tools-Options
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_Options"
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Caption = GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_Options").Description = GetDescription4StatusBar("Mnu_Options")
    
    'Tools-stampa-doc-sel
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Tools").Tools.Add ToolID, "Mnu_stampa_doc_sel"
    BarMenu.Bands("Band_Tools").Tools("Mnu_stampa_doc_sel").Caption = "Stampa documenti con selezione" 'GetCaption4MenuBar("Mnu_Options")
    BarMenu.Bands("Band_Tools").Tools("Mnu_stampa_doc_sel").Description = "Avvia la procedura per selezionare i documenti da stampare" 'GetDescription4StatusBar("Mnu_Options")
    
    'Tools-Export
    BarMenu.Bands.Add "Mnu_Band_Export"
    BarMenu.Bands("Mnu_Band_Export").Type = ddBTPopup
    BarMenu.Bands("Mnu_Band_Export").DockingArea = ddDAPopup
    
    'Tools-Export-ExportPDF
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportPDF"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").SetPicture 0, gResource.GetBitmap(IDB_ACROBAT_16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Description = GetDescription4StatusBar("Mnu_ExportPDF")
    
    'Tools-Export-ExportWord
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportWord"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").SetPicture 0, gResource.GetBitmap(IDB_STD_WORD16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Description = GetDescription4StatusBar("Mnu_ExportWord")
    
    'Tools-Export-ExportExcel
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportExcel"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").SetPicture 0, gResource.GetBitmap(IDB_STD_EXCEL16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Description = GetDescription4StatusBar("Mnu_ExportExcel")
    
    'Tools-Export-ExportHtml
    ToolID = ToolID + 1
    BarMenu.Bands("Mnu_Band_Export").Tools.Add ToolID, "Mnu_ExportHtml"
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").SetPicture 0, gResource.GetBitmap(IDB_STD_HTML16), &HC0C0C0
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Description = GetDescription4StatusBar("Mnu_ExportHtml")


    'Help
    BarMenu.Bands.Add "Band_Help"
    BarMenu.Bands("Band_Help").Type = ddBTPopup
    BarMenu.Bands("Band_Help").DockingArea = ddDATop
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Menu").Tools.Add ToolID, "Help"
    BarMenu.Bands("Band_Menu").Tools("Help").Caption = GetCaption4MenuBar("Help")
    BarMenu.Bands("Band_Menu").Tools("Help").Description = GetDescription4StatusBar("Help")
    BarMenu.Bands("Band_Menu").Tools("Help").SubBand = "Band_Help"
    
    'Help-HelpOnLine
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_HelpOnLine"
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Caption = GetCaption4MenuBar("Mnu_HelpOnLine")
    BarMenu.Bands("Band_Help").Tools("Mnu_HelpOnLine").Description = GetDescription4StatusBar("Mnu_HelpOnLine")
    
    'Help-Arg
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Arg"
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Caption = GetCaption4MenuBar("Mnu_Arg")
    BarMenu.Bands("Band_Help").Tools("Mnu_Arg").Description = GetDescription4StatusBar("Mnu_Arg")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "SepMnu_Web"
    BarMenu.Bands("Band_Help").Tools("SepMnu_Web").ControlType = ddTTSeparator
    
    'Help-Web
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Web"
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").SetPicture 0, gResource.GetBitmap(IDB_DMT_WEB16), &HC0C0C0
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Caption = GetCaption4MenuBar("Mnu_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Web").Description = GetDescription4StatusBar("Mnu_Web")
    
    'Help-Blog
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Agg_Web"
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").SetPicture 0, gResource.GetBitmap(IDB_AGG_WEB16), &HC0C0C0
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Caption = GetCaption4MenuBar("Mnu_Agg_Web")
    BarMenu.Bands("Band_Help").Tools("Mnu_Agg_Web").Description = GetDescription4StatusBar("Mnu_Agg_Web")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "SepMnu_Info"
    BarMenu.Bands("Band_Help").Tools("SepMnu_Info").ControlType = ddTTSeparator
    
    'Help-Info
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Help").Tools.Add ToolID, "Mnu_Info"
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Caption = GetCaption4MenuBar("Mnu_Info")
    BarMenu.Bands("Band_Help").Tools("Mnu_Info").Description = GetDescription4StatusBar("Mnu_Info")
    
    'PopUp
    BarMenu.Bands.Add "Band_PopUp"
    BarMenu.Bands("Band_PopUp").Type = ddBTPopup
    BarMenu.Bands("Band_PopUp").DockingArea = ddDAPopup
    
    'PopUp-RunApplication
    ToolID = ToolID + 1
    BarMenu.Bands("Band_PopUp").Tools.Add ToolID, "Mnu_RunApplication"
    BarMenu.Bands("Band_PopUp").Tools("Mnu_RunApplication").Caption = GetCaption4MenuBar("Mnu_RunApplication")
    
    'PopUp-SearchObject
    ToolID = ToolID + 1
    BarMenu.Bands("Band_PopUp").Tools.Add ToolID, "Mnu_SearchObject"
    BarMenu.Bands("Band_PopUp").Tools("Mnu_SearchObject").Caption = GetCaption4MenuBar("Mnu_SearchObject")
    
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: InitToolBar
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Inizializzazione della ToolBar
'
'**/
Private Sub InitToolBar(ByRef ToolID As Integer)


    BarMenu.Bands.Add "Standard"
    BarMenu.Bands("Standard").DockLine = 2
    BarMenu.Bands("Standard").Type = ddBTNormal
    BarMenu.Bands("Standard").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("Standard").GrabHandleStyle = ddGSNormal
    BarMenu.Bands.Add BAND_CLOSE_PREVIEW
    BarMenu.Bands(BAND_CLOSE_PREVIEW).DockLine = 2
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Type = ddBTMenuBar
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Caption = "Chiudi"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).DockingArea = ddDATop
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Visible = False

    'New
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "New"
    BarMenu.Bands("Standard").Tools("New").ToolTipText = GetToolTipText4ToolBar("New")
    BarMenu.Bands("Standard").Tools("New").Description = GetDescription4StatusBar("New")
    
    'Save
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Save"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep2"
    BarMenu.Bands("Standard").Tools("Sep2").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("Save").ToolTipText = GetToolTipText4ToolBar("Save")
    BarMenu.Bands("Standard").Tools("Save").Description = GetDescription4StatusBar("Save")
    
    'Print
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Print"
    BarMenu.Bands("Standard").Tools("Print").ToolTipText = GetToolTipText4ToolBar("Print")
    BarMenu.Bands("Standard").Tools("Print").Description = GetDescription4StatusBar("Print")
    
    'PrePrint
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "PrePrint"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep3"
    BarMenu.Bands("Standard").Tools("Sep3").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("PrePrint").ToolTipText = GetToolTipText4ToolBar("PrePrint")
    BarMenu.Bands("Standard").Tools("PrePrint").Description = GetDescription4StatusBar("PrePrint")
    
    'Cut
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Cut"
    BarMenu.Bands("Standard").Tools("Cut").ToolTipText = GetToolTipText4ToolBar("Cut")
    BarMenu.Bands("Standard").Tools("Cut").Description = GetDescription4StatusBar("Cut")
    
    'Copy
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Copy"
    BarMenu.Bands("Standard").Tools("Copy").ToolTipText = GetToolTipText4ToolBar("Copy")
    BarMenu.Bands("Standard").Tools("Copy").Description = GetDescription4StatusBar("Copy")
    
    'Paste
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Paste"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep"
    BarMenu.Bands("Standard").Tools("Sep").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("Paste").ToolTipText = GetToolTipText4ToolBar("Paste")
    BarMenu.Bands("Standard").Tools("Paste").Description = GetDescription4StatusBar("Paste")
    
    'Delete
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Delete"
    BarMenu.Bands("Standard").Tools("Delete").ToolTipText = GetToolTipText4ToolBar("Delete")
    BarMenu.Bands("Standard").Tools("Delete").Description = GetDescription4StatusBar("Delete")
    
    'Clear
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Clear"
    BarMenu.Bands("Standard").Tools("Clear").ToolTipText = GetToolTipText4ToolBar("Clear")
    BarMenu.Bands("Standard").Tools("Clear").Description = GetDescription4StatusBar("Clear")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SepNewSearch"
    BarMenu.Bands("Standard").Tools("SepNewSearch").ControlType = ddTTSeparator
    
    'NewSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "NewSearch"
    BarMenu.Bands("Standard").Tools("NewSearch").ToolTipText = GetToolTipText4ToolBar("NewSearch")
    BarMenu.Bands("Standard").Tools("NewSearch").Description = GetDescription4StatusBar("NewSearch")
    
    'ExecuteSearch
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "ExecuteSearch"
    BarMenu.Bands("Standard").Tools("ExecuteSearch").ToolTipText = GetToolTipText4ToolBar("ExecuteSearch")
    BarMenu.Bands("Standard").Tools("ExecuteSearch").Description = GetDescription4StatusBar("ExecuteSearch")
    
    'ChangeView
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "ChangeView"
    BarMenu.Bands("Standard").Tools("ChangeView").ControlType = ddTTButtonDropDown
    BarMenu.Bands("Standard").Tools("ChangeView").SubBand = "Band_ChangeView"
    BarMenu.Bands("Standard").Tools("ChangeView").ToolTipText = GetToolTipText4ToolBar("ChangeView")
    BarMenu.Bands("Standard").Tools("ChangeView").Description = GetDescription4StatusBar("ChangeView")
    BarMenu.Bands.Add "Band_ChangeView"
    BarMenu.Bands("Band_ChangeView").Type = ddBTPopup
    BarMenu.Bands("Band_ChangeView").DockingArea = ddDATop
    
    'ChangeView - Form
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_FormView"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").SetPicture 0, gResource.GetBitmap(IDB_STD_FORM16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Caption = GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Description = GetDescription4StatusBar("Mnu_FormView")
    
    'ChangeView - Tabella
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_TableView"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").SetPicture 0, gResource.GetBitmap(IDB_STD_GRID16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Caption = GetCaption4MenuBar("Mnu_TableView")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Description = GetDescription4StatusBar("Mnu_TableView")
    
     'ChangeView - Filtro
    ToolID = ToolID + 1
    BarMenu.Bands("Band_ChangeView").Tools.Add ToolID, "Mnu_SearchFilter"
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").SetPicture 0, gResource.GetBitmap(IDB_FILTRO16), &HC0C0C0
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Caption = GetCaption4MenuBar("Mnu_SearchFilter")
    BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Description = GetDescription4StatusBar("Mnu_SearchFilter")
    
    'SearchPrevious
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SearchPrevious"
    BarMenu.Bands("Standard").Tools("SearchPrevious").ToolTipText = GetToolTipText4ToolBar("SearchPrevious")
    BarMenu.Bands("Standard").Tools("SearchPrevious").Description = GetDescription4StatusBar("SearchPrevious")
    
    'SearchNext
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "SearchNext"
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep4"
    BarMenu.Bands("Standard").Tools("Sep4").ControlType = ddTTSeparator
    BarMenu.Bands("Standard").Tools("SearchNext").ToolTipText = GetToolTipText4ToolBar("SearchNext")
    BarMenu.Bands("Standard").Tools("SearchNext").Description = GetDescription4StatusBar("SearchNext")
        
    'Export
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Export"
    BarMenu.Bands("Standard").Tools("Export").ControlType = ddTTButtonDropDown
    BarMenu.Bands("Standard").Tools("Export").SubBand = "Band_Export"
    BarMenu.Bands("Standard").Tools("Export").ToolTipText = GetToolTipText4ToolBar("Export")
    BarMenu.Bands("Standard").Tools("Export").Description = GetDescription4StatusBar("Mnu_Export")
    BarMenu.Bands.Add "Band_Export"
    BarMenu.Bands("Band_Export").Type = ddBTPopup
    BarMenu.Bands("Band_Export").DockingArea = ddDATop
    
    'ExportPDF
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportPDF"
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Caption = GetCaption4MenuBar("Mnu_ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").ToolTipText = GetToolTipText4ToolBar("ExportPDF")
    BarMenu.Bands("Band_Export").Tools("ExportPDF").Description = GetDescription4StatusBar("ExportPDF")
    
    'ExportWord
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportWord"
    BarMenu.Bands("Band_Export").Tools("ExportWord").Caption = GetCaption4MenuBar("Mnu_ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").ToolTipText = GetToolTipText4ToolBar("ExportWord")
    BarMenu.Bands("Band_Export").Tools("ExportWord").Description = GetDescription4StatusBar("ExportWord")
    
    'ExportExcel
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportExcel"
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Caption = GetCaption4MenuBar("Mnu_ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").ToolTipText = GetToolTipText4ToolBar("ExportExcel")
    BarMenu.Bands("Band_Export").Tools("ExportExcel").Description = GetDescription4StatusBar("ExportExcel")
    
    'ExportHtml
    ToolID = ToolID + 1
    BarMenu.Bands("Band_Export").Tools.Add ToolID, "ExportHtml"
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Caption = GetCaption4MenuBar("Mnu_ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").ToolTipText = GetToolTipText4ToolBar("ExportHtml")
    BarMenu.Bands("Band_Export").Tools("ExportHtml").Description = GetDescription4StatusBar("ExportHtml")
    
    'Bottone chiusura anteprima
    ToolID = ToolID + 1
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools.Add ToolID, "ClosePreview"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Style = ddSText
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Caption = "&Chiudi"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").ToolTipText = "Chiudi anteprima"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Description = "Esci da modalità Anteprima di stampa"
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").ControlType = ddTTButton
    BarMenu.Bands(BAND_CLOSE_PREVIEW).Tools("ClosePreview").Visible = True
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep5"
    BarMenu.Bands("Standard").Tools("Sep5").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Web"
    BarMenu.Bands("Standard").Tools("Web").ToolTipText = GetToolTipText4ToolBar("Web")
    BarMenu.Bands("Standard").Tools("Web").Description = GetDescription4StatusBar("Web")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Agg_Web"
    BarMenu.Bands("Standard").Tools("Agg_Web").ToolTipText = GetToolTipText4ToolBar("Agg_Web")
    BarMenu.Bands("Standard").Tools("Agg_Web").Description = GetDescription4StatusBar("Agg_Web")
    
        
    BarMenu.Bands.Add "StandardPO"
    BarMenu.Bands("StandardPO").DockLine = 3
    BarMenu.Bands("StandardPO").Type = ddBTNormal
    BarMenu.Bands("StandardPO").Flags = ddBFDockTop Or ddBFDockLeft Or ddBFFloat Or ddBFDockRight Or ddBFDockBottom
    BarMenu.Bands("StandardPO").GrabHandleStyle = ddGSNormal

    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdSalvaComeNuovo"
    BarMenu.Bands("StandardPO").Tools("cmdSalvaComeNuovo").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdSalvaComeNuovo").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdSalvaComeNuovo").ToolTipText = "Salva come nuovo" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdSalvaComeNuovo").Description = "Salva come nuovo"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdSalvaComeNuovo").Caption = "Salva come nuovo"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep11"
    BarMenu.Bands("StandardPO").Tools("Sep11").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdListaPrelievo"
    BarMenu.Bands("StandardPO").Tools("cmdListaPrelievo").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdListaPrelievo").SetPicture 0, gResource.GetBitmap(IDB_ARTICOLI_DA_ASSEGNARE16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdListaPrelievo").ToolTipText = "Lista prelievo" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdListaPrelievo").Description = "Lista prelievo"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdListaPrelievo").Caption = "Lista prelievo"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep12"
    BarMenu.Bands("StandardPO").Tools("Sep12").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdLavorazioni"
    BarMenu.Bands("StandardPO").Tools("cmdLavorazioni").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdLavorazioni").SetPicture 0, gResource.GetBitmap(IDB_COLLO_QTA_PRELEVATA_REFRESH16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdLavorazioni").ToolTipText = "Controllo lavorazioni" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdLavorazioni").Description = "Controllo lavorazioni"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdLavorazioni").Caption = "Controllo lavorazioni"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep13"
    BarMenu.Bands("StandardPO").Tools("Sep13").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdAnalizzaOrdine"
    BarMenu.Bands("StandardPO").Tools("cmdAnalizzaOrdine").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdAnalizzaOrdine").SetPicture 0, gResource.GetBitmap(IDB_FORMULE16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdAnalizzaOrdine").ToolTipText = "Analizza ordine" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdAnalizzaOrdine").Description = "Analizza ordine"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdAnalizzaOrdine").Caption = "Analizza ordine"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep14"
    BarMenu.Bands("StandardPO").Tools("Sep14").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdTour"
    BarMenu.Bands("StandardPO").Tools("cmdTour").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdTour").SetPicture 0, gResource.GetBitmap(IDB_COLLCOMPLETI16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdTour").ToolTipText = "Tour" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdTour").Description = "Tour"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdTour").Caption = "Collegamento Tour"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep15"
    BarMenu.Bands("StandardPO").Tools("Sep15").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdConferimento"
    BarMenu.Bands("StandardPO").Tools("cmdConferimento").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdConferimento").SetPicture 0, gResource.GetBitmap(IDB_IMPCENTRO16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdConferimento").ToolTipText = "Conferimento" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdConferimento").Description = "Crea un conferimento in base alle righe di ordine"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdConferimento").Caption = "Conferimento"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep16"
    BarMenu.Bands("StandardPO").Tools("Sep16").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "cmdContratto"
    BarMenu.Bands("StandardPO").Tools("cmdContratto").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("cmdContratto").SetPicture 0, gResource.GetBitmap(IDB_ACT_ORDER_CUSTOMER_EVASION_16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("cmdContratto").ToolTipText = "Contratto" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdContratto").Description = "Preleva informazioni di testa e di righe presenti in un contratto"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("cmdContratto").Caption = "Dati da contratto"  'GetDescription4StatusBar("Mnu_FormView")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep17"
    BarMenu.Bands("Standard").Tools("Sep17").ControlType = ddTTSeparator
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "cmdAvviaQualita"
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").Style = ddSIconText
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").SetPicture 0, gResource.GetBitmap(IDB_ACT_ORDER_CUSTOMER_EVASION_16), &HC0C0C0
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").ToolTipText = "Gestione qualità" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").Description = "Gestione qualità"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdAvviaQualita").Caption = "Qualità"  'GetDescription4StatusBar("Mnu_FormView")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep18"
    BarMenu.Bands("Standard").Tools("Sep18").ControlType = ddTTSeparator
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "cmdAvviaListaQualita"
    BarMenu.Bands("Standard").Tools("cmdAvviaListaQualita").Style = ddSIconText
    BarMenu.Bands("Standard").Tools("cmdAvviaListaQualita").SetPicture 0, gResource.GetBitmap(IDB_ACT_ORDER_CUSTOMER_EVASION_16), &HC0C0C0
    BarMenu.Bands("Standard").Tools("cmdAvviaListaQualita").ToolTipText = "Elenco qualità" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdAvviaListaQualita").Description = "Elenco qualità"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("cmdAvviaListaQualita").Caption = "Elenco qualità"  'GetDescription4StatusBar("Mnu_FormView")
    
    
    BarMenu.RecalcLayout
End Sub





'**+
'Nome: ChooseAboutSaving
'
'Parametri:
'Ritorna i valori vbYes, vbNo o vbCancel a seconda della risposta data
'
'Valori di ritorno:
'
'Funzionalità:
'Richiesta della registrazione di un record
'**/
Private Function ChooseAboutSaving() As Integer
    If m_Changed Then
        gResource.CustomStrings.Clear
        gResource.CustomStrings.Add Chr(34) & "Documento del " & dtData.Text & " n° " & lngNumero.Value & Chr(34), 1

        ChooseAboutSaving = fnMsgQuestionWithCancel((gResource.GetCustomizedMessage(MESS_QUERYSAVE)), TheApp.FunctionName)
    End If
End Function


'**+
'Nome: ChooseAboutSavingOkCancel
'
'Parametri:
'
'Valori di ritorno:
'Ritorna i valori vbOK o vbCancel a seconda della risposta data
'
'Funzionalità:
'Come ChooseAboutSaving ma con pulsanti Ok e Annulla
'**/
Private Function ChooseAboutSavingOkCancel() As Integer
    Dim sRecord As String

    sRecord = Caption2Display(False)
  
    gResource.CustomStrings.Clear
    gResource.CustomStrings.Add Chr(34) & sRecord & Chr(34), 1
    ChooseAboutSavingOkCancel = fnMsgQuestionOKCancel((gResource.GetCustomizedMessage(MESS_QUERYSAVE)), m_App.FunctionName)
    
End Function

'**+
'Nome: ActivateBarButtons
'
'Parametri:
'Buttons - Variabile lunga 8 byte con la combinazione
'della maschera di bit che indica di quali bottoni cambiare
'lo stato di abilitazione
'Enable - Valore booleano che indica lo stato di abilitazione
'da applicare
'
'Valori di ritorno:
'
'Funzionalità:
'Abilita o meno gruppi di bottoni e voci di menu
'**/
Private Sub ActivateBarButtons(ByVal Buttons As Currency, ByVal Enable As Boolean)

    'Pulsanti della Toolbar
    '----------------------
    If (Buttons And BTN_NEW) Then BarMenu.Bands("Standard").Tools("New").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_SAVE) Then BarMenu.Bands("Standard").Tools("Save").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PRINT) Then BarMenu.Bands("Standard").Tools("Print").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PREVIEW) Then BarMenu.Bands("Standard").Tools("PrePrint").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_CUT) Then BarMenu.Bands("Standard").Tools("Cut").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_COPY) Then BarMenu.Bands("Standard").Tools("Copy").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PASTE) Then BarMenu.Bands("Standard").Tools("Paste").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_DELETE) Then BarMenu.Bands("Standard").Tools("Delete").Enabled = CheckRights("Cancellazione", Enable)
    If (Buttons And BTN_CLEAR) Then BarMenu.Bands("Standard").Tools("Clear").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_FIND) Then BarMenu.Bands("Standard").Tools("NewSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCH) Then BarMenu.Bands("Standard").Tools("ExecuteSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_VIEWMODE) Then BarMenu.Bands("Standard").Tools("ChangeView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_PREVIOUS) Then BarMenu.Bands("Standard").Tools("SearchPrevious").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_NEXT) Then BarMenu.Bands("Standard").Tools("SearchNext").Enabled = CheckRights("Selezione", Enable)
    If Not oExportActivity Is Nothing Then
        If (Buttons And BTN_EXPORT) Then oExportActivity.EnableItems CheckRights("Stampa", Enable)
    End If
    If (Buttons And BTN_EXPORT) Then BarMenu.Bands("Standard").Tools("Export").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_WORD) Then BarMenu.Bands("Band_Export").Tools("ExportWord").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_EXCEL) Then BarMenu.Bands("Band_Export").Tools("ExportExcel").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_HTML) Then BarMenu.Bands("Band_Export").Tools("ExportHtml").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PDF) Then BarMenu.Bands("Band_Export").Tools("ExportPDF").Enabled = CheckRights("Stampa", Enable)
    
    If (Buttons And BTN_SEARCHFORM) Then
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_FormView").Checked = Not Enable
    End If
    
    If (Buttons And BTN_SEARCHTABLE) Then
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Enabled = CheckRights("Selezione", Enable)
        BarMenu.Bands("Band_ChangeView").Tools("Mnu_TableView").Checked = Not Enable
    End If
    
    If (Buttons And BTN_FILTER) Then BarMenu.Bands("Band_ChangeView").Tools("Mnu_SearchFilter").Enabled = CheckRights("Selezione", Enable)
    
    'VOCI DI MENU
    '------------
    
    'Menu File
    '---------
    If (Buttons And BTN_NEW) Then BarMenu.Bands("Band_File").Tools("Mnu_New").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_SAVE) Then BarMenu.Bands("Band_File").Tools("Mnu_Save").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PRINT) Then BarMenu.Bands("Band_File").Tools("Mnu_Print").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PREVIEW) Then BarMenu.Bands("Band_File").Tools("Mnu_PrePrint").Enabled = CheckRights("Stampa", Enable)
    
    'Menu Edit
    '---------
    If (Buttons And BTN_CUT) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Cut").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_COPY) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Copy").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_PASTE) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Paste").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_DELETE) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Delete").Enabled = CheckRights("Cancellazione", Enable)
    If (Buttons And BTN_CLEAR) Then BarMenu.Bands("Band_Edit").Tools("Mnu_Clear").Enabled = CheckRights("Modifica", Enable)
    If (Buttons And BTN_FIND) Then BarMenu.Bands("Band_Edit").Tools("Mnu_NewSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCH) Then BarMenu.Bands("Band_Edit").Tools("Mnu_ExecuteSearch").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_PREVIOUS) Then BarMenu.Bands("Band_Edit").Tools("Mnu_SearchPrevious").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_NEXT) Then BarMenu.Bands("Band_Edit").Tools("Mnu_SearchNext").Enabled = CheckRights("Selezione", Enable)
    
    'Menu Visualizza
    '---------------
    If (Buttons And BTN_FILTER) Then BarMenu.Bands("Band_View").Tools("Mnu_SearchFilter").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCHFORM) Then BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_SEARCHTABLE) Then BarMenu.Bands("Band_View").Tools("Mnu_TableView").Enabled = CheckRights("Selezione", Enable)
    If (Buttons And BTN_VIEWMODE) Then
        BarMenu.Bands("Band_View").Tools("Mnu_FormView").Enabled = CheckRights("Selezione", Enable)
    End If

    'Menu Export
    '-----------
    If Not oExportActivity Is Nothing Then
        If (Buttons And BTN_EXPORT) Then oExportActivity.EnableItems CheckRights("Stampa", Enable)
    End If
    If (Buttons And BTN_EXPORT) Then BarMenu.Bands("Band_Tools").Tools("Mnu_Export").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_WORD) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportWord").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_EXCEL) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportExcel").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_HTML) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportHtml").Enabled = CheckRights("Stampa", Enable)
    If (Buttons And BTN_PDF) Then BarMenu.Bands("Mnu_Band_Export").Tools("Mnu_ExportPDF").Enabled = CheckRights("Stampa", Enable)
End Sub

'**+
'Nome: NewSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni da compiere in caso di richiesta di una nuova ricerca
'**/
Private Sub NewSearch()

    'Refresh dello stato del Form
    m_Changed = False
    m_Saved = False
    m_Search = True
    
    'Annulla una eventuale operazione di inserimento di un nuovo record
    If m_Document.TableNew Then
        m_Document.AbortNew
        RefreshFormFields
    End If
    
    'Ripristina la vista del Form
    BrwMain.Visible = True
    
    'Predispone la modalità DefineFilter della Browse
    BrwMain.AbortFilterEdit = False
    BrwMain.GuiMode = dgFilterDefinition
    BrwMain.SetFocus
    
    'Refresh dello stato dei bottoni delle barre dei menu per la modalità ricerca
    SetStatus4Modality 2 'Find
    
End Sub

'**+
'Nome: ExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue la ricerca impostata.
'
'**/
Private Sub ExecuteSearch()
    Dim Cond As DmtGridCtl.dgCondition
    Dim Field As DmtDocManLib.Field
    Dim OLDCursor As Integer
    Dim sWhere As String
    
    
    'Gestione della clessidra
    OLDCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    
    
    'Se non è stato selezionato nessun filtro dal controllo DocTypeExplorer
    'viene creato un filtro temporaneo in memoria e reso il filtro attivo
    If Not m_FilterSelected Then
        
        'Comunica all'oggetto DocType i valori da usare per la ricerca
        sWhere = fnFillDocTypeCondition
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"
        
        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")
                           
        'Aggiunge al filtro eventuali condizioni aggiuntive restituite dalla funzione fnFillDocTypeCondition
        If sWhere <> "" Then m_ActiveFilter.AddCondition sWhere
        
    End If
    
    
    'Comunica al documento il nuovo filtro da usare
    Set m_Document.ActiveFilter = m_ActiveFilter
    
    'Viene effettuata la ricerca
    m_Document.OpenDoc
    
    
    'Assegnazione del riferimento alla fonte dati (binding sul recordset del documento)
    Set BrwMain.Recordset = m_Document.Dataset.Recordset
    
    'Ripristina il cursore
    Screen.MousePointer = OLDCursor
    
    'Operazioni da effettuare dopo l'esecuzione della ricerca.
    AfterExecuteSearch
End Sub

'**+
'Nome: AfterExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Determina quali operazioni compiere dopo ExecuteSearch
'              in funzione dell'esito della ricerca.
'
'**/
Private Sub AfterExecuteSearch()

    If Not (m_Document.EOF = True And m_Document.BOF = True) Then
        'La ricerca ha avuto esito positivo
        'Attiva la vista tabellare
        BrwMain.Visible = True
        BrwMain.SetFocus

        'Imposta i menu e la toolbar per la modalità tabellare
        SetStatus4Modality Browse

        'Attiva le procedure di creazione di un nuovo filtro solo se l'ExecuteSearch
        'non è stata chiamata da una selezione del DocTypeExplorer

        'Se l'ExecuteSearch non è stata chiamata da un filtro del riquadro attività
        'si permette di salvare il nuovo filtro ed aggiungerlo nel ramo dei filtri.
        If Not m_FilterSelected Then
            oFiltersActivity.NewFilterBegin
        End If

        'Imposta i suggerimenti da visualizzare sulla Statusbar in funzione
        'della modalità di visualizzazione corrente.
        'Ad esempio in alcuni casi le frasi sono al Singolare/Plurare.
        'Le impostazioni sottostanti servono soltanto all'avvio del programma dopo la prima
        'ricerca. (in quanto ChangeView non è stata ancora eseguita)
        'La Sub RefreshDescriptions4StatusBar deve essere chiamata anche in ChangeView()--> Vedi.
        RefreshDescriptions4StatusBar

        m_Search = False
    Else
        'La ricerca ha avuto esito negativo. Viene mostrato un messaggio
        'e si torna in modalità ricerca.

        'Per questioni estetiche viene subito mostrata la modalità FilterDefinition
        'al posto della browse vuota e quindi viene mostrato il messaggio.
        BrwMain.GuiMode = dgFilterDefinition

        'Se si è selezionato il filtro "Nessun record" non occorre
        'visualizzare il messaggio
        If m_ActiveFilter.NothingSelected = False Or m_FilterSelected = False Then
            'Messaggio  "Nessun elemento trovato"
            sbMsgInfo gResource.GetMessage(MESS_NORECFOUND), m_App.FunctionName
        End If

        'Si torna in modalità form (modalità ricerca)
        OnNewSearch
    End If
    
End Sub



'**+
'Autore: Diamante s.p.a
'Data creazione: 26/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: fnFillDocTypeCondition
'
'Parametri:
'
'Valori di ritorno: String - in base alle esigenze specifiche di una manutenzione è possibile montare ad hoc
'                                 una clausola WHERE che potrà poi essere presa in considerazione nel filtro di selezione
'                                 con il metodo AddCondition dell'oggetto DmtDocManLib.Filter
'
'Funzionalità: Comunica all'oggetto DocType i valori da usare per la ricerca
'
'**/
Private Function fnFillDocTypeCondition() As String
    Dim Field As DmtDocManLib.Field
    Dim Cond As DmtGridCtl.dgCondition
    Dim sWhere As String
    
    
    'NOTA per l'uso dei campi RANGE
    '--------------------------------------------------------------------------------------------------
    'E' consentito l'inserimento, nella modalità filtri e nel caso di campi di tipo range, del solo il valore iniziale
    '(in questo caso vengono filtrati tutti gli elementi maggiori o uguali a quello inserito)
    'o solo quello finale (in questo vengono filtrati tutti gli elementi minori o uguali a quello inserito).
    'Questo funzionamento vale per tutte le tipologie di campo.
    
    'Nel caso di condizione RANGE la sintassi da usare è del tipo della riga sotto:
    'm_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
    '--------------------------------------------------------------------------------------------------
    
    sWhere = vbNullString
    
    'Ripulisce la collezione Fields dell'oggetto DocType
    For Each Field In m_DocType.Fields
        Field.Value = Empty
    Next
    
    m_DocType.Fields("IDAzienda").Value = m_App.IDFirm
    m_DocType.Fields("IDFiliale").Value = m_App.Branch
    
    Me.BrwMain.Refresh
    'Comunica all'oggetto DocType i valori da usare per la ricerca
    For Each Cond In BrwMain.Conditions
        If Cond.IsHeader = False Then
              Select Case Cond.ConditionType
                  
                  'Condizione boolean
                  Case dgCondTypeBoolean
                      m_DocType.Fields(Cond.FieldName).Value = IIf(IsEmpty(Cond.FromValue), Empty, Abs(CDbl(Cond.FromValue = "SI")))
                      
                  'Condizione associata ad una combo box
                  Case dgCondTypeComboDB
                      m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValueID
                  
                  'Condizione di tipo text, numeric, data, time
                  Case dgCondTypeText
                      If Cond.RangeChecked = True Then
                          m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                  Case dgCondTypeNumber
                      If Cond.RangeChecked = True Then
                          m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                  
                  Case dgCondTypeDate
                      If Cond.RangeChecked = True Then
                          m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                          
                          'sWhere = Cond.FieldName & ">=" & fnNormDate(Cond.FromValue)
                          'sWhere = sWhere & " AND " & Cond.FieldName & "<=" & fnNormDate(Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                  
                  Case dgCondTypeTime
                      If Cond.RangeChecked = True Then
                          sWhere = Cond.FieldName & ">=" & fnNormTime(Cond.FromValue)
                          sWhere = sWhere & " AND " & Cond.FieldName & "<=" & fnNormTime(Cond.ToValue)
                      Else
                          m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      End If
                
                  'Altre condizioni
                  Case Else
                      m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                      
        
              End Select
        End If
    Next Cond
        
    fnFillDocTypeCondition = sWhere
    
End Function



'**+
'Nome: CheckRights
'
'Parametri:
'ActionName - Nome della azione
'Enable - Valore da modificare o ritornare inalterato
'
'Valori di ritorno:
'Il valore in Enable o False se l'azione non è abilitata
'per il tipo di documento
'
'Funzionalità:
'Controlla se l'azione passata è abilitata per il tipo documento
'**/
Private Function CheckRights(ByVal ActionName As String, ByVal Enable As Boolean) As Boolean
    Dim Action As DmtDocManLib.Action
    Dim Dummy As String
    
    If m_DocType.Actions.Count = 0 Then
        CheckRights = Enable
        Exit Function
    End If
    For Each Action In m_DocType.Actions
        If Action.Name = "TUTTE LE AZIONI" Then
            CheckRights = Enable
            Exit Function
        End If
    Next
    On Error GoTo ActionNotFound
    Dummy = m_DocType.Actions(ActionName).Name
    CheckRights = Enable
    Exit Function
ActionNotFound:
    CheckRights = False
End Function

'**+
'Nome: IsFieldInput
'
'Parametri:
'Control - un oggetto Control da controllare
'
'Valori di ritorno:
'Se il controllo è abilitato all'input torna vero altrimenti falso
'
'Funzionalità:
'Controllo se un certo controllo è usabile come campo
'di input dei dati del Form
'**/
Private Function IsFieldInput(ByVal Control As Control) As Boolean
    'Controlla se il Controllo è di Immissione
    IsFieldInput = IsFieldInput Or TypeName(Control) = "TextBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "CheckBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "ComboBox"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "OptionButton"
End Function

'**+
'Nome: OnStart
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Altre inizializzazioni dopo quelle predefinite
'**/
Private Sub OnStart()
    Dim NewLink As DmtDocManLib.Link
    
'**************************SOTTO DOCUMENTO DELLE COMMISSIONI*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    'Set m_DocumentsLink = m_Document.AddDocumentsLink("RV_POCommissioniPerDoc")
    
    'Impostazioni dell'oggetto DocumentsLink
    'm_DocumentsLink.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    'm_DocumentsLink.PrimaryKey = "IDRV_POCommissioniPerDoc" '<-- Specifica il campo chiave primaria
   
    'Crea un Link LEFT JOIN sul campo "Articolo.Articolo"
    'Set NewLink = m_DocumentsLink.AddLink("IDRV_POTipoCommissione", "RV_POTipoCommissione", ltLeft, "IDRV_POTipoCommissione")
    'NewLink.AddLinkColumn "RV_POTipoCommissione.TipoCommissione"

End Sub

'**+
'Nome: OnSave
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Save
'**/
Private Function OnSave() As Boolean
On Error GoTo ERR_OnSave
Dim Field As DmtDocManLib.Field
Dim DocLink As DmtDocManLib.DocumentsLink
Dim sSQL As String
Dim NumeroInserimenti As Long
Dim OLD_CURSOR As Long
Dim IDTour As Long
Dim NuovoOggetto As Boolean


    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Exit Function
    End If
    
    
    OnSave = True
    'Controlli preliminari sulla validità e consistenza dei dati da salvare
    If Not PermissionToSave Then
        OnSave = False
        Exit Function
    End If
        
    'Se la proprietò IDOggetto dell'oggetto cDocument è uguale a 0 (zero)
    'vuol dire che si tratta di un nuovo documento e quindi procediamo con il suo
    'inserimento altrimenti effettuiamo un aggiornamento del documento esistente
    bSaving = 1

    If Me.cboCambioValuta.CurrentID = 0 Then
        oDoc.Field "Link_Val_cambio", Null, sTabellaTestata
        oDoc.Field "Val_valore_cambio", Null, sTabellaTestata
        oDoc.Field "Val_data_cambio", Null, sTabellaTestata
    Else
        oDoc.Field "Link_Val_cambio", Me.cboCambioValuta.CurrentID, sTabellaTestata
        oDoc.Field "Val_valore_cambio", Me.txtValoreCambioValuta.Value, sTabellaTestata
        oDoc.Field "Val_data_cambio", Me.txtDataCambio.Text, sTabellaTestata
    End If
    
    
    OLD_CURSOR = Cn.CursorLocation
    Cn.CursorLocation = adUseClient

    If oDoc.IDOggetto = 0 Then
        If GET_ESISTENZA_NUMERO_DOCUMENTO(oDoc.IDTipoOggetto, oDoc.IDOggetto) = True Then
            'MsgBox "Numero documento esistente", vbInformation, "Salvataggio documento"
            Me.lngNumero.Value = fnDocumentNumber(Me.dtData.Text)
            If (Me.txtNListaPrelievo.Value = 1) Then
                txtNOrdinePadre.Value = Me.lngNumero.Value
            End If
            MsgBox "Al documento verrà associato il numero " & Me.lngNumero.Value, vbInformation, "Salva documento"
        End If
        'If Me.chkProtocolloICE.Value = vbChecked Then
        '    AggiornaProtocolloICE 0
        'End If
        Me.Caption = "SALVATAGGIO IN CORSO....................."
        If (Me.txtNListaPrelievo.Value = 1) Then
            txtNOrdinePadre.Value = Me.lngNumero.Value
            Me.txtDataDocPadre.Value = Me.dtData.Value
            
            oDoc.Field "RV_PONumeroOrdinePadre", Me.txtNOrdinePadre.Value, sTabellaTestata
            oDoc.Field "RV_PODataOrdinePadre", Me.txtDataDocPadre.Value, sTabellaTestata
        End If
        
        'Crea il nuovo documento
        oDoc.Insert
        
        NuovoOggetto = True
    Else
        If GET_ESISTENZA_NUMERO_DOCUMENTO(oDoc.IDTipoOggetto, oDoc.IDOggetto) = True Then
            MsgBox "Numero documento esistente", vbInformation, "Salvataggio documento"
            Me.lngNumero.SetFocus
            Exit Function
        End If
        If (Me.txtNListaPrelievo.Value = 1) Then
            txtNOrdinePadre.Value = Me.lngNumero.Value
            Me.txtDataDocPadre.Value = Me.dtData.Value
            
            oDoc.Field "RV_PONumeroOrdinePadre", Me.txtNOrdinePadre.Value, sTabellaTestata
            oDoc.Field "RV_PODataOrdinePadre", Me.txtDataDocPadre.Value, sTabellaTestata
        End If
        'Aggiorna il documento esistente
        'If Me.chkProtocolloICE.Value = vbChecked Then
        '    AggiornaProtocolloICE 1
        'End If
       
        'oDoc.AllowCreateMovements = False
        Me.Caption = "SALVATAGGIO IN CORSO....................."
        'Me.lblInfoTesta.Font.Bold = True
        'Me.lblInfoTesta.ForeColor = vbBlue
        DoEvents
                
        If oDoc.Update = False Then
            MsgBox oDoc.LastErrorNumber
            Me.lblInfoTesta.Caption = Caption2Display(False)
            'Me.lblInfoTesta.Font.Bold = True
            'Me.lblInfoTesta.ForeColor = vbBlue
            Screen.MousePointer = 0
            DoEvents
        End If
        
       
        NuovoOggetto = False
        'oDoc.AllowCreateMovements = True
        'oDoc.CreateMovements
    
    End If

    'Se la proprietà oDoc.LastErrorNumber = 0 vuol dire che il savataggio è andato a buon fine
    If oDoc.LastErrorNumber = 0 Then
        
        'Dopo aver salvato il documento con la DmtDocs dobbiamo rinfrescare
        'il contenuto del modello ad oggetti in modo da mostrare correttamente i nuovi dati
        Screen.MousePointer = 11
        'MsgBox oDoc.IDOggetto
        Link_IDOggetto_OLD = oDoc.IDOggetto
        fnAggiornaDescrizioneDocumento
        CONSOLIDA_RETTIFICA
        
        If Me.txtNListaPrelievo.Value = 1 Then
            AGGIORNA_LINK_ORDINE_PADRE oDoc.IDOggetto
            If ATTIVA_COMMISSIONI_DA_ORDINE = 1 Then
                fnAggiornaCommissioniPerCliente
                IMPOSTA_COMMISSIONI_PER_TIPO_PEDANA Me.cdAnagrafica.KeyFieldID, oDoc.IDOggetto, Me.cboAltroSito.CurrentID
            End If
        End If
        IDTour = GET_LINK_TOUR
        If IDTour > 0 Then
            SCRIVI_RICERCA_RIGHE IDTour
        End If
        'fnAggiornaCommissioniPerCliente
        
        'Me.GrigliaCommissioni.RefreshEuroCirce
        DoEvents
        'AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI 0, 0
        
        ELIMINA_RIGHE_DOCUMENTO_VUOTE oDoc.IDOggetto, sTabellaDettaglio
        SalvaTipoLavorazioniUtilizzare oDoc.IDOggetto
        
        Screen.MousePointer = 0
        
        If NuovoOggetto = True Then
            
            RIPRISTINA_GRIGLIA_CONDIZIONI
            
        End If
        
        m_Document.OpenDoc
        'Dopo aver rinfrescato il contenuto del modello ad oggetti ci posizioniamo sul record corretto
        
        
        Cn.CursorLocation = OLD_CURSOR
        
        'oDoc.ClearValues
        
        m_Document.FindLocalData "IDOggetto = " & Link_IDOggetto_OLD, sdSearchForward
        
        m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        'm_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
        m_Semaphore.SetObjectAction m_DocType.ID, oDoc.IDOggetto, SemAllActions
        'Refresh delle variabili di stato
        m_Changed = False
        m_Search = False
        m_Saved = True
        Me.lblInfoTesta.Caption = ""
        Me.Caption = Caption2Display(False)
        'Refresh dello stato della ToolBar standard in modalità variazione
        SetStatus4Modality Modify
        'If GET_PARAMETRO_GESTIONE_CHIUSURA_LOTTI = True Then
        '    Link_Oggetto = oDoc.IDOggetto
        '    frmChiusuraConferimento.Show vbModal
        '    contArray = 0
        '    AzzerraArray
        'End If
        
        sbPopalaListaCommissioni
    Else
        sbMsgError "Si è verificato un errore durante il salvataggio del documento!" & vbCrLf & "L'errore è:" & oDoc.LastErrorNumber, TheApp.FunctionName
    End If
bSaving = 0
Exit Function
ERR_OnSave:
    MsgBox Err.Description, vbCritical, "Salvataggio"
    MsgBox oDoc.LastErrorNumber
    Me.lblInfoTesta.Caption = ""
    Me.Caption = Caption2Display(False)
    'Me.lblInfoTesta.Font.Bold = True
    'Me.lblInfoTesta.ForeColor = vbBlue
    Screen.MousePointer = 0
End Function

'**+
'Nome: OnSaveDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando OnSaveDocumentsLink
'**/
Private Sub OnSaveDocumentsLink(ByVal DocumentLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnDelete
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Delete
'**/
Private Sub OnDelete()
On Error GoTo ERR_OnDelete
Dim sToRemove As String
Dim DocLink As DmtDocManLib.DocumentsLink
Dim sSQL As String
Dim Link_Oggetto As Long
Dim Testo As String

    'Se si è in modalità tabellare potrebbe essere necessario sincronizzare
    'il documento con il record evidenziato nella browse
    If BrwMain.Visible = True Then
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    
    
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemDeleteAction) Then
        Exit Sub
    End If
    
    'Se in fase di inserimento di un nuovo
    'record non c'è niente da fare
    If oDoc.IDOggetto = 0 And Not BrwMain.Visible Then
        Exit Sub
    End If
'    If Me.lvwArticoli.ListItems.Count > 0 Then
'        MsgBox "ATTENZIONE!!" & vbCrLf & "Per cancellare totalmente il documento è necessario prima eliminare le righe manualmente", vbInformation, "Impossibile eliminare il documento"
'        Exit Sub
'    End If
    'Conferma della cancellazione
    gResource.CustomStrings.Clear
    sToRemove = Caption2Display
    gResource.CustomStrings.Add Chr(34) & sToRemove & Chr(34), 1
    If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYREMOVE), m_App.FunctionName) = vbYes Then
    
        
        'If DeleteAll = False Then
        '    MsgBox "ATTENZIONE!!" & vbCrLf & "Ci sono stati errori di elaborazione per il calcolo dei colli dei lotti disponibili" & vbCrLf & "Se il problema persiste eseguire questi comandi:" & vbCrLf & _
        '    "- Eliminare una riga documento per volta" & vbCrLf & _
        '    "- Salvare il documento" & vbCrLf & _
        '    "- Eliminare definitivamente il documento", vbCritical, "Eliminazione dati"
        '    Exit Sub
        'End If
       
        If Not (m_Document.EOF Or m_Document.BOF) Then
            'Cancella l'eventuale blocco sul record da cancellare.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        End If
        
        Link_Oggetto = oDoc.IDOggetto
                
        If Me.chkChiuso.Value = vbChecked Then
            MsgBox "Impossibile eliminare un ordine chiuso", vbCritical, TheApp.FunctionName
            Exit Sub
        End If
                
        Testo = GET_CONTROLLO_BLOCCO_ORDINE_TOUR
        If Len(Testo) > 0 Then
            MsgBox Testo, vbCritical, TheApp.FunctionName
            Exit Sub
        End If
        Testo = GET_CONTROLLO_LAVORAZIONI_ORDINE(oDoc.IDOggetto)
        If Len(Testo) > 0 Then
            MsgBox Testo, vbCritical, TheApp.FunctionName
            Exit Sub
        End If
        
        'Cancelliamo il docomento con la DmtDocs
        If oDoc.DeleteWithTO(oDoc.IDOggetto, oDoc.IDTipoOggetto) <> 0 Then
            'In ogni caso, se la cancellazione fa a buon fine, eliminiamo anche il record
            'con il modello ad oggetti in modo da tenere sincronizzata il
            
            If Not ((m_Document.EOF) And (m_Document.BOF)) Then
                m_Document.Delete
            End If
            
            GET_DATI_TOUR Link_Oggetto
            
            'ELIMINAZIONE TOUR''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = "DELETE FROM RV_POTourRighe "
            sSQL = sSQL & "WHERE IDOggettoOrdine=" & Link_Oggetto
            Cn.Execute sSQL
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If LINK_TOUR > 0 Then
                GET_CAMBIO_POSIZIONE LINK_TOUR, POSIZIONE_TOUR
            End If
            'ELIMINAZIONE TIPI DI LAVORAZIONE''''''''''''''''''''''''''''''''''''''''''''''''
            sSQL = "DELETE FROM RV_POOrdineTipoLavorazione "
            sSQL = sSQL & "WHERE IDOggetto=" & Link_Oggetto
            Cn.Execute sSQL
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            If (m_Document.EOF = True And m_Document.BOF = True) Then
                'Se è stato cancellato l'ultimo record si va in modalità inserimento
                NewRecord
            Else
                'Refresh dello stato della ToolBar standard e dei menu
                If BrwMain.Visible Then
                    'Va in modalità tabellare
                    SetStatus4Modality Browse
                Else
                    'Essendo in modalità variazione occorre controllare se il record su cui
                    'ci si è posizionati è bloccato.
                    'Se non lo è lo si blocca e si procede altrimenti si andrà in modalità tabellare.
                    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions) Then
                        'Il record è bloccato.
                        'Va in modalità tabellare
                        BrwMain.Visible = True
                        SetStatus4Modality Browse
                    Else
                        'Il record non è bloccato.
                        
                        m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
                        m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
                        
                        'Va in modalità variazione
                        SetStatus4Modality Modify
                    End If
                
                     RefreshDescriptions4StatusBar
                End If
            End If
        
            'Refresh delle variabili di stato
            m_Changed = False
            m_Saved = True
            m_Search = False
        Else
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
        End If
    End If
Exit Sub
ERR_OnDelete:
    MsgBox Err.Description, vbCritical, "OnDelete"
End Sub

'**+
'Nome: OnDeleteDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni relative ai documents link sul comando Delete
'**/
Private Sub OnDeleteDocumentsLink(ByVal DocumentLink As DmtDocManLib.DocumentsLink)
End Sub

'**+
'Nome: OnClear
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Clear
'**/
Private Sub OnClear()
'Se si è in modalità Filtro occorre ripulire i campi di immissione altrimenti,
'se si è in modalità Form, si cancella il contenuto di tutti i controlli
    
    
    If BrwMain.Visible And BrwMain.GuiMode = dgFilterDefinition Then
        '---Modalità Filtro---
        'Ripulisce i campi di immissione delle condizioni di ricerca.
        BrwMain.Conditions.ClearValues
    Else
        '---Modalità Form---
        'Ripulisce i campi del form
        ClearFormFields
        'dtData.SetFocus
        
        'Se si era in modalità Nuovo viene disabilitato il pulsante Salva
        'e si ripristina la modalità stessa.
        If m_Document.TableNew Then
            ActivateBarButtons BTN_SAVE, False
            m_Changed = False
            m_Saved = True
        End If
    End If
End Sub

'**+
'Nome: OnExecuteSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ExecuteSearch
'**/
Private Sub OnExecuteSearch()

    AGGIORNA_NUMERAZIONE_ORDINE
    'Nota: utilizzo la chiamata al metodo ApplyFilter della dmtGrid piuttosto
    'che la chiamata diretta di ExecuteSearch perchè in questo modo la dmtGrid
    'può gestire internamente le conditions di ricerca.
    'Verrà generato l'evento BrwMain_OnApplyFilter()
    '
    'ExecuteSearch
    '
    BrwMain.ApplyFilter
    
        
End Sub

'**+
'Nome: OnMoveCurrentRecord
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando di riposizionamento del record corrente
'**/
Private Sub OnMoveCurrentRecord(ByVal TIpo As Integer, ByVal sToolName As String)
    Dim iResponse As Integer
        
        
        
    iResponse = ChooseAboutSaving
    If iResponse = vbYes Then
        OnSave
        'Se la registrazione non è andata a buon fine esce
        If Not m_Saved Then
            Exit Sub
        End If
    End If
    If iResponse <> vbCancel Then
       Select Case TIpo
           Case SRCNEXT
               SearchNext
           Case SRCPREVIOUS
               SearchPrevious
       End Select
       m_Changed = False
       ActivateBarButtons BTN_SAVE, False
    End If
End Sub

'**+
'Nome: OnRepositionDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando di riposizionamento del record corrente
'per i DocumentsLink
'**/
Private Sub OnRepositionDocumentsLink(ByVal DocumentsLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnChangeView
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ChangeView
'**/
Private Sub OnChangeView(ByVal sToolName As String)
    Dim iResponse   As Integer
    
    If Not BrwMain.Visible And m_Changed Then
        iResponse = ChooseAboutSaving
        
        If iResponse = vbYes Then
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
        End If
        
        If iResponse <> vbCancel Then
            'cbc 20/04/1999
            'se si è scelto NO ripulisce i campi e va in modalità tabellare annullando
            'le ultime modifiche
            RefreshFormFields
            ChangeView sToolName
            m_Changed = False
        End If
    Else
        ChangeView sToolName
    End If
    
End Sub

'**+
'Nome: OnToolBarOptions
'
'Parametri:
'**+
'Nome: OnToolBarOptions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ToolBar
'
'
'**/
Private Sub OnToolBarOptions()
    Dim dlgToolBars As frmToolBars
    Dim bVisible As Boolean
    
    On Error Resume Next
    
    Set dlgToolBars = New frmToolBars
    'Imposta un riferimento al form chiamante
    Set dlgToolBars.FormClient = Me
    dlgToolBars.Show vbModal, Me
    Set dlgToolBars = Nothing
    
    'All'uscita dal form di dialogo la visibilità della toolbar dei filtri dipende dalla
    'visibilità del Riquadro attività e dall'impostazione fatta nel dialogo.
    bVisible = GetSetting(REGISTRY_KEY, TheApp.Name & "Settings", "Riquadro attività", True)
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: OnOptions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Option
'**/
Private Sub OnOptions()
    Dim dlgOption As frmOption
    
    Set dlgOption = New frmOption
    Set dlgOption.FormClient = Me
    dlgOption.Show vbModal, Me
    
    
    Set dlgOption = Nothing
    
    'Impedisce il 'blocco' della Toolbar alla chiusura di un form di dialogo.
    
End Sub

'**+
'Nome: OnInfo
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Info
'**/
Private Sub OnInfo()
    Dim dlgInfo As frmInformazioni
    
    Set dlgInfo = New frmInformazioni
    dlgInfo.Show vbModal, Me
    
    'Impedisce il 'blocco' della Toolbar alla chiusura di un form di dialogo.
    
End Sub
'**+
'Nome: OnPrint
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Print
'**/
'**+
'Nome: OnPrint
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Print
'**/
Private Sub OnPrint_OLD(ByVal ToolName As String)
Dim lFlags As Long
Dim OLDCursor As Integer
Dim sStr As String
Dim Field As DmtDocManLib.Field
Dim IDReportDefault As Long
Dim Testo As String

        
    
    OLDCursor = Screen.MousePointer
    
    IDReportDefault = oReportsActivity.DefaultReportID
    
    'Se il filtro attivo è "Nessun record" è possibile eseguire una stampa/esportazione soltanto se
    'si è in modalità form. In tal caso, infatti, verrà passato al Crystals Reports un filtro
    'creato ad hoc sull'ID del record attuale.
    If m_ActiveFilter.NothingSelected And BrwMain.Visible Then
        sStr = "Impossibile effettuare l'operazione richiesta." & vbCrLf
        sStr = sStr & "Prima di procedere occorre eseguire un filtro."
        sbMsgInfo sStr, m_App.FunctionName
        Screen.MousePointer = OLDCursor
        Exit Sub
    End If
    
    'Se non esiste un report attivo occorre annullare l'operazione.
   'Se non esiste un report attivo occorre annullare l'operazione.
    If Len(oReportsActivity.SelectedReportName) > 0 Then
        Set m_Report = m_DocType.Reports.Item(oReportsActivity.SelectedReportName)
    End If
    If m_Report Is Nothing Then
        sbMsgError "Impossibile eseguire - Nessun report predefinito.", m_App.FunctionName
        GoTo OnPrint_Exit
    End If
    m_iNumeroCopieDefault = m_Report.Copies
    m_OrientamentoDefault = m_Report.Orientation
    
    
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    If m_Changed Then
        Select Case ChooseAboutSavingOkCancel
            Case vbOK
                OnSave
                'Se la registrazione non è andata a buon fine esce
                If Not m_Saved Then
                    GoTo OnPrint_Exit
                End If
                
            Case vbCancel
                GoTo OnPrint_Exit
        End Select
    End If
    
    
        Set oReport = New dmtReportLib.dmtReport
        Set oReport.Connection = Cn
        oReport.Password = m_App.Password
        oReport.User = m_App.User
            
            
        'Viene inserita la condizione di ricerca basata sull'ID del record corrente.
        

        'm_DocType.Fields("IDOggetto").Value = m_Document.Fields("IDOggetto").Value
        ''m_DocType.Fields("IDOggetto").Value = oDoc.IDOggetto
        'Viene creato un filtro temporaneo per il Crystals Reports.
        ''m_DocType.RemoveFilter "Form"
            'Imposta l'idfiliale di appartenenza del documento da stampare
        oReport.BranchID = oDoc.IDFiliale 'IDFiliale
        'Imposta l'identificativo del tipo di documento
        oReport.DocTypeID = fncIDTipoOggettoPrg(App.EXEName)
            
    
    
            
    
    fncImpostaDefaultReport oReportsActivity.SelectedReportID, fnGetTipoOggetto(App.EXEName)
    
    Me.ActivityBox.Redraw = True
    DoEvents
    Select Case ToolName
    
        Case "PrePrint", "Mnu_PrePrint"
            On Error GoTo ErrorHandler
            
            Screen.MousePointer = vbHourglass
            If Not BrwMain.Visible Then
                
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            
            Else
            
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            
            End If
 
            oReport.Preview 0, 0, 0
            
            'm_Document.Preview m_Report, "", 0, 0, 0, CInt(ScaleWidth / Screen.TwipsPerPixelX), CInt(ScaleHeight / Screen.TwipsPerPixelY), True
            'lFlags = SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOMOVE
            SetWindowPos m_PreviewWindowHandle, HWND_TOP, 0, 0, 0, 0, lFlags
            
        Case "Print", "Mnu_Print"
            'PrintDocument ToolName
            If Not BrwMain.Visible Then
                    'Modalità Form - deve stampare solo il record corrente

                    fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                    oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                    oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                    oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
                        
                    oReport.DoPrint oReport.PrinterName
                        
                Else
                    Testo = "ATTENZIONE!!!" & vbCrLf
                    Testo = Testo & "Verranno stampati tutti i documenti presenti nell'elenco" & vbCrLf
                    Testo = Testo & "Si vuole procedere ugualmente?"
                    
                    If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then Exit Sub
                    
                    
                    If Not ((BrwMain.Recordset.EOF) And (BrwMain.Recordset.BOF)) Then
                        BrwMain.Recordset.MoveFirst
                        While Not Me.BrwMain.Recordset.EOF
                        
                        fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                        oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                        
                        oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                        oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
                       
                        oReport.DoPrint oReport.PrinterName
                        Me.BrwMain.Recordset.MoveNext
                        Wend
                     End If
                End If

            
            
        Case "ExportWord", "Mnu_ExportWord"

            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If

            oReport.Export recWord
            
            'ExportDocument ecWord
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Word, TheApp.Name

        Case "ExportExcel", "Mnu_ExportExcel"
            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If
        
        
        
            oReport.Export recExcel
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Excel, TheApp.Name
            
        Case "ExportHtml", "Mnu_ExportHtml"
            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If
            oReport.Export recHtml
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), HTML, TheApp.Name
        
        Case "ExportPDF", "Mnu_ExportPDF"
            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If
            
            oReport.Export recPDF
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), PDF, TheApp.Name
        
        Case "MailWord"
            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If
            SendDocument ecWord
            
            
        Case "MailExcel"
            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If
            SendDocument ecExcel
            
        Case "MailHtml"
            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If
            SendDocument ecHtml
        
        Case "MailPDF"
            If Not BrwMain.Visible Then
                fnDeleteTabellaRicorsione oDoc.IDUtente, oDoc.IDTipoOggetto
                
                oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
                oReport.Where = "IDOggetto = " & oDoc.IDOggetto
                oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            Else
                fnDeleteTabellaRicorsione TheApp.IDUser, Me.BrwMain("IDTipoOggetto").Value
                oDoc.Prepare2Print Me.BrwMain("IDAzienda").Value, TheApp.IDUser, Me.BrwMain("IDOggetto").Value, Me.BrwMain("IDTipoOggetto").Value
                
                oReport.Where = "IDOggetto = " & Me.BrwMain("IDOggetto").Value
                oReport.Where = oReport.Where & " AND IDUtente = " & TheApp.IDUser
            End If
            SendDocument ecPdf
            'oReport.SendMail recPDF
        
    End Select
    
    fncImpostaDefaultReport IDReportDefault, fnGetTipoOggetto(App.EXEName)
    Me.ActivityBox.Redraw = True
    DoEvents
    
   
OnPrint_Exit:
    Set Field = Nothing
    Screen.MousePointer = OLDCursor
    Exit Sub
    
ErrorHandler:
    Const ERROR_PRINTING_ABORTED = 3
    Const ERROR_PRINTING_CANCELLED = 4
    Select Case Err.Number
        Case 20507
            'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
            sbMsgInfo "File di report non trovato", m_App.FunctionName
        Case ERROR_PRINTING_ABORTED, ERROR_PRINTING_CANCELLED
            'non deve far niente, è stato già segnalato da CrystalReport
        Case Else
            If Len(Trim(Err.Description)) > 0 Then
                sbMsgInfo Err.Description, m_App.FunctionName
            End If
    End Select

    'Si è verificato un errore durante la procedura di anteprima.
    Screen.MousePointer = OLDCursor
    
    'Ripristina la situazione del form
    m_PreviewWindowHandle = 0
    PicForm.Visible = True

    BrwMain.Visible = m_TabMode
    ActivityBox.Visible = BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked
    FormRecalcLayout
    'SetStatus4Modality Preview, ClosePrw
        
    Set Field = Nothing
End Sub
'**+
'Nome: OnNewSearch
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando NewSearch
'**/
Private Sub OnNewSearch()
On Error GoTo ERR_OnNewSearch
    Dim iResponse As Integer

    m_FilterSelected = False
    
    If Not m_Changed Then
        NewSearch
    Else
        'cbc 20/04/1999
        'deve mostrare il messaggio con Si, No, Annulla
        iResponse = ChooseAboutSaving
        If iResponse = vbYes Then
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
        End If
        If iResponse <> vbCancel Then
            'se si è scelto NO ripristina i dati precedenti annullando le ultime modifiche
            'e predispone la modalità ricerca.
            RefreshFormFields
            NewSearch
            m_Changed = False
        End If
    
    End If
Exit Sub
ERR_OnNewSearch:
    MsgBox Err.Description, vbCritical, "OnNewSearch"
End Sub

'**+
'Nome: OnNew
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando New
'**/
Private Sub OnNew(ByVal sToolName As String)
    
    Select Case DoNewDocument
        Case vbYes
            'Si è risposto affermativamente alla
            'richiesta di Update delle modifiche apportate
            OnSave
            'Se la registrazione non è andata a buon fine esce
            If Not m_Saved Then
                Exit Sub
            End If
            NewRecord
            
        Case vbCancel
            'Si è risposto Annulla alla richiesta di Update
            Exit Sub
            
        Case Else
            'Si è premuto il tasto <No> alla richiesta di Update
            NewRecord
    End Select
End Sub

'**+
'Nome: OnNewDocumentsLink
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando New per i documentslink
'**/
Private Sub OnNewDocumentsLink(ByVal DocumentsLink As DmtDocManLib.DocumentsLink)

End Sub

'**+
'Nome: OnSummary
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Summary
'**/
Private Sub OnSummary()
    Dim lRes As Long
    
    lRes = WinHelp(hwnd, App.HelpFile, HELP_FINDER, 0)
End Sub

'**+
'Nome: OnFastHelp
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando FastHelp
'**/
Private Sub OnFastHelp()
    frmMain.WhatsThisMode
End Sub

'**+
'Nome: OnHelpOnLine
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando HelpOnLine
'**/
Private Sub OnHelpOnLine()
    Dim lRes As Long
    
    If Not ActiveControl Is Nothing Then
        If ActiveControl.HelpContextID <> 0 Then
            lRes = WinHelp(hwnd, App.HelpFile, HELP_CONTEXT, ActiveControl.HelpContextID)
        Else
            ExecuteMenuCommand "Mnu_Arg"
        End If
    End If
End Sub

'**+
'Nome: OnArg
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Arg
'**/
Private Sub OnArg()
    Dim lRes As Long
    
    If m_App.ContextHelpID <> 0 Then
        lRes = WinHelp(hwnd, App.HelpFile, HELP_CONTEXT, m_App.ContextHelpID)
    Else
        ExecuteMenuCommand "Mnu_Summary"
    End If
End Sub

'**+
'Nome: OnViewAssistant
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando ViewAssistant
'**/
Private Sub OnViewAssistant()
    BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked = Not BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked
    If BarMenu.Bands("Band_Assistant").Tools("Mnu_ViewAssistant").Checked Then
    Else
    End If
End Sub

'**/
'Autore                 : Diamante S.p.a
'
'Nome                   : OnFolders
'
'Parametri:
'
'
'Valori di ritorno:
'
'Funzionalità:
'Permette la visualizzazione o meno del DocTypeExplorer e della relativa toolbar.
'**/
Private Sub OnFolders()
    ActivityBox.Visible = Not ActivityBox.Visible
    BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked = ActivityBox.Visible
    FormRecalcLayout
    
    BarMenu.RecalcLayout
End Sub

'**+
'Nome: OnRunApplication
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando RunApplication
'**/
Private Sub OnRunApplication(ByVal sToolName As String)
End Sub
Private Sub ACSAnaDest_LostFocus()

    If (ACSAnaDest.IDAnagrafica = fnNotNullN(oDoc.Field("RV_POIDAnagraficaDestinazione", , sTabellaTestata))) Then Exit Sub
    
    sbImpostaDatiDocumento
End Sub

Private Sub ACSCliente_ChangedElement()
    If Me.ACSCliente.IDAnagrafica > 0 Then
        If Me.ACSCliente.IDAnagrafica <> Me.cdAnagrafica.KeyFieldID Then
            Me.cdAnagrafica.Load Me.ACSCliente.IDAnagrafica
        End If
    End If
End Sub

Private Sub cboAccordoCommerciale_Change()
    sbImpostaDatiDocumento

End Sub

Private Sub cboAliquotaArticolo_Click()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & IIf(IsNull(Me.cboAliquotaArticolo.CurrentID), 0, Me.cboAliquotaArticolo.CurrentID)

Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Me.txtAliquotaArticolo.Value = IIf(IsNull(rs!AliquotaIva), 0, rs!AliquotaIva)
    Else
        Me.txtAliquotaArticolo.Value = 0
    End If
rs.CloseResultset
Set rs = Nothing

CalcolaTotaleRiga

End Sub

Private Sub CboAliquotaImballo_Click()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & IIf(IsNull(Me.CboAliquotaImballo.CurrentID), 0, Me.CboAliquotaImballo.CurrentID)

Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Me.txtAliquotaImballo.Value = IIf(IsNull(rs!AliquotaIva), 0, rs!AliquotaIva)
    Else
        Me.txtAliquotaImballo.Value = 0
    End If
rs.CloseResultset
Set rs = Nothing

CalcolaTotaleRiga

End Sub
Private Sub cboAltroSito_Click()
Dim IDListinoDestinazione As Long
Dim IDListinoDefault As Long


    If Me.cboAltroSito.CurrentID <> oDoc.Field("Link_Nom_ult_sito", , sTabellaTestata) Then
        If Me.cboAltroSito.CurrentID > 0 Then
            IDListinoDestinazione = GET_LISTINO_PER_DESTINAZIONE(Me.cdAnagrafica.KeyFieldID, Me.cboAltroSito.CurrentID)
            If IDListinoDestinazione > 0 Then
                Me.cboListino.WriteOn IDListinoDestinazione
            End If
        Else
            '''''CALCOLO DEL LISTINO CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''
            IDListinoDefault = GET_LISTINO_DEFAULT(Me.cdAnagrafica.KeyFieldID)
            If IDListinoDefault = 0 Then
                IDListinoDefault = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
            End If
            
            Me.cboListino.WriteOn IDListinoDefault
            oDoc.Field "Link_Doc_Listino", IDListinoDefault, sTabellaTestata
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End If
    
    oDoc.ReadDataFromCliFoSite Me.cboAltroSito.CurrentID, sTabellaTestata
End Sub

Private Sub cboAspettoEsteriore_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub cboBancaAzienda_Click()
    sbImpostaDatiDocumento
    oDoc.Field "Link_Doc_contratto_bancario_Az", Me.cboBancaAzienda.CurrentID, sTabellaTestata
End Sub

Private Sub cboBancaCliente_Click()
    sbImpostaDatiDocumento
    oDoc.Field "Link_Nom_contratto_bancario", Me.cboBancaCliente.CurrentID, sTabellaTestata
    
End Sub


Private Sub cboCambioValuta_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT DataCambio, Valore FROM Cambio "
sSQL = sSQL & "WHERE IDCambio=" & Me.cboCambioValuta.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtValoreCambioValuta.Value = 0
    Me.txtDataCambio.Value = 0
Else
    Me.txtValoreCambioValuta.Value = fnNotNullN(rs!Valore)
    Me.txtDataCambio.Text = fnNotNull(rs!DataCambio)
End If

rs.CloseResultset
Set rs = Nothing


sbImpostaDatiDocumento
End Sub





Private Sub cboIvaCliente_Click()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    sbImpostaDatiDocumento
End Sub

Private Sub cboListino_Click()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    sbImpostaDatiDocumento
End Sub

Private Sub cboListinoAzienda_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub cboLuogoPresaMerce_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub cboMagazzino_Click()
    
    sbImpostaDatiDocumento
End Sub

Private Sub cboPagamento_Click()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    'bSaving = 1
    
    oDoc.ReadDataFromPayment Me.cboPagamento.CurrentID, sTabellaTestata
    
    oDoc.Field "Doc_data_inizio_scadenza", oDoc.DataEmissione, sTabellaTestata
    
    oDoc.Field "Link_Val_valuta", Me.cboValuta.CurrentID, sTabellaTestata
    
    oDoc.Scadenze.ParametersChanged = True
    'sbImpostaDatiDocumento

End Sub

Private Sub cboPorto_Click()
    sbImpostaDatiDocumento
End Sub


Private Sub cboRaggrFatturato_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub cboSezionale_Click()
On Error GoTo ERR_cboSezionale_Click
    oDoc.IDSezionale = Me.cboSezionale.CurrentID
    
    If oDoc.IDOggetto = 0 Then
        oDoc.Field "Doc_numero", fnDocumentNumber(Me.dtData.Text), sTabellaTestata
    Else
        If oDoc.Field("Link_Doc_sezionale", , sTabellaTestata) <> Me.cboSezionale.CurrentID Then
            oDoc.Field "Doc_numero", fnDocumentNumber(Me.dtData.Text), sTabellaTestata
        End If
    End If
    
    oDoc.Field "Link_Doc_sezionale", oDoc.IDSezionale, sTabellaTestata
Exit Sub
ERR_cboSezionale_Click:
    MsgBox Err.Description, vbCritical, "cboSezionale_Click"
End Sub

Private Sub cboTipoOrdine_Click()
    sbImpostaDatiDocumento
End Sub
Private Sub cboTrasporto_Click()
'    If Me.cboTrasporto.CurrentID = 3 Then
'        Me.cboVettore.Enabled = True
'    Else
'        Me.cboVettore.WriteOn 0
'        Me.cboVettore.Enabled = False
'    End If

    sbImpostaDatiDocumento
End Sub
Private Sub cboUnitaDiMisura_Click()
    Link_UMCoop = fnGetUMCoop(Me.cboUnitaDiMisura.CurrentID)
End Sub


Private Sub cboValuta_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM Cambio "
sSQL = sSQL & "WHERE IDValuta=" & Me.cboValuta.CurrentID
sSQL = sSQL & " AND IDValutaDiRiferimento=" & oDoc.DBDefaults.Link_Val_valuta_nazionale
sSQL = sSQL & " ORDER BY DataCambio DESC"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboCambioValuta.WriteOn 0
Else
    Me.cboCambioValuta.WriteOn fnNotNullN(rs!IDCambio)
End If

rs.CloseResultset
Set rs = Nothing

sbImpostaDatiDocumento

End Sub


Private Sub cboVettore_Click()
    oDoc.ReadDataFromCarrier Me.cboVettore.CurrentID, MainCarrier, sTabellaTestata
    sbImpostaDatiDocumento
End Sub

Private Sub cboVettoreSuccessivo_Click()
sbImpostaDatiDocumento
End Sub

Private Sub CDAgenteTesta_ChangeElement()
    sbImpostaDatiDocumento
End Sub

Private Sub cdAnagrafica_ChangeElement()
On Error Resume Next
Dim IDListinoDefault As Long
Dim IDPagamento As Long
Dim IDAnagraficaDest As Long

    AggiornaAltreDestinazioni
    AggiornaContrattiBancariCliente
    AggiornaAccordiCommerciali
    
    Me.ACSCliente.IDAnagrafica = 0
    Me.ACSCliente.Description = ""
    Me.ACSCliente.Code = ""
    Me.ACSCliente.SecondDescription = ""
    
    Me.ACSCliente.sbLoadCFByIDAnagrafica 0, Me.cdAnagrafica.KeyFieldID
    
    If oDoc.IDOggetto = 0 Then
    'Legge tutti i dati relativi al cliente selezionato
        'Me.cboPagamento.WriteOn 0
        'sbImpostaDatiDocumento
        
        oDoc.ReadDataFromCliFo cdAnagrafica.KeyFieldID
        
        LINK_CLIENTE_IVA = fnNotNullN(oDoc.Field("Link_Nom_IVA", , sTabellaTestata))
        
        oDoc.ReadDataFromAgent GET_LINK_AGENTE_CLIENTE(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm)
        
''        If GET_CONTROLLO_NUMERO_LETTERE_INTENTO(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm, Year(Me.dtData.Text)) = 1 Then
'            Me.txtIDLetteraIntento.Value = GET_LINK_LETTERA_INTENTO(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm, Year(Me.dtData.Text))
'            LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
''        End If
        Me.txtIDLetteraIntento.Value = fnNotNullN(oDoc.Field("Link_Nom_lettera_intento", , sTabellaTestata))
        LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
        Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA
        
        'If TIPO_SCONTO_CLIENTE = 1 Then
        '    Me.lngScontoDocPer.Value = fnNotNullN(oDoc.Field("Nom_sconto", , sTabellaTestata))
        'End If
        
        '''''CALCOLO DEL LISTINO CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''
        IDListinoDefault = GET_LISTINO_DEFAULT(Me.cdAnagrafica.KeyFieldID)
        If IDListinoDefault = 0 Then
            IDListinoDefault = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
        End If
        
        Me.cboListino.WriteOn IDListinoDefault
        oDoc.Field "Link_Doc_Listino", IDListinoDefault, sTabellaTestata
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If Me.dtData.Value > 0 Then
            Me.cboAccordoCommerciale.WriteOn GET_LINK_ACCORDO_COMMERCIALE_PREDEFINITO(Me.cdAnagrafica.KeyFieldID, oDoc.IDTipoAnagrafica, Me.dtData.Text)
            IDPagamento = GET_MODALITA_PAGAMENTO(Me.dtData.Text, Me.cdAnagrafica.KeyFieldID)
            If IDPagamento > 0 Then
                Me.cboPagamento.WriteOn IDPagamento
            End If
        End If
        
        If Me.cboPagamento.CurrentID = 0 Then
            Me.cboPagamento.WriteOn oDoc.DBDefaults.IDPagamentoDocDefault
        End If
        
        '''''ANAGRAFICA DI DESTINAZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
        IDAnagraficaDest = GET_LINK_CLIENTE_DESTINAZIONE(Me.cdAnagrafica.KeyFieldID)
        
        Me.ACSAnaDest.IDAnagrafica = 0
        Me.ACSAnaDest.Description = ""
        Me.ACSAnaDest.Code = ""
        Me.ACSAnaDest.SecondDescription = ""
        
        If IDAnagraficaDest > 0 Then
            Me.ACSAnaDest.sbLoadCFByIDAnagrafica 0, IDAnagraficaDest
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Else
        If oDoc.Field("Link_nom_Anagrafica", , sTabellaTestata) <> Me.cdAnagrafica.KeyFieldID Then
            Me.cboPagamento.WriteOn 0
            
            oDoc.ReadDataFromCliFo cdAnagrafica.KeyFieldID
            LINK_CLIENTE_IVA = fnNotNullN(oDoc.Field("Link_Nom_IVA", , sTabellaTestata))
            Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA
            If TIPO_SCONTO_CLIENTE = 1 Then
                Me.lngScontoDocPer.Value = fnNotNullN(oDoc.Field("Nom_sconto", , sTabellaTestata))
            End If
            If Me.dtData.Value > 0 Then
                Me.cboAccordoCommerciale.WriteOn GET_LINK_ACCORDO_COMMERCIALE_PREDEFINITO(Me.cdAnagrafica.KeyFieldID, oDoc.IDTipoAnagrafica, Me.dtData.Text)
                IDPagamento = GET_MODALITA_PAGAMENTO(Me.dtData.Text, Me.cdAnagrafica.KeyFieldID)
                If IDPagamento > 0 Then
                    Me.cboPagamento.WriteOn IDPagamento
                End If
            End If
            If Me.cboPagamento.CurrentID = 0 Then
                Me.cboPagamento.WriteOn oDoc.DBDefaults.IDPagamentoDocDefault
            End If
            
            '''''ANAGRAFICA DI DESTINAZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''
            IDAnagraficaDest = GET_LINK_CLIENTE_DESTINAZIONE(Me.cdAnagrafica.KeyFieldID)
            
            Me.ACSAnaDest.IDAnagrafica = 0
            Me.ACSAnaDest.Description = ""
            Me.ACSAnaDest.Code = ""
            Me.ACSAnaDest.SecondDescription = ""
            
            If IDAnagraficaDest > 0 Then
                Me.ACSAnaDest.sbLoadCFByIDAnagrafica 0, IDAnagraficaDest
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
    End If
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    
    GET_INFO_CLIENTE_FIDO_BLOCCO Me.cdAnagrafica.KeyFieldID
    
    If LINK_BLOCCO_CLIENTE = 1 Then
        sbMsgInfo "Il cliente risulta bloccato, pertanto non sarà possibile salvare il documento", m_App.FunctionName
    End If
    
    
    sbImpostaDatiDocumento
    
    Me.cboPorto.SetFocus
End Sub
Private Sub GET_DEFAULT_ARTICOLO_PER_VIVAIO(IDArticolo As Long)
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

N_PIANALI_PER_CARRELLO = 0
N_PIANTE_PER_PIANALE = 0
N_PROLUNGHE_PER_CARRELLO = 0
N_SETTIMANA_INIZIO = 0
N_SETTIMANA_FINE = 0


sSQL = "SELECT IDIvaVendita, AliquotaIva, Articolo, IDUnitaDiMisuraVendita, "
sSQL = sSQL & "NonRiportoIntrastat, MassaNettaInKg, IDNomenclaturaCombinata, "
sSQL = sSQL & "RV_POIDNaturaTransazione, RV_POIDCalibro, RV_POIDTipoCategoria, RV_POIDTipoLavorazione, "
sSQL = sSQL & "RV_POIDImballoVendita, IDTipoProdotto, RV_POQuantitaPerCollo, RV_POMoltiplicatore, PesoNetto, RV_POIDUnitaDiMisuraLiq, "
sSQL = sSQL & "RV_POIDTipoPesoArticolo, IDPDCAvere, RV_PO01_PianaliPerCarrello, RV_PO01_PiantePerPianale, RV_PO01_SettimanaDa, RV_PO01_SettimanaA, "
sSQL = sSQL & "RV_PO01_IDTipoPianta, RV_PO01_DiametroVaso, RV_PO01_IDArticoloPianale, RV_PO01_IDArticoloProlunga, RV_PO01_IDTipoPedana, QuantitaProlunga "

sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then

    N_PIANALI_PER_CARRELLO = fnNotNullN(rs!RV_PO01_PianaliPerCarrello)
    N_PIANTE_PER_PIANALE = fnNotNullN(rs!RV_PO01_PiantePerPianale)
    N_PROLUNGHE_PER_CARRELLO = fnNotNullN(rs!QuantitaProlunga)
    N_SETTIMANA_INIZIO = fnNotNullN(rs!RV_PO01_SettimanaDa)
    N_SETTIMANA_FINE = fnNotNullN(rs!RV_PO01_SettimanaA)

End If



rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub CDArticolo_ChangeElement()
On Error Resume Next
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim Link_Lotto As Long
Dim LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL As Long

If GESTIONE_ORDINE_VIVAIO = 1 Then
    GET_DEFAULT_ARTICOLO_PER_VIVAIO Me.CDArticolo.KeyFieldID
End If

    If Me.CDArticolo.KeyFieldID > 0 Then
        If GET_ARTICOLO_ANNULLATO(Me.CDArticolo.KeyFieldID) = True Then
            MsgBox "Articolo annullato", vbInformation, "Inserimento dati"
            Me.txtDescrizioneArticolo.Text = ""
            Me.CDArticolo.Load 0
            Exit Sub
        End If
        
        sSQL = "SELECT IDIvaVendita, AliquotaIva, Articolo, IDUnitaDiMisuraVendita, "
        sSQL = sSQL & "NonRiportoIntrastat, MassaNettaInKg, IDNomenclaturaCombinata, "
        sSQL = sSQL & "RV_POIDNaturaTransazione, RV_POIDCalibro, RV_POIDTipoCategoria, RV_POIDTipoLavorazione, "
        sSQL = sSQL & "RV_POIDImballoVendita, IDTipoProdotto, RV_POQuantitaPerCollo, RV_POMoltiplicatore, PesoNetto, RV_POIDUnitaDiMisuraLiq, "
        sSQL = sSQL & "RV_POIDTipoPesoArticolo, IDPDCAvere, RV_PO01_PianaliPerCarrello, RV_PO01_PiantePerPianale, RV_PO01_SettimanaDa, RV_PO01_SettimanaA, "
        sSQL = sSQL & "RV_PO01_IDTipoPianta, RV_PO01_DiametroVaso, RV_PO01_IDArticoloPianale, RV_PO01_IDArticoloProlunga, RV_PO01_IDTipoPedana, QuantitaProlunga "
        
        sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
        sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
        sSQL = sSQL & "WHERE IDArticolo = " & Me.CDArticolo.KeyFieldID
        
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF = False Then
            Link_Articolo = Me.CDArticolo.KeyFieldID
            
            N_PIANALI_PER_CARRELLO = fnNotNullN(rs!RV_PO01_PianaliPerCarrello)
            N_PIANTE_PER_PIANALE = fnNotNullN(rs!RV_PO01_PiantePerPianale)
            N_PROLUNGHE_PER_CARRELLO = fnNotNullN(rs!QuantitaProlunga)
            N_SETTIMANA_INIZIO = fnNotNullN(rs!RV_PO01_SettimanaDa)
            N_SETTIMANA_FINE = fnNotNullN(rs!RV_PO01_SettimanaA)
            
            
            LINK_UM_LIQUIDAZIONE = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
            LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL = fnNotNullN(rs!RV_POIDTipoPesoArticolo)
            
'             If fnNotNullN(rs!RV_POQuantitaPerCollo) = 0 Then
'                QUANTITA_PER_COLLO = 1
'            Else
'                QUANTITA_PER_COLLO = fnNotNullN(rs!RV_POQuantitaPerCollo)
'            End If
'
'            If fnNotNullN(rs!RV_POMoltiplicatore) = 0 Then
'                Moltiplicatore = 1
'            Else
'                Moltiplicatore = fnNotNullN(rs!RV_POMoltiplicatore)
'            End If
'
'            PESO_LORDO = fnNotNullN(rs!PesoNetto)

            If LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL = 0 Then
                TIPO_PESO_ARTICOLO = TIPO_PESO_ARTICOLO
            Else
                TIPO_PESO_ARTICOLO = LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL
            End If
            
            
            If ((bVariazioneDettaglio = False) Or (Me.CDArticolo.KeyFieldID <> oDoc.Field("Link_art_articolo", , sTabellaDettaglio))) Then
                If LINK_CLIENTE_IVA = 0 Then
                    Me.cboAliquotaArticolo.WriteOn fnNotNullN(rs!IDIvaVendita)
                Else
                    Me.cboAliquotaArticolo.WriteOn LINK_CLIENTE_IVA
                End If
                
                Me.txtDescrizioneArticolo.Text = fnNotNull(rs!Articolo)
                Me.cboCalibro.WriteOn fnNotNullN(rs!RV_POIDCalibro)
                Me.cboTipoCategoria.WriteOn fnNotNullN(rs!RV_POIDTipoCategoria)
                Me.cboTipoLavorazione.WriteOn fnNotNullN(rs!RV_POIDTipoLavorazione)
                Me.cboUnitaDiMisura.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
                Me.CDImballo.Load fnNotNullN(rs!RV_POIDImballoVendita)
                Me.txtQuantitaPerCollo.Value = fnNotNullN(rs!RV_POQuantitaPerCollo)
                Me.txtPesoPerCollo.Value = fnNotNullN(rs!PesoNetto)
                Me.txtMoltiplicatore.Value = fnNotNullN(rs!RV_POMoltiplicatore)
                
                If (Me.CDImballo.KeyFieldID = 0) Then
                    If DATI_DA_CONTRATTO = False Then
                        GET_IMBALLO_PER_ARTICOLO_CONF Me.CDArticolo.KeyFieldID
                    End If
                End If
                
                'VIVAIO
                Me.CDTipoPedana.Load fnNotNullN(rs!RV_PO01_IDTipoPedana)
                Me.CDArticoloPianale.Load fnNotNullN(rs!RV_PO01_IDArticoloPianale)
                Me.CDArticoloProlunga.Load fnNotNullN(rs!RV_PO01_IDArticoloProlunga)
                Me.ACSSocio.sbLoadCFByIDAnagrafica 7, GET_LINK_FORNITORE(Me.CDArticolo.KeyFieldID)
                
                If TIPO_PESO_ARTICOLO <= 1 Then
                    Me.txtPesoLordo.Value = Me.txtPesoPerCollo.Value * Me.txtColli.Value
                Else
                    Me.txtPesoNetto.Value = Me.txtPesoPerCollo.Value
                    Me.txtPesoLordo.Value = Me.txtPesoNetto.Value + Me.txtTara.Value
                End If
                If DATI_DA_CONTRATTO = False Then
                    GET_CONFIGURAZIONE_IMPORTI Me.cdAnagrafica.KeyFieldID, Me.CDArticolo.KeyFieldID, fnNotNullN(oDoc.Field("Link_Doc_listino", , sTabellaTestata)), fnNotNullN(oDoc.Field("Link_Doc_listino_base", , sTabellaTestata)), Me.txtQta_UM.Value
                End If
                'VIVAIO
                GET_CALCOLO_VIVAIO txtQta_UM.Value
            
            End If
            
            
            
            If DATI_DA_CONTRATTO = False Then
                If CDTipoPedana.KeyFieldID = 0 Then Me.CDTipoPedana.Load GET_LINK_PEDANA_CLIENTE(Me.cdAnagrafica.KeyFieldID)
                If fnGestioneArticoli <= 1 Then
                    If rs!IDTipoProdotto = Link_TipoGrezzo Then
                        MsgBox "ATTENZIONE!!!" & vbCrLf & "Stai inserendo un articolo di tipo GREZZO", vbInformation, "Inserimento articoli"
                    End If
                End If
                rs.CloseResultset
                Set rs = Nothing
                            
                CalcolaTotaleRiga
                If GESTIONE_ORDINE_VIVAIO = 0 Then
                    If Me.CDTipoPedana.KeyFieldID > 0 Then
                        If Me.txtQuantitaPedana.Value > 0 Then
                            If Me.CDImballo.KeyFieldID > 0 Then
                                If Me.txtQuantitaPerPedana.Value = 1 Then
                                    If Me.txtQuantitaPerPedana.Enabled = True Then
                                        If (bVariazioneDettaglio = False) Then
                                            Me.txtQuantitaPerPedana.SetFocus
                                        End If
                                    End If
        '                            If Me.txtImportoUnitarioImballo.Enabled = True Then
        '                                Me.txtImportoUnitarioImballo.SetFocus
        '                            End If
                                Else
                                    If Me.txtColli.Enabled = True Then
                                        Me.txtColli.SetFocus
                                    End If
                                End If
                            Else
                                If Me.CDImballo.Enabled = True Then
                                    Me.CDImballo.SetFocus
                                End If
                            End If
                        Else
                            If Me.txtQuantitaPedana.Enabled = True Then
                                Me.txtQuantitaPedana.SetFocus
                            End If
                        End If
                    Else
                        If Me.CDTipoPedana.Enabled = True Then
                            Me.CDTipoPedana.SetFocus
                        End If
                    End If
                Else
                    Me.CDImballo.SetFocus
                    
                    If Me.CDImballo.KeyFieldID > 0 Then
                        Me.txtColli.SetFocus
                    End If
                End If
            End If
        End If
    End If




End Sub



Private Sub CDImballo_ChangeElement()
'On Error GoTo ERR_CDImballo_ChangeElement
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
oDoc.Tables(sTabellaDettaglio).SetActiveRetail A_Riga(1)

If Me.CDImballo.KeyFieldID > 0 Then
    If Me.CDImballo.KeyFieldID <> fnNotNullN(oDoc.Field("Link_Art_Articolo", , sTabellaDettaglio)) Then
        If GET_ARTICOLO_ANNULLATO(Me.CDImballo.KeyFieldID) = True Then
            MsgBox "Articolo annullato", vbInformation, "Inserimento dati"
            Me.txtDescrizioneImballo.Text = ""
            Me.CDImballo.Load 0
            Exit Sub
        End If
        sSQL = "SELECT Articolo.IDIvaVendita, Iva.AliquotaIva, Articolo.Tara, Articolo.Articolo, Articolo.NonRiportoIntrastat, Articolo.MassaNettaInKg, "
        sSQL = sSQL & "Articolo.IDNomenclaturaCombinata , Articolo.RV_POIDNaturaTransazione, RV_POTipoImballo.Rendere, Articolo.IDPDCAvere "
        sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
        sSQL = sSQL & "RV_POTipoImballo ON Articolo.RV_POIDTipoImballo = RV_POTipoImballo.IDRV_POTipoImballo LEFT OUTER JOIN "
        sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
        sSQL = sSQL & "WHERE IDArticolo = " & Me.CDImballo.KeyFieldID
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF = False Then
            If FLAG_IVA_IMBALLO_A_RENDERE = 1 Then
                If Abs(fnNotNullN(rs!Rendere)) = 1 Then
                    Me.CboAliquotaImballo.WriteOn fnNotNullN(rs!IDIvaVendita)
                Else
                    If FLAG_IVA_UGUALE = 1 Then
                        If LINK_CLIENTE_IVA = 0 Then
                            If Me.CDArticolo.KeyFieldID > 0 Then
                                Me.CboAliquotaImballo.WriteOn Me.cboAliquotaArticolo.CurrentID
                            Else
                                Me.CboAliquotaImballo.WriteOn fnNotNullN(rs!IDIvaVendita)
                            End If
                        Else
                            Me.CboAliquotaImballo.WriteOn LINK_CLIENTE_IVA
                        End If
                    Else
                        Me.CboAliquotaImballo.WriteOn fnNotNullN(rs!IDIvaVendita)
                    End If
                End If
            Else
                If FLAG_IVA_UGUALE = 1 Then
                    If LINK_CLIENTE_IVA = 0 Then
                        If Me.CDArticolo.KeyFieldID > 0 Then
                            Me.CboAliquotaImballo.WriteOn Me.cboAliquotaArticolo.CurrentID
                        Else
                            Me.CboAliquotaImballo.WriteOn fnNotNullN(rs!IDIvaVendita)
                        End If
                    Else
                        Me.CboAliquotaImballo.WriteOn LINK_CLIENTE_IVA
                    End If
                Else
                    If LINK_CLIENTE_IVA > 0 Then
                        Me.CboAliquotaImballo.WriteOn LINK_CLIENTE_IVA
                    Else
                        Me.CboAliquotaImballo.WriteOn fnNotNullN(rs!IDIvaVendita)
                    End If
                End If
            End If

            Me.txtTaraUnitaria.Value = fnNotNullN(rs!Tara)
            Me.txtDescrizioneImballo.Text = fnNotNull(rs!Articolo)

            
        End If
         
        rs.CloseResultset
        Set rs = Nothing
        
        If DATI_DA_CONTRATTO = False Then
            GET_CONFIGURAZIONE_IMPORTI_IMBALLI Me.cdAnagrafica.KeyFieldID, Me.CDImballo.KeyFieldID, fnNotNullN(oDoc.Field("Link_Doc_listino", , sTabellaTestata)), fnNotNullN(oDoc.Field("Link_Doc_listino_base", , sTabellaTestata)), Me.txtColli.Value
        End If
        
        If bVariazioneDettaglio = False Then
            If DATI_DA_CONTRATTO = False Then
                Me.txtQuantitaPerPedana.Value = GET_QUANTITA_IMBALLO_PER_TIPO_PEDANA(Me.CDTipoPedana.KeyFieldID, Me.CDImballo.KeyFieldID)
                Me.txtColli.Value = Me.txtQuantitaPedana.Value * Me.txtQuantitaPerPedana.Value
                Me.chkImportoImballoInArticolo.Value = GET_PREZZO_IMBALLO_INCLUSO(Me.CDImballo.KeyFieldID, Me.cdAnagrafica.KeyFieldID)
                GET_CONFEZIONI_IMBALLO_PER_ARTICOLO Me.CDArticolo.KeyFieldID, Me.CDImballo.KeyFieldID
            End If
        End If
        
        If (txtQuantitaPerPedana.Enabled = True) Then
            Me.txtQuantitaPerPedana.SetFocus
        End If
        
'        If Me.txtImportoUnitarioImballo.Enabled = True Then
'            Me.txtImportoUnitarioImballo.SetFocus
'        End If
        
        CalcolaTotaleRiga
    End If
        
End If


Exit Sub
ERR_CDImballo_ChangeElement:
    MsgBox Err.Description, vbCritical, "CDImballo_ChangeElement"
End Sub
Private Sub CDImballoPrimario_ChangeElement()
    If bVariazioneDettaglio = False Then GET_PROP_CONFEZIONI Me.CDArticolo.KeyFieldID, Me.CDImballo.KeyFieldID, Me.CDImballoPrimario.KeyFieldID

    If bLoadingRiga = False Then txtColli_LostFocus
End Sub
Private Sub CDPedana_ChangeElement()
If Me.CDPedana.KeyFieldID > 0 Then
    If Me.CDPedana.KeyFieldID <> fnNotNullN(oDoc.Field("RV_POIDArticoloPedana", , sTabellaDettaglio)) Then
        If GET_ARTICOLO_ANNULLATO(Me.CDPedana.KeyFieldID) = True Then
            MsgBox "Articolo annullato", vbInformation, "Inserimento dati"
            Me.CDTipoPedana.Load 0
            Me.CDPedana.Load 0
        End If
    End If
End If
End Sub

Private Sub CDTipoPedana_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If ((bVariazioneDettaglio = False) Or (Me.CDTipoPedana.KeyFieldID <> oDoc.Field("RV_POIDTipoPedana", , sTabellaDettaglio))) Then
    sSQL = "SELECT IDArticoloImballo "
    sSQL = sSQL & "FROM RV_POTipoPedana "
    sSQL = sSQL & "WHERE IDRV_POTipoPedana=" & Me.CDTipoPedana.KeyFieldID
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        Me.CDPedana.Load 0
    Else
        Me.CDPedana.Load fnNotNullN(rs!IDArticoloImballo)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    Me.txtQuantitaPerPedana.Value = GET_QUANTITA_IMBALLO_PER_TIPO_PEDANA(Me.CDTipoPedana.KeyFieldID, Me.CDImballo.KeyFieldID)
    
    Me.txtColli.Value = Me.txtQuantitaPedana.Value * Me.txtQuantitaPerPedana.Value + Me.txtColliSfusi.Value
    
    txtColli_LostFocus
End If
End Sub

Private Sub chkChiuso_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub chkConfDaContratto_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub chkLordoIVA_Click()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    sbImpostaDatiDocumento
End Sub



Private Sub chkOrdineCompletato_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub chkRaggruppaScadenze_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub chkRaggruppBolle_Click()
    sbImpostaDatiDocumento
End Sub
Private Sub cmdAnalizzaOrdine_Click()
On Error GoTo ERR_cmdAnalizzaOrdine_Click
    If oDoc.IDOggetto > 0 Then
        LINK_ORDINE_SELEZIONATO = fnNotNullN(oDoc.Field("RV_POIDOrdinePadre", , sTabellaTestata))
        LINK_ART_ORD_PADRE_SEL = 0
        frmAnalizzaOrdine.Show vbModal
    End If
Exit Sub

ERR_cmdAnalizzaOrdine_Click:
    MsgBox Err.Description, vbCritical, "cmdAnalizzaOrdine_Click"
End Sub

Private Sub cmdConferimento_Click()

    frmConferimento.Show vbModal
    
End Sub

Private Sub cmdDuplicaRiga_Click()
On Error Resume Next
    If bVariazioneDettaglio = False Then Exit Sub
    
    bVariazioneDettaglio = False
    
    Me.txtColli.SetFocus
    
    lblStatoRiga.Caption = "RIGA DUPLICATA"
    
    RIGA_DUPLICATA = 1
End Sub

Private Sub cmdElencoRett_Click()
    
    LINK_RIGA_SELEZIONATA_RETT = fnNotNullN(oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio))
    
    frmRettifica.Show vbModal
    
End Sub

Private Sub cmdElimina_Click()
Dim Testo As String
     'Se è stato selezionato una riga nella listview degli articoli
    
    
    If Not lvwArticoli.SelectedItem Is Nothing Then
        'Rimuoviamo il dettaglio selezionato dall'oggetto cDocument
        
        If GET_CONTROLLO_COLLEGAMENTO_CONF(NumeroRecordPerModifica) = True Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "La riga dell'ordine è collegata ad un conferimento, continuando con questa operazione il collegamento andrà perso"
            Testo = Testo & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            
            If (MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio dati")) = vbNo Then
                
                Exit Sub
            End If
        End If
        
        
        If oDoc.Field("RV_PORigaCompleta", , sTabellaDettaglio) = 1 Then
            
            If oDoc.Field("IDValoriOggettoDettaglio", , sTabellaDettaglio) > 0 Then
                If fnNotNullN(oDoc.Field("RV_POIDConferimentoRighe", , sTabellaDettaglio)) > 0 Then
                    ArrayConfMod(contArray) = fnNotNullN(oDoc.Field("RV_POIDConferimentoRighe", , sTabellaDettaglio))
                    contArray = contArray + 1
                End If
            End If
                
            SALVA_RETTIFICA
                
            If A_Riga(0) > 0 Then oDoc.Tables(sTabellaDettaglio).RemoveRetail A_Riga(0)
            If A_Riga(1) > 0 Then oDoc.Tables(sTabellaDettaglio).RemoveRetail A_Riga(1) - 1
        Else
            If oDoc.Field("IDValoriOggettoDettaglio", , sTabellaDettaglio) > 0 Then
                If fnNotNullN(oDoc.Field("RV_POIDConferimentoRighe", , sTabellaDettaglio)) > 0 Then
                    ArrayConfMod(contArray) = fnNotNullN(oDoc.Field("RV_POIDConferimentoRighe", , sTabellaDettaglio))
                    contArray = contArray + 1
                End If
            End If
            
            SALVA_RETTIFICA
            
            If A_Riga(0) > 0 Then oDoc.Tables(sTabellaDettaglio).RemoveRetail A_Riga(0)
            If A_Riga(1) > 0 Then oDoc.Tables(sTabellaDettaglio).RemoveRetail A_Riga(1)
        End If
        
        'Aggiorna il contenuto della listview articoli
        sbPopalaListaArticoli False
        'Ricalcola il totale del documento
        sbCalcolaDocumento
        'Si predispone per l'inserimento di un nuovo dettaglio
        cmdNuovo_Click
    End If
    
    
End Sub



Private Sub cmdEliminaRifContratto_Click()
Dim Testo As String

If txtIDContratto.Value = 0 Then Exit Sub

Testo = "ATTENZIONE!!!" & vbCrLf
Testo = Testo & "Sei sicuro di voler eliminare il riferimento del contratto?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Elimina riferimento contratto") = vbNo Then Exit Sub

Me.txtIDContratto.Value = 0


End Sub

Private Sub cmdEliminaRifLetInt_Click()
On Error GoTo ERR_cmdEliminaRifLetInt_Click
Dim Testo As String
If Me.txtIDLetteraIntento.Value = 0 Then Exit Sub
Testo = "Sei sicuro di voler eliminare il riferimento alla lettera d'intento?" & vbCrLf
If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento lettera d'intento") = vbNo Then Exit Sub

Me.txtIDLetteraIntento.Value = 0

LINK_CLIENTE_IVA = GET_LINK_IVA_CLIENTE(Me.cdAnagrafica.KeyFieldID)

LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
    
Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA
Exit Sub
ERR_cmdEliminaRifLetInt_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRifLetInt_Click"
End Sub

Private Sub cmdLavorazioni_Click()
If oDoc.IDOggetto > 0 Then
    LINK_ORDINE_PADRE = fnNotNullN(oDoc.Field("RV_POIDOrdinePadre", , sTabellaTestata))
    frmAssegnazioneMerce.Show vbModal
End If
End Sub

Private Sub cmdLetteraIntento_Click()
On Error GoTo ERR_cmdLetteraIntento_Click
    If Me.cdAnagrafica.KeyFieldID = 0 Then Exit Sub
    frmLetteraIntento.Show vbModal
    
    If Me.txtIDLetteraIntento.Value > 0 Then
        LINK_CLIENTE_IVA = GET_LINK_IVA_CLIENTE(Me.cdAnagrafica.KeyFieldID)
        
        LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
            
        Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA
    End If
    sbImpostaDatiDocumento
Exit Sub
ERR_cmdLetteraIntento_Click:
    MsgBox Err.Description, vbCritical, "cmdLetteraIntento_Click"
End Sub

Private Sub cmdListaPrelievo_Click()
On Error GoTo ERR_cmdListaPrelievo_Click
Dim LINK_OGGETTO_DA_COPIARE As Long
Dim LINK_ANAGRAFICA As Long
Dim LINK_DESTINAZIONE As Long
Dim DATA_DOCUMENTO_PADRE As String
Dim NUMERO_DOCUMENTO_PADRE As Long
Dim Testo As String
    
    If oDoc.IDOggetto = 0 Then Exit Sub

    If LINK_SEZIONALE_LISTA = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = "Per poter utilizzare questa funzione "
        Testo = "impostare il sezionale per la lista prelievo nei parametri cooperativa." & vbCrLf
        
        MsgBox Testo, vbCritical, "Creazione nuova lista prelievo"
        
        Exit Sub
    End If
        
    If LINK_OGGETTO_ORDINE_PADRE_REGISTRY = 0 Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "Con questo comando verrà creato una nuova lista di prelievo per l'ordine numero " & Me.txtNOrdinePadre.Text & " del " & Me.txtDataDocPadre.Text & vbCrLf
        Testo = Testo & "Continuare?"
        
        If MsgBox(Testo, vbQuestion + vbYesNo, "Creazione nuova lista prelievo") = vbNo Then Exit Sub
        
        
        If (m_Changed = True) Then
        
            If OnSave = False Then Exit Sub
            
        End If
    
    End If
    
    Screen.MousePointer = 11
        
    LINK_OGGETTO_DA_COPIARE = fnNotNullN(oDoc.Field("RV_POIDOrdinePadre", , sTabellaTestata))
    
    LINK_ANAGRAFICA = Me.cdAnagrafica.KeyFieldID
    LINK_DESTINAZIONE = Me.cboAltroSito.CurrentID
    DATA_DOCUMENTO_PADRE = Me.txtDataDocPadre.Text
    NUMERO_DOCUMENTO_PADRE = Me.txtNOrdinePadre.Value

    
    Me.Caption = "INSERIMENTO NUOVO LISTA PRELIEVO..............."
    DoEvents
    
    NewRecord
    Me.cdAnagrafica.Load 0
    Me.cboSezionale.WriteOn LINK_SEZIONALE_LISTA
    Me.cdAnagrafica.Load LINK_ANAGRAFICA
    Me.cboAltroSito.WriteOn LINK_DESTINAZIONE
    Me.txtDataDocPadre.Text = DATA_DOCUMENTO_PADRE
    Me.txtNOrdinePadre.Value = NUMERO_DOCUMENTO_PADRE
    Me.txtNListaPrelievo.Value = GET_NUMERO_LISTA_PRELIEVO(LINK_OGGETTO_DA_COPIARE)
    
    oDoc.Field "RV_POIDOrdinePadre", LINK_OGGETTO_DA_COPIARE, sTabellaTestata
    oDoc.Field "RV_PONumeroOrdinePadre", Me.txtNOrdinePadre.Value, sTabellaTestata
    oDoc.Field "RV_PODataOrdinePadre", Me.txtDataDocPadre.Value, sTabellaTestata
    oDoc.Field "RV_PONumeroListaPrelievo", Me.txtNListaPrelievo.Value, sTabellaTestata
    oDoc.Field "RV_POOrdineCompletato", 0, sTabellaTestata
    
    IMPOSTA_DATI_LISTA_PRELIEVO LINK_OGGETTO_DA_COPIARE
    
    ABILITA_CONTROLLI
    
    
    Me.Caption = Caption2Display
    
    Screen.MousePointer = 0
    'Me.cdAnagrafica.SetFocus
Exit Sub
ERR_cmdListaPrelievo_Click:
    MsgBox Err.Description, vbCritical, "cmdListaPrelievo_Click"
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdNuovo_Click()


    If Me.txtNListaPrelievo.Value > 1 Then Exit Sub

    bVariazioneDettaglio = False

    Me.CDArticolo.Load 0
    Link_Articolo = 0
    Me.txtDescrizioneArticolo.Text = ""
    Me.cboAliquotaArticolo.WriteOn 0
    Me.txtAliquotaArticolo.Value = 0
    Me.cboUnitaDiMisura.WriteOn 0
    Me.txtImportoUnitarioArticolo.Value = 0
    
    Me.chkImportoImballoInArticolo.Value = vbUnchecked
    
    Me.CDImballo.Load 0
    Me.txtDescrizioneImballo.Text = ""
    Me.txtTaraUnitaria.Value = 0
    Me.CboAliquotaImballo.WriteOn 0
    Me.txtAliquotaImballo.Value = 0
    Me.txtImportoUnitarioImballo.Value = 0
    Me.txtColli.Value = 0
    
    Me.txtPesoLordo.Value = 0
    Me.txtPesoNetto.Value = 0
    Me.txtTara.Value = 0
    Me.txtPezzi.Value = 0
    Me.txtQta_UM.Value = 0
   
    Me.CDPedana.Load 0
    'Me.txtDescrizionePedana.Text = ""
    Me.txtQuantitaPedana.Value = 0
    Me.txtColliSfusi.Value = 0
    txtQuantitaPerPedana.Value = 0
    
    Me.txtQuantitaPerCollo.Value = 0
    Me.cboUMRigaOrdine.WriteOn LINK_TIPO_UM_RIGA_ORDINE
    Me.CDTipoPedana.Load 0
    Me.txtSconto1.Value = 0
    Me.txtSconto2.Value = 0
    Me.txtImponibileUnitario.Value = 0
    Me.txtAnnotazioniDiRiga.Text = ""
    Me.txtAnnotazioniDiRigaLav.Text = ""
    'IMBALLO PRIMARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Me.CDImballoPrimario.Load 0
    Me.txtTaraConfImballo.Value = 0
    Me.txtNumeroConfImballo.Value = 0
    
    

    
    'If TIPO_SCONTO_CLIENTE = 2 Then
    '    Me.txtSconto1.Value = fnNotNullN(oDoc.Field("Nom_sconto", , sTabellaTestata))
    'End If
    
    Me.txtImportoListinoArticolo.Value = 0
    Me.txtScontoImpListino.Value = 0
    
    Me.cboCalibro.WriteOn 0
    Me.cboTipoCategoria.WriteOn 0
    Me.cboTipoLavorazione.WriteOn 0
    
    
    Me.CDArticoloPianale.Load 0
    Me.CDArticoloProlunga.Load 0
    
    Me.ACSSocio.IDAnagrafica = 0
    Me.ACSSocio.Description = ""
    Me.ACSSocio.SecondDescription = ""
    
    N_PIANALI_PER_CARRELLO = 0
    N_PIANTE_PER_PIANALE = 0
    N_PROLUNGHE_PER_CARRELLO = 0
    N_SETTIMANA_INIZIO = 0
    N_SETTIMANA_FINE = 0
    
    Me.txtAnnotazioniPerSocio.Text = ""
    Me.txtRaggrRigaOrdine.Text = ""
    Me.txtQuantitaPedanaEff.Value = 0
    
    Me.cboReportNew.WriteOn 0
    Me.cboReportPedNew.WriteOn 0
    
    'GESTIONE RAGGRUPPAMENTO ORDINE
    LINK_LOTTO_PROD_LAV = 0
    Me.txtRaggrRigaOrdine.Locked = False
    If (ATTIVA_SEL_LOTTO_PROD_IN_LAV = 1) Then
        Me.txtRaggrRigaOrdine.Locked = True
        Me.txtRaggrRigaOrdine.Text = GET_DESCRIZIONE_LOTTO_PROD_LAV(LINK_LOTTO_PROD_LAV)
    End If
    
    CalcolaTotaleRiga
    
    Me.CDArticolo.SetFocus
    
    A_Riga(0) = 0
    A_Riga(1) = 0
    Me.txtAndamentoOrdineDett.Text = ""
    lblStatoRiga.Caption = "NUOVA RIGA"
    
    RIGA_DUPLICATA = 0
    
    
End Sub
Private Function GET_TIPO_IMPORTO_ARTICOLO_DA_LIQUIDAZIONE() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDTipoImportoArticolo "
sSQL = sSQL & "FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

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
Private Sub cmdSalva_Click()

If Me.txtNListaPrelievo.Value > 1 Then Exit Sub

AggiornaRiga = 1

If PermessoSalvataggio = True Then
    If bVariazioneDettaglio = False Then
        NumeroRiga = NumeroRiga + 1
    End If
        
        SalvataggioRiga
        
        oDoc.PerformTable sTabellaDettaglio, False
        
        'Aggiorna il contenuto della listview degli articoli
        sbPopalaListaArticoli False
        'Ricalcola il documento
        sbCalcolaDocumento
        
        AggiornaRiga = 0
        
        GET_TOTALI_PEDANE
        
        'Se eravamo in presenza di un nuovo dettaglio
        If RIGA_DUPLICATA = 0 Then
            If bVariazioneDettaglio = False Then
                If lvwArticoli.ListItems.Count > 0 Then
                    'Si predispone per l'inserimento di un altro dettaglio
                    Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(lvwArticoli.ListItems.Count)
                    lvwArticoli.SelectedItem.EnsureVisible
                End If
                cmdNuovo_Click
            Else
                Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(NumeroRecordLista)
                lvwArticoli.SelectedItem.EnsureVisible
            End If
        Else
            Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(lvwArticoli.ListItems.Count)
            lvwArticoli.SelectedItem.EnsureVisible
            
            lblStatoRiga.Caption = "RIGA IN MODIFICA"
        End If
        
        RIGA_DUPLICATA = 0
        
End If
End Sub
Private Function PermessoSalvataggio() As Boolean
Dim Testo As String

PermessoSalvataggio = True

If Me.CDArticolo.KeyFieldID > 0 Then
    If Me.cboUnitaDiMisura.CurrentID = 0 Then
        MsgBox "Deve essere impostata l'unità di misura dell'articolo", vbCritical, "Impossibile salvare"
        PermessoSalvataggio = False
        Me.cboUnitaDiMisura.SetFocus
    
        Exit Function
    End If
End If

If Me.CDArticolo.KeyFieldID > 0 Then
    If bVariazioneDettaglio = False Then
        If Me.CDImballo.KeyFieldID = 0 Then
            Testo = "L'imballo non è stato inserito" & vbCrLf
            Testo = Testo & "Vuoi continuare l'operazione di salvataggio?"
            
            If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then
                PermessoSalvataggio = False
                Me.CDImballo.SetFocus
                Exit Function
            End If
        End If
    End If
End If

'If Link_Articolo > 0 Then
'    If Me.cboAliquotaArticolo.CurrentID <> Me.CboAliquotaImballo.CurrentID Then
'        If ControllaParametroIvaBloccata = True Then
'
'            If Me.CDImballo.KeyFieldID > 0 Then
'                MsgBox "ATTENZIONE!!!" & vbCrLf & "Per impostazioni nei parametri della filiale l'aliquota IVA dell'imballo deve essere uguale all'aliquota IVA del prodotto venduto", vbInformation, "Controllo parametri filiale"
'                PermessoSalvataggio = False
'                Exit Function
'
'            End If
'
'        End If
'    End If
'End If
If Me.CDArticolo.KeyFieldID > 0 Then
    If Me.cboAliquotaArticolo.CurrentID <> Me.CboAliquotaImballo.CurrentID Then
        If Me.CDImballo.KeyFieldID > 0 Then
            If ControllaParametroIvaBloccata(Me.CDImballo.KeyFieldID) = True Then
                MsgBox "ATTENZIONE!!!" & vbCrLf & "Per impostazioni nei parametri della filiale l'aliquota IVA dell'imballo deve essere uguale all'aliquota IVA del prodotto venduto", vbInformation, "Controllo parametri filiale"
                PermessoSalvataggio = False
                Exit Function
            End If
        End If
    End If
End If
If Me.CDArticolo.KeyFieldID > 0 Then
    If Me.txtColli.Value <= 0 Then
        MsgBox "ATTENZIONE!!!" & vbCrLf & "I colli devono essere maggiori di zero"
        PermessoSalvataggio = False
        Exit Function
    End If
End If

If (Me.CDArticolo.KeyFieldID > 0) Or (Me.CDImballo.KeyFieldID > 0) Then
    If Me.txtColli.Value <= 0 Then
        MsgBox "ATTENZIONE!!!" & vbCrLf & "I colli devono avere un valore maggiore di zero", vbInformation, "Salvataggio dati"
        PermessoSalvataggio = False
        Exit Function
    End If
End If
If (Me.CDImballo.KeyFieldID > 0) Then
    If ((Me.txtImportoUnitarioImballo.Value = 0) And (Me.chkImportoImballoInArticolo.Value = vbChecked)) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "È stato impostato il visto dell'importo imballo incluso nell'importo della merce, ma l'importo unitario dell'imballo non è stato impostato" & vbCrLf
        Testo = Testo & "Vuoi continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then
             PermessoSalvataggio = False
             Exit Function
        End If
    End If
    
End If

If GET_CONTROLLO_COLLEGAMENTO_CONF(NumeroRecordPerModifica) = True Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "La riga dell'ordine è collegata ad un conferimento, continuando con questa operazione il collegamento andrà perso"
    Testo = Testo & vbCrLf
    Testo = Testo & "Vuoi continuare?"
    
    If (MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio dati")) = vbNo Then
        PermessoSalvataggio = False
        Exit Function
    End If
End If

End Function
Private Function ControllaParametroIvaBloccata(IDImballo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IvaBloccata, IvaImballoARendere FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDAzienda=" & m_App.IDFirm & " AND "
sSQL = sSQL & "IDFIliale=" & m_App.Branch & " AND "
sSQL = sSQL & "IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    ControllaParametroIvaBloccata = False
Else
    If fnNotNullN(rs!IvaBloccata) = 0 Then
        ControllaParametroIvaBloccata = False
    Else
        If fnNotNullN(rs!IvaImballoARendere) = 1 Then
            ControllaParametroIvaBloccata = False
        Else
            ControllaParametroIvaBloccata = fnNotNullN(rs!IvaBloccata)
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub cmdSalvaComeNuovo_Click()
On Error GoTo ERR_cmdSalvaComeNuovo_Click
Dim LINK_OGGETTO_DA_COPIARE As Long
Dim LINK_ANAGRAFICA As Long
Dim LINK_PAGAMENTO As Long
Dim LINK_DESTINAZIONE As Long
Dim LINK_LUOGO_PRESA_MERCE As Long
Dim LINK_AGENTE As Long
Dim CAUSALE_TRASP As String
Dim LINK_TIPO_ORD As Long

Dim LINK_PORTO As Long
Dim LINK_TRASP As Long
Dim LINK_VETTORE As Long
Dim TARGA_AUT As String
Dim ISTR_MITT As String
Dim LINK_VETTORE_SUCC As Long
Dim LINK_ASPETTO As Long

Dim NOTE_FATT As String
Dim NOTE_INTERNE As String
Dim NOTE_EVASIONE As String

Dim LINK_CONTRATTO_COLLEGATO As Long

Dim Testo As String
    
    If oDoc.IDOggetto = 0 Then Exit Sub
    
    Testo = "ATTENZIONE!!!!" & vbCrLf
    Testo = Testo & "Con questo comando verrà creato un nuovo documento simile a questo visualizzato" & vbCrLf
    Testo = Testo & "Continuare?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Creazione nuovo documento") = vbNo Then Exit Sub

    If OnSave = False Then Exit Sub
    
    LINK_OGGETTO_DA_COPIARE = oDoc.IDOggetto
    LINK_ANAGRAFICA = Me.cdAnagrafica.KeyFieldID
    LINK_DESTINAZIONE = Me.cboAltroSito.CurrentID
    LINK_PAGAMENTO = Me.cboPagamento.CurrentID
    LINK_LUOGO_PRESA_MERCE = Me.cboLuogoPresaMerce.CurrentID
    LINK_AGENTE = Me.CDAgenteTesta.KeyFieldID
    CAUSALE_TRASP = Me.txtCausaleDocumento.Text
    LINK_TIPO_ORD = Me.cboTipoOrdine.CurrentID
    LINK_PORTO = Me.cboPorto.CurrentID
    LINK_TRASP = Me.cboTrasporto.CurrentID
    LINK_VETTORE = Me.cboVettore.CurrentID
    TARGA_AUT = Me.txtTargaAutomezzo.Text
    ISTR_MITT = Me.txtIstruzioniMittente.Text
    LINK_VETTORE_SUCC = Me.cboVettoreSuccessivo.CurrentID
    LINK_ASPETTO = Me.cboAspettoEsteriore.CurrentID
    NOTE_FATT = Me.txtAnnotazioni.Text
    NOTE_INTERNE = Me.txtAnnotazioniInterna.Text
    NOTE_EVASIONE = Me.txtDescrizioneRigaDoc.Text
    LINK_CONTRATTO_COLLEGATO = Me.txtIDContratto.Value
    
    frmRiporta.Show vbModal
    

    
    Me.Caption = "INSERIMENTO NUOVO DOCUMENTO..............."
    DoEvents
    
    NewRecord
    Me.cdAnagrafica.Load LINK_ANAGRAFICA
    
    If CONF_RIPORTA_CONTRATTO = True Then
        If RIP_AGENTE_CONTR = 1 Then Me.CDAgenteTesta.Load LINK_AGENTE
        If RIP_ALTRA_DEST_CONTR = 1 Then Me.cboAltroSito.WriteOn LINK_DESTINAZIONE
        If RIP_ASPETTO_EST_CONTR = 1 Then Me.cboAspettoEsteriore.WriteOn LINK_ASPETTO
        If RIP_CAUS_CONTR = 1 Then Me.txtCausaleDocumento.Text = CAUSALE_TRASP
        If RIP_ISTRUZIONI_CONTR = 1 Then Me.txtIstruzioniMittente.Text = ISTR_MITT
        If RIP_LUOGO_MERCE_CONTR = 1 Then Me.cboLuogoPresaMerce.WriteOn LINK_LUOGO_PRESA_MERCE
        If RIP_NOTE_FATT_CONTR = 1 Then Me.txtAnnotazioni.Text = NOTE_FATT
        If RIP_NOTE_FINALI_CONTR = 1 Then Me.txtDescrizioneRigaDoc.Text = NOTE_EVASIONE
        If RIP_NOTE_INTERNE_CONTR = 1 Then Me.txtAnnotazioniInterna.Text = NOTE_INTERNE
        If RIP_PAGAMENTO_CONTR = 1 Then Me.cboPagamento.WriteOn LINK_PAGAMENTO
        If RIP_PORTO_CONTR = 1 Then Me.cboPorto.WriteOn LINK_PORTO
        If RIP_TARGA_CONTR = 1 Then Me.txtTargaAutomezzo.Text = TARGA_AUT
        If RIP_TIPO_ORD_CONTR = 1 Then Me.cboTipoOrdine.WriteOn LINK_TIPO_ORD
        If RIP_TRASPORTO_CONTR = 1 Then Me.cboTrasporto.WriteOn LINK_TRASP
        If RIP_VETTORE_CONTR = 1 Then Me.cboVettore.WriteOn LINK_VETTORE
        If RIP_VETTORE_SUCC_CONTR = 1 Then Me.cboVettoreSuccessivo.WriteOn LINK_VETTORE_SUCC
        Me.txtIDContratto.Value = LINK_CONTRATTO_COLLEGATO
    End If
    
    
    Me.Caption = "INSERIMENTO DETTAGLIO NEL NUOVO DOCUMENTO..............."
    DoEvents
    
    SCRIVE_NUOVE_RIGHE LINK_OGGETTO_DA_COPIARE
    
    If VISUALIZZA_ANDAMENTO_ORDINE = 1 Then
        GET_ANDAMENTO_ORDINE fnNotNullN(oDoc.Field("RV_POIDOrdinePadre", , sTabellaTestata))
    End If
    
    Me.Caption = Caption2Display

Exit Sub
ERR_cmdSalvaComeNuovo_Click:
    MsgBox Err.Description, vbCritical, "cmdSalvaComeNuovo_Click"

End Sub

Private Sub cmdTour_Click()
    If oDoc.IDOggetto > 0 Then
        frmTour.Show vbModal
    End If
End Sub

Private Sub Command1_Click()
If Me.txtNListaPrelievo.Value > 1 Then Exit Sub

AggiornaRiga = 1

If PermessoSalvataggio = True Then
    If bVariazioneDettaglio = False Then
        NumeroRiga = NumeroRiga + 1
    End If
        
        If oDoc.Field("IDValoriOggettoDettaglio", , sTabellaDettaglio) > 0 Then
            If oDoc.Field("RV_POIDConferimentoRighe", , sTabellaDettaglio) <> Link_RigaConferimento Then
                ArrayConfMod(contArray) = fnNotNullN(oDoc.Field("RV_POIDConferimentoRighe", , sTabellaDettaglio))
                contArray = contArray + 1
            End If
            
            SALVA_RETTIFICA
            
        End If
        
        SalvataggioRiga
        
        oDoc.PerformTable sTabellaDettaglio, False
        
        
        'Aggiorna il contenuto della listview degli articoli
        sbPopalaListaArticoli False
        'Ricalcola il documento
        sbCalcolaDocumento
        
        AggiornaRiga = 0
        
        GET_TOTALI_PEDANE
        
        'Se eravamo in presenza di un nuovo dettaglio
        If RIGA_DUPLICATA = 0 Then
            If bVariazioneDettaglio = False Then
                If lvwArticoli.ListItems.Count > 0 Then
                    'Si predispone per l'inserimento di un altro dettaglio
                    Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(lvwArticoli.ListItems.Count)
                    lvwArticoli.SelectedItem.EnsureVisible
                End If
                cmdNuovo_Click
            Else
                Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(NumeroRecordLista)
                lvwArticoli.SelectedItem.EnsureVisible
            End If
        Else
            Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(lvwArticoli.ListItems.Count)
            lvwArticoli.SelectedItem.EnsureVisible
            
            lblStatoRiga.Caption = "RIGA IN MODIFICA"
        End If
        
        RIGA_DUPLICATA = 0
        
End If
End Sub

Private Sub curScontoDocImp_LostFocus()
    
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub



Private Sub curSpeseIncasso_LostFocus()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub

Private Sub curSpeseTrasporto_LostFocus()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub






Private Sub dtData_LostFocus()
Dim LinkPagamento_Prec As Long
Dim IDPagamento As Long
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    'If Me.cboPagamento.CurrentID > 0 Then
    '    cboPagamento_Click
    'End If
    'cboSezionale_Click
    
    If Me.dtData.Text <> oDoc.Field("Doc_data", , sTabellaTestata) Then
        sbImpostaDatiDocumento
        cboSezionale_Click
        LinkPagamento_Prec = Me.cboPagamento.CurrentID
        Me.cboPagamento.WriteOn 0
        Me.cboPagamento.WriteOn LinkPagamento_Prec
        
        
        
        If Me.dtData.Value > 0 Then
            IDPagamento = GET_MODALITA_PAGAMENTO(Me.dtData.Text, Me.cdAnagrafica.KeyFieldID)
            If IDPagamento > 0 Then
                Me.cboPagamento.WriteOn IDPagamento
            End If
        End If
        
        If (Me.txtNListaPrelievo = 1) Then
            Me.txtNOrdinePadre.Value = Me.lngNumero.Value
            Me.txtDataDocPadre.Value = Me.dtData.Value
        End If
        
    End If
    
End Sub

Private Sub Form_Activate()
    'Il codice di Form_Activate deve essere eseguito soltanto la prima volta,
    'all'avvio del programma.
    '
    'La variabile m_bOnFirstTime è usata per evitare di eseguire il codice seguente
    'quando si chiude un Form di dialogo e si riattiva frmMain.
    '
    'Queste inizializzazioni non sono state effettuate nella Sub Main() per evitare di
    'rendere visibili le variabili m_DocType, m_Document e m_Changed.
    If m_bOnFirstTime = True Then

        'm_bOnFirstTime = False

        'Se il filtro di default restituisce dei record si va in modalità variazione
        'ma solo se il primo record non è bloccato altrimenti si va in modalità tabellare
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            'Il filtro ha restituito almeno un record
             
            'Controlla se il primo record su cui si dovrebbe andare in variazione è bloccato.
            If m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions) Then
                'Il primo record NON è bloccato
                'allora si effettua il blocco e si va in modalità Variazione
                
                m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
                    
                'La vista alla partenza deve essere quella del Form
                BrwMain.Visible = False

                'Imposta la modalità variazione
                SetStatus4Modality 1 'Modify
                
            Else
                'Il primo record è bloccato
                'allora si parte in modalità tabellare
                
                BrwMain.Visible = True
                
                SetStatus4Modality Browse
            End If

            RefreshDescriptions4StatusBar
        Else
            'Il filtro di default non ha restituito nessun record.
            'Si va in modalità inserimento nuovo record
            NewRecord
            
        End If
        
        
        If LINK_OGGETTO_ORDINE_PADRE_REGISTRY > 0 Then
            cmdListaPrelievo_Click
        End If
        
        m_bOnFirstTime = False
        
    End If
    
    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    End If
End Sub

Private Sub Form_Initialize()
    ActivityBox.Visible = True
    
    'Impostazione iniziale del flag
    m_bOnFirstTime = True
    
    bEnableGuiEvent = True
End Sub

Private Sub Form_Load()
    'La vista tabellare deve trovarsi sopra tutti gli altri controlli
    BrwMain.ZOrder
    
    'IMPOSTA IL CONTROLLO CHE CONTIENTE I TUTTI GLI ALTRI CONTROLLI
    Set DMTSplitBar1.dContainer = Me.PicForm2
    'IMPOSTA L'UNITA' DI MISURA DEL FORM
    DMTSplitBar1.ScaleMode = DMTSplit_Twips
    'INIZIALIZZA LA SPLIT BAR
    DMTSplitBar1.SetSplitBar Me.ScaleHeight, Me.ScaleWidth, Me.PicForm.ScaleHeight, Me.PicForm.ScaleWidth

End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If ActivityBox.Visible Then
        imgSplitter.Left = ActivityBox.Width + ActivityBox.Left
    End If
    If m_PreviewWindowHandle > 0 Then
        MoveWindow m_PreviewWindowHandle, CInt(BarMenu.ClientAreaLeft / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaTop / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaWidth / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaHeight / Screen.TwipsPerPixelX), True
    Else
        FormRecalcLayout
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bHandled As Boolean
    
    m_EatKey = False
    
    ShortCut KeyCode, Shift
    
    If KeyCode = 0 And Shift = 0 Then
        m_EatKey = True
    Else
        m_EatKey = False
    End If
    
    If KeyCode = vbKeyF8 Then
        Me.SSTab1.Tab = 0
    End If
    If KeyCode = vbKeyF9 Then
        Me.SSTab1.Tab = 1
    End If
    If KeyCode = vbKeyReturn Then
        If SSTab1.Tab = 1 Then
            cmdSalva_Click
        End If
    End If
    If KeyCode = vbKeyF2 Then
        ExecuteMenuCommand "New"
    End If
    If KeyCode = vbKeyF3 Then
        ExecuteMenuCommand "Save"
    End If
    If KeyCode = vbKeyF7 Then
        ExecuteMenuCommand "Delete"
    End If
    If KeyCode = vbKeyF5 Then
        ExecuteMenuCommand "NewSearch"
    End If
    If KeyCode = vbKeyF10 Then
    End If



    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If m_EatKey Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' ATTENZIONE
    '-------------------------------------------------------------------------------------
    ' In questo metodo qualsiasi riferimento a proprietà o metodi di un oggetto dovrebbe
    ' essere 'protetto' dal test
    '
    '                                         If obj Is Nothing then .....
    '
    ' perchè il form potrebbe essere scaricato prima che l'oggetto stesso vengana istanziato.
    '-------------------------------------------------------------------------------------

    
    Cancel = FormUnload
    
    'Distrugge il riferimento al recordset
    If Cancel = 0 Then
        Set BrwMain.Recordset = Nothing
    End If

End Sub

Private Sub Form_Terminate()

    'Distrugge tutti gli oggetti allocati e provvede ad eliminare gli eventuali blocchi
    'effettuati dalla Semaforo.
    '(Inserire in DestroyObjects il codice per la distruzione degli oggetti allocati)
    DestroyObjects

End Sub

Private Sub brwMain_DblClick()
    
    
'-------------------------------------------------------------------
    'Il documento si sincronizza con la browse
'    If BrwMain.ListIndex > 0 Then
'        m_Document.Move BrwMain.ListIndex - 1
'    End If
'-------------------------------------------------------------------
'NOTA: La versione attuale della dmtGrid effettua automaticamente il
'      Move sul documento.
'-------------------------------------------------------------------
    
    'Se si è in modalità FilterDefinition il DblClick e la pressione
    'di Invio non devono avere alcun effetto
    If BrwMain.GuiMode <> dgFilterDefinition Then
        
        ChangeView
        BrowseReposition
        
        m_Document.AbortNew
        m_Changed = False
        ActivateBarButtons BTN_SAVE, False
    
    End If
End Sub

Private Sub brwMain_KeyDown(KeyCode As Integer, Shift As Integer)

    'Alla pressione del tasto INVIO dalla modalità tabellare si passa in modalità form.
    If KeyCode = vbKeyReturn And BrwMain.GuiMode = dgNormal And BrwMain.Visible Then
        brwMain_DblClick
    End If
    
    'Viene intercettata la pressione del tasto CANC
    'e la si comunica al form.
    If KeyCode = vbKeyDelete Then
    
        'Prima di cancellare sincronizzo il documento con la selezione
        'fatta nella browse
        If BrwMain.GuiMode = dgNormal And BrwMain.ListIndex > 0 Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    
        ShortCut KeyCode, Shift
    End If
    
End Sub


'Quando si selezionano i documenti dalla modalità tabellare la Caption del form
'va costruita leggendo i valori direttamente dalla riga selezionata nella griglia
'e non da un campo del documento perchè in modalità tabellare non viene eseguito
'il Move sul documento.
Private Sub BrwMain_Reposition(ByVal AllColumns As DmtGridCtl.dgColumns)
    If Not (m_Document.EOF = True Or m_Document.BOF = True) Then
        'Monta la caption del form principale
        Me.Caption = Caption2Display(True)
    End If
End Sub


Private Sub BrwMain_OnChangeGuiMode()
    'Se si cambia modalità tramite il menù presente nel controllo
    'dmtGrid occorre effettuare delle impostazioni preliminari nella UserInterface
    
    If bEnableGuiEvent Then
    
        'Modalità FilterDefinition
        If BrwMain.GuiMode = dgFilterDefinition Then
            'Annulla una eventuale operazione di inserimento di un nuovo record
            If m_Document.TableNew Then
                m_Document.AbortNew
            End If
            
            'Impostazioni per la modalità Ricerca
            SetStatus4Modality 2 'Find
        End If
        
        'Modalità tabellare
        If BrwMain.GuiMode = dgNormal Then
            'Se si è premuto il pulsante "Visualizzazione tabellare" dalla browse
            'in modalità FilterDefinition e con il recordset vuoto, non si deve andare in
            'modalità tabellare (browse vuota) ma si deve restare in modalità ricerca.
            If (m_Document.EOF = True And m_Document.BOF = True) Then
                BrwMain.GuiMode = dgFilterDefinition
            Else
                'Impostazioni per la modalità tabellare
                SetStatus4Modality 3 'Browse
            End If
        End If
    
    End If
End Sub

'Scatenato prima che venga visualizzata la Toolbar della DmtGrid
Private Sub BrwMain_BeforeShowActions()
    
    'Quando si è in modalità FilterDefinition si può andare in
    'modalità tabellare solo se il documento contiene almeno un record.
    If BrwMain.GuiMode = dgFilterDefinition Then
        'Abilita/disabilita il pulsante Modalità Tabellare della dmtGrid
        BrwMain.Actions("TableMode").Enabled = (m_Document.EOF <> True And m_Document.BOF <> True)
    End If
End Sub

'Scatenato quando dalla Browse ( in modalità FilterDefinition ) si clicca su esegui ricerca.
Private Sub BrwMain_OnApplyFilter(ByVal Filter As String)
    ExecuteSearch
End Sub



Private Sub lblDocument_Click(Index As Integer)
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtOleDbLib.adoResultset

If Index = 36 Then
If Me.txtDescrizioneImballo.Text = "" Then Exit Sub
    If Me.CDImballo.Code = "" Then
        'Crea un'istanza dell'oggetto Find
        Set oSearch = New dmtFind.Find
        
        'Assegna la connessione aperta
        oSearch.Database = Cn
        
        'La Caption della finestra di ricerca
        oSearch.Caption = "Imballi"
        

        oSearch.AddDisplayField "Articolo", "Articolo", 1
        oSearch.AddDisplayField "Codice Articolo", "CodiceArticolo", 1
        
        
        oSearch.Filters.Add "Articolo", txtDescrizioneImballo.Text
        oSearch.Start = Me.txtDescrizioneImballo.Text
    
            sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
            sSQL = sSQL & "FROM Articolo "
            sSQL = sSQL & "WHERE ((IDTipoProdotto = " & Link_TipoImballo & ") "
            'sSQL = sSQL & "AND (GestioneLotti=" & fnNormBoolean(1) & ") "
            sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & "))"
            
        
        
        
        oSearch.SQL = fnAnsi2Jet(sSQL)
        
        
       
        
        Set oRes = oSearch.Exec
        
        
        If Not oRes.EOF Then
            Me.CDImballo.Load fnNotNullN(oRes!IDArticolo)
        End If
        
        Set oRes = Nothing
        Set oSearch = Nothing
        
    End If
        

End If

If Index = 24 Then
If Me.txtDescrizioneArticolo.Text = "" Then Exit Sub
    If Me.CDArticolo.Code = "" Then
        Set oSearch = New dmtFind.Find
        
        oSearch.Database = Cn
        
        oSearch.Caption = "Articoli"
        
        
        oSearch.AddDisplayField "Articolo", "Articolo", 1
        oSearch.AddDisplayField "Codice Articolo", "CodiceArticolo", 1

        
        oSearch.Filters.Add "Articolo", Me.txtDescrizioneArticolo.Text
        
        oSearch.Start = Me.txtDescrizioneArticolo.Text
        
            sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo, IDUnitaDiMisuraAcquisto "
            sSQL = sSQL & "FROM Articolo "
            sSQL = sSQL & "WHERE (((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL)) "
            sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & "))"
            
        oSearch.SQL = fnAnsi2Jet(sSQL)
        
        
       
        Set oRes = oSearch.Exec
        
        
        If Not oRes.EOF Then
            Me.CDArticolo.Load fnNotNullN(oRes!IDArticolo)
        End If
        Set oRes = Nothing
        Set oSearch = Nothing
        
    End If
        

End If

End Sub


Private Sub lngNumero_Change()

    sbImpostaDatiDocumento
End Sub

Private Sub lngNumero_LostFocus()
    If (Me.txtNListaPrelievo.Value) = 1 Then
        Me.txtNOrdinePadre.Value = Me.lngNumero.Value
    End If
End Sub

Private Sub lngScontoDocPer_LostFocus()
    
    sbImpostaDatiDocumento
    sbCalcolaDocumento
    
End Sub

Private Sub lvwArticoli_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Annulla l'inserimento di un nuovo dettaglio
    
    bVariazioneDettaglio = True
    Var_LostFocus_Colli = 0
    'Va in variazione della riga di dettaglio selezionata
    
    bLoadingRiga = True
    
    sbVariazioneRiga
    
    bLoadingRiga = False
    'Aggiorna il contenuto della listview degli articoli
    CalcolaTotaleRiga
    
    If VISUALIZZA_ANDAMENTO_ORDINE = 1 Then
'        Me.lvwArticoli.Height = 4215
'        Me.lvwArticoli.Top = 3840
        GET_ANDAMENTO_ORDINE_LAVORAZIONE oDoc.IDOggetto, Me.CDArticolo.KeyFieldID, Me.txtRaggrRigaOrdine.Text
    End If
'    Else
'        Me.lvwArticoli.Height = 4455
'        Me.lvwArticoli.Top = 3600
'    End If
    
    DISEGNA_FORM
    
    Me.txtScontoImpListino.Value = GET_SCONTO_IMPORTO_LISTINO
    Me.txtCollegamentoRigaOrdine.Text = GET_DESCRIZIONE_COLLEGAMENTO(NumeroRecordPerModifica)
    NumeroRecordLista = Me.lvwArticoli.SelectedItem.Index
    
    lblStatoRiga.Caption = "RIGA IN MODIFICA"
    RIGA_DUPLICATA = 0
End Sub

Private Sub m_App_OnRun(ByVal Proc As Process)
    Dim Parameter As DMTRunAppLib.Parameter

    On Error GoTo ErrorHandler
    
    Set m_Process = Proc
    Set m_DocType = m_Process.IDocType
    
    '.................................................................................................................................
    '.................................................................................................................................
    'Gestione preliminare della Semaforo per il controllo dei conflitti di multiutenza
    
    
    'Inizializza la Semaforo
    InitSemaphore
    
    ' Verifica se l'applicazione corrente è bloccata da altri gestori.
    ' (Il controllo avviene sul Tipo Oggetto correntemente trattato.)
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, SemAllObjects, SemAllActions) Then
        '-------------------------------------------------------------
        'Il programma è bloccato da un'altra manutenzione in esecuzione.
        '-------------------------------------------------------------
        
        'Scarica il form
        Unload Me
       
        'Prima di terminare il programma è bene distruggere tutti gli oggetti allocati
        DestroyObjects
       
        'Termina il programma
        End
    End If
    
    '----------------------------------------------------
    'Il programma non è bloccato e prosegue normalmente.
    '----------------------------------------------------
    
    'Ripulisce la tabella semaforo.
    'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
    SemaphoreUnlock
    
    'Imposta gli eventuali blocchi (semaforo) su altre manutenzioni.
    SemaphoreLock
    '.................................................................................................................................
    '.................................................................................................................................
    
    
    Select Case Proc.Name
        '*
        'Inserire il codice per la gestione del processo
        '*
        Case "Manutenzione"
        '   For Each Parameter In Proc.Parameters
        '       Select Case Parameter.Name
        '       *
        '       Inserire il codice per la gestione del parametro
        '       *
        '       Case ParameterName??????
        '       End Select
        '   next
           Start 'di solito
    
    Case Else
    
'''''''        Dim ErrorMsg As String
'''''''
'''''''        ErrorMsg = "No processes to execute" & vbCrLf
'''''''        ErrorMsg = ErrorMsg & "This application is able to execute these processes:" & vbCrLf
'''''''        '*
'''''''        'Inserire i processi che l'applicazione sa eseguire
'''''''        '*
'''''''        Err.Raise ERR_NO_PROCESSES, , ErrorMsg
        
    End Select
    Exit Sub
ErrorHandler:
    SemaphoreUnlock
    ShowErrorLog
End Sub

Private Sub m_Document_OnReposition()
    
    'oDoc.ClearValues
    If Not m_Document.TableNew Then
        'Se EOF = true o BOF = true vuol dire che si è andati oltre l'ultimo o
        'prima del primo record. In tal caso non si deve fare il refresh dei
        'controlli del form.
        If Not (m_Document.EOF Or m_Document.BOF) Then
            BrowseReposition
            
        End If
    Else
        'Nel caso di inserimento nuovo record ripulisce i campi del form
       
        ClearFormFields
    End If
    
    'Set Me.GrigliaCommissioni.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink.TableName).Data
    
    'CONTROLLA_BLOCCHI_INSERIMENTI
    
    If oDoc.IDOggetto = 0 Then
        NuovoDocumento = 0
        Me.cboSezionale.Enabled = True
    Else
        NuovoDocumento = 1
        fnGrigliaCommissioni
        Me.cboSezionale.Enabled = False
    End If
    
    If VISUALIZZA_ANDAMENTO_ORDINE = 1 Then
        GET_ANDAMENTO_ORDINE fnNotNullN(oDoc.Field("RV_POIDOrdinePadre", , sTabellaTestata))
    End If
    
    'fnEliminaDatiTemporanei
    ControllaNumeroRiga
    Me.SSTab1.Tab = 0
    'chkCessione_Click
    
    GET_TOTALI_PEDANE
    
    sbPopalaListaArticoli True
    sbPopalaListaScadenze
    sbPopalaListaIva
    sbPopalaListaCommissioni
    RiepilogoTotaliDocumento
    GET_DATI_TOUR oDoc.IDOggetto
    
    ABILITA_CONTROLLI
    
    CREA_RECORDSET_RETTIFICA
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 07/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: AutoLostFocus
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Forza un LostFocus del controllo attivo ed attende la gestione di eventuali eventi associati.
'                  Alla fine ripristina il fuoco sul controllo iniziale.
'                  Usata quando si clicca sulla toolbar e quando si utilizza l'acceleratore per il salvataggio SHIFT + F12
'                  (in tal caso infatti non viene scatenato l'evento BarMenu_Click)
'
'**/
Private Sub AutoLostFocus()
    Dim Ctr As Control

    
    'Se si è in modalità FilterDefinition non si deve spostare il fuoco
    'altrimenti Taglia, Copia e Incolla (dalla toolbar) non possono funzionare
    If BrwMain.GuiMode <> dgFilterDefinition And Not Me.ActiveControl Is Nothing Then
    
        'Memorizza il controllo che ha il fuoco
        Set Ctr = Me.ActiveControl
    
        'Forza il lost focus del controllo attivo
        Globali.SetFocus PicForm.hwnd
        
        'Vengono gestiti gli eventi LostFocus (se previsti)
        DoEvents
        
        'Ripristina il fuoco sul controllo.
        Ctr.SetFocus
        
    End If

End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 11/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: InitSemaphore
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Inizializzazione del semaforo per la gestione
'                  dei conflitti in caso di multiutenza

'
'**/
Private Sub InitSemaphore()

    Set m_Semaphore = New Semaforo.dmtSemaphore
    Set m_Semaphore.Database = m_App.Database.Connection
    Set m_Semaphore.objRes = gResource
    
    m_Semaphore.IDUser = m_App.IDUser
    m_Semaphore.IDBranch = m_App.Branch
    m_Semaphore.IDFunction = m_App.FunctionID
    
End Sub


'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreLock
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                 ////////////////////////////////////////////////////////////////////////
'                     Impostare qui gli eventuali blocchi sulle altre manutenzioni
'                 ////////////////////////////////////////////////////////////////////////
'**/
Private Sub SemaphoreLock()
    If Not m_Semaphore Is Nothing Then
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        m_Semaphore.SetObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions

    End If
End Sub

'**+
'Autore: Diamante s.p.a
'Data creazione: 12/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: SemaphoreUnlock
'
'Parametri:
'
'Valori di ritorno:

'Funzionalità:
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'                     Sbloccare qui le altre manutenzioni (bloccate precedentemente in SemaphoreLock)
'                 //////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Sub SemaphoreUnlock()
    If Not m_Semaphore Is Nothing Then
    
        'Ripulisce la tabella semaforo per quanto riguarda il Tipo Oggetto e l'utente correnti
        m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        
        
        
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        'Personalizzare, se necessario, le righe sottostanti
        '/////////////////////////////////////////////////////////////////////////////////////////////////
        
        'Sblocca le manutenzioni bloccate precedentemente
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_XXX, SemAllObjects, SemAllActions
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_YYY, SemAllObjects, SemAllActions
'        m_Semaphore.ClearObjectAction TO_TIPO_OGGETTO_ZZZ, SemAllObjects, SemAllActions
    
    End If
End Sub




'**+
'Autore: Diamante s.p.a
'Data creazione: 11/09/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: DestroyObjects
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'                  ////////////////////////////////////////////////////////////////////////////////////////////////////
'                  /         Inserire qui il codice per distruggere (prima che venga terminato il programma)     /
'                  /         tutti gli oggetti allocati                                                                              /
'                  ////////////////////////////////////////////////////////////////////////////////////////////////////
'
'**/
Private Sub DestroyObjects()
    
    'Sblocca gli eventuali gestori bloccati da questa manutenzione
    SemaphoreUnlock

    Set m_Report = Nothing
    Set m_ActiveFilter = Nothing
    Set m_Document = Nothing
    Set m_Process = Nothing
    Set m_App = Nothing
    Set m_Semaphore = Nothing
End Sub



'**+
'Autore: Diamante s.p.a
'Data creazione: 15/10/03
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: oDocChangeNotify_ChangeValue
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
' Intercetta la modifica dei valori di ogni singolo campo di tutte
' le tabelle presenti nel documento ed effettua eventuali controlli
' sul valore e/o ne riporta semplicemente il valore a video
'
'**/
Private Sub oDocChangeNotify_ChangeValue(ByVal Table As DmtTables.cTable, ByVal Field As DmtTables.cField, ByVal Value As Variant)
    'Normalizza il valore a seconda del tipo di campo
    Select Case Field.FieldType
        Case adTypeBIGINT, adTypeBIT, adTypeDECIMAL, adTypeDOUBLE, adTypeFLOAT, adTypeINTEGER, adTypeNUMERIC, adTypeREAL, adTypeSMALLINT, adTypeTINYINT
            Value = fnNotNullN(Value)
        Case Else
            Value = fnNotNull(Value)
    End Select
    
    Select Case Field.Name
        
        Case "Doc_data"
            oDoc.DataEmissione = Value
            dtData.Text = Value
        
            'Quando cambia la data imposta l'Esercizio contabile di riferimento
            fnSetEsercizio Value
            
            If fnNotNullN((oDoc.Field("Doc_numero", , sTabellaTestata)) = 0) Then
                oDoc.Field "Doc_numero", fnDocumentNumber(), sTabellaTestata
            End If
            
            'ed il numero documento
        Case "Doc_numero"
            oDoc.Numero = Value
            lngNumero.Value = Value
        Case "Doc_prezzi_lordo_IVA"
            chkLordoIVA.Value = Abs(CLng(Val(Value)))
            'Quando cambia il flag Prezzi lordo Iva ricalcola il documento
            sbCalcolaDocumento
        'Case "Nom_IVA_in_sospensione"
        '    Me.chkSpospensioneIva.Value = Abs(CLng(Val(Value)))
            'Quando cambia il flag Prezzi lordo Iva ricalcola il documento
            sbCalcolaDocumento
        Case "Link_Doc_sezionale"
            cboSezionale.WriteOn Val(Value)
            oDoc.IDSezionale = Me.cboSezionale.CurrentID
            oDoc.Field "Link_Doc_sezionale", oDoc.IDSezionale, sTabellaTestata

        Case "Link_Doc_listino"
            cboListino.WriteOn Val(Value)
        Case "Link_Nom_IVA"
            LINK_CLIENTE_IVA = Val(Value)
            Me.cboIvaCliente.WriteOn Val(Value)
        Case "Link_Doc_pagamento"
            cboPagamento.WriteOn Val(Value)
            'Quando cambia la modalità di pagamento ricalcola il documento
            sbCalcolaDocumento
        Case "Doc_causale_trasporto"
            Me.txtCausaleDocumento.Text = Value
        Case "Link_Nom_anagrafica"
            cdAnagrafica.Load Val(Value)
            TIPO_SCONTO_CLIENTE = GET_TIPO_SCONTO(Val(Value))
        Case "Nom_partita_IVA"
            txtPartitaIva.Text = Value
        Case "Nom_indirizzo"
            txtIndirizzo.Text = Value
        Case "Nom_cap"
            txtCAP.Text = Value
        Case "Nom_comune"
            txtComune.Text = Value
        Case "Nom_provincia"
            txtProvincia.Text = Value
        Case "Spe_incasso_neutro"
            curSpeseIncasso.Value = Value
            'Quando cambiano le spese incasso ricalcola il documento
            sbCalcolaDocumento
        Case "Spe_trasporto_neutro"
            curSpeseTrasporto.Value = Value
            'Quando cambiano le spese di trasporto ricalcola il documento
            sbCalcolaDocumento
        Case "Sco_percentuale_fine_documento"
            lngScontoDocPer.Value = Value
            'Quando cambia lo sconto in percentuale di fine documento ricalcola il documento
            sbCalcolaDocumento
        Case "Sco_ad_importo_fine_documento"
            curScontoDocImp.Value = Value
            'Quando cambia lo sconto ad importo di fine documento ricalcola il documento
            sbCalcolaDocumento
        Case "Tot_imponibile_corr"
            curTotImponibile.Value = Value
            'Me.curTotImponibile_Corpo.Value = Value
        Case "Tot_imposta_corr"
            curTotImposta.Value = Value
            'Me.curTotImposta_Corpo.Value = Value
        Case "Tot_documento_corr"
            curTotDocumento.Value = Value
            'Me.curTotDocumento_Corpo.Value = Value
             oDoc.Scadenze.ParametersChanged = True
        Case "Tot_arrotondamenti"
            curTotArrotondamenti.Value = Value
            'Me.curTotArrotondamenti_Corpo.Value = Value
        Case "Tot_netto_a_pagare_corr"
            curNettoAPagare.Value = Value
            'Me.curNettoAPegare_Corpo.Value = Value
        Case "Link_Nom_porto"
            Me.cboPorto.WriteOn Value
        Case "Link_Doc_spedizione"
            Me.cboTrasporto.WriteOn Value
        Case "Link_Vet_vettore"
            Me.cboVettore.WriteOn Value
        'Case "Doc_data_inizio_trasporto"
        '    Me.txtDataTrasporto.Text = Value
        'Case "Doc_ora_inizio_trasporto"
        '    Me.txtOraTrasporto.Text = Value
        Case "Tot_numero_colli"
            Me.txtColliTotali.Value = Value
        Case "Tot_peso"
            Me.txtPesoTotale.Value = Value
        Case "Link_Nom_ult_sito"
            Me.cboAltroSito.WriteOn Value
        Case "Nom_ult_sito_cap"
                Me.txtCapAltroSito.Text = Value
        Case "Nom_ult_sito_indirizzo"
                Me.txtIndirizzoAltroSito.Text = Value
        Case "Nom_ult_sito_comune"
                Me.txtComuneAltroSito.Text = Value
        Case "Nom_ult_sito_provincia"
                Me.txtProvinciaAltroSito.Text = Value
        Case "Nom_raggruppamento_bolle"
            Me.chkRaggruppBolle.Value = Abs(CLng(Val(Value)))
        Case "Nom_raggruppamento_scadenze"
            Me.chkRaggruppaScadenze.Value = Abs(CLng(Val(Value)))

        Case "Link_Doc_aspetto_esteriore"
            Me.cboAspettoEsteriore.WriteOn Val(Value)
        Case "Doc_annotazioni_variazio"
            Me.txtAnnotazioni.Text = Value

        'Case "RV_POIstruzioniMittente"
        '    Me.txtIstruzioniMittente.Text = Value
        'Case "RV_POTargaAutomezzo"
        '    Me.txtTargaAutomezzo.Text = Value
        
        Case "Val_data_cambio"
            Me.txtDataCambio.Text = Value
            'sbCalcolaDocumento
        Case "Val_valore_cambio"
            Me.txtValoreCambioValuta.Value = Value
            'sbCalcolaDocumento
        Case "Link_Val_valuta"
            Me.cboValuta.WriteOn Val(Value)
            sbCalcolaDocumento
        Case "Link_Val_cambio"
            Me.cboCambioValuta.WriteOn Val(Value)
            'sbCalcolaDocumento
        Case "Tot_netto_a_pagare_naz"
            curNettoAPagare_naz.Value = Value
            'curNettoAPegare_Corpo_naz.Value = Value
        Case "RV_POIDLuogoPresaMerce"
            Me.cboLuogoPresaMerce.WriteOn Value
        Case "RV_POIDTrasportatoreSuccessivo"
            Me.cboVettoreSuccessivo.WriteOn Value
        Case "Doc_data_presso_nom"
            Me.txtDataOrdineCliente.Text = Value
        Case "Doc_numero_presso_nom"
            Me.txtNumeroOrdineCliente.Text = Value
        Case "RV_PODataArrivoMerce"
            Me.txtDataTrasporto.Text = Value
        Case "RV_POOraArrivoMerce"
            Me.txtOraTrasporto.Text = Value
        Case "Doc_data_prevista_evasione"
            Me.txtDataPartenza.Text = Value
        Case "Doc_ordine_chiuso"
            chkChiuso.Value = Abs(CLng(Val(Value)))
        Case "Link_Doc_listino_base"
            Me.cboListinoAzienda.WriteOn fnNotNullN(Value)
        Case "RV_POAnnotazioniInterna"
            Me.txtAnnotazioniInterna.Text = Value
        Case "RV_PODescrizioneCorpoDocEv"
            Me.txtDescrizioneRigaDoc.Text = Value
        Case "RV_POIDTipoOrdine"
            Me.cboTipoOrdine.WriteOn Value
        Case "RV_PODataArrivoMerceLuogo"
            Me.txtDataArrivoLuogo.Text = Value
        Case "RV_POOraArrivoMerceLuogo"
            Me.txtOraArrivoLuogo.Text = Value
        Case "Link_Nom_lettera_intento"
            Me.txtIDLetteraIntento.Value = Value
        Case "Link_Doc_magazzino"
            Me.cboMagazzino.WriteOn Value
        Case "Link_Doc_contratto_bancario_az"
            Me.cboBancaAzienda.WriteOn Value
        Case "Link_Nom_contratto_bancario"
            Me.cboBancaCliente.WriteOn Value
        Case "Link_Nom_accordi_commerciali"
            Me.cboAccordoCommerciale.WriteOn Value
        Case "Link_Doc_agente"
            Me.CDAgenteTesta.Load fnNotNullN(Value)
        Case "RV_POIstruzioniMittente"
            Me.txtIstruzioniMittente.Text = fnNotNull(Value)
        Case "RV_POTargaAutomezzo"
            Me.txtTargaAutomezzo.Text = fnNotNull(Value)
            
        Case "RV_PONumeroOrdinePadre"
            Me.txtNOrdinePadre.Value = fnNotNullN(Value)
        Case "RV_PODataOrdinePadre"
            Me.txtDataDocPadre.Value = fnNotNullN(Value)
        Case "RV_PONumeroListaPrelievo"
            Me.txtNListaPrelievo.Value = fnNotNullN(Value)
        Case "RV_POOrdineCompletato"
            Me.chkOrdineCompletato.Value = Value
        Case "RV_PONumeroPedanePrelievo"
            Me.txtNPedaneTesta.Value = Value
        Case "RV_POIDAnagraficaDestinazione"
            If Value > 0 Then
                Me.ACSAnaDest.IDAnagrafica = 0
                Me.ACSAnaDest.Code = ""
                Me.ACSAnaDest.Description = ""
                Me.ACSAnaDest.SecondDescription = ""
                Me.ACSAnaDest.sbLoadCFByIDAnagrafica 0, Value
            Else
                Me.ACSAnaDest.IDAnagrafica = 0
                Me.ACSAnaDest.Code = ""
                Me.ACSAnaDest.Description = ""
                Me.ACSAnaDest.SecondDescription = ""
            End If
        Case "RV_POIDOggettoContratto"
            Me.txtIDContratto.Value = Value
        Case "Link_Nom_raggrup_fatturato"
            Me.cboRaggrFatturato.WriteOn fnNotNullN(Value)
        Case "RV_POConfPresaVisContratto"
            Me.chkConfDaContratto.Value = Abs(fnNotNullN(Value))
        Case "Doc_causale_documento"
            Me.txtCausaleDocumentoEF.Text = Value
        Case "RV_POFatturaProforma"
            Me.chkStampaFattProForma.Value = Abs(fnNotNullN(Value))
    End Select

    'Solo per la modalita' form attiva il salva
    If Not (BrwMain.Visible) Then Change
End Sub

'Funzionalita': Imposta l'IDEsercizio in base alla data documento
Function fnSetEsercizio(Optional ByVal sDate As String) As Long
    Dim Rw As Rowset
        
    'Se non viene passata la data come parametro prendiamo la data del documento
    If Len(Trim(sDate)) = 0 Then
        sDate = oDoc.Field("Doc_data", , sTabellaTestata)
    End If
    If sDate <> "" Then
        'Effettua una Query sulla tabella Esercizio cercando
        'l'eserzio che rientra nella data specificata
        Set Rw = TheApp.Database.CreateRowset
        Rw.Columns.Add "IDEsercizio"
        Rw.Tables.Add "Esercizio"
        Rw.Where = "DataInizio <= " & fnNormDate(sDate)
        Rw.Where = Rw.Where & " AND DataFine >= " & fnNormDate(sDate)
        Rw.Where = Rw.Where & " AND IDAzienda = " & TheApp.IDFirm
        
        Rw.Refresh
        If Not Rw.EOF Then
            'Imposta la proprietà IDEserzio dell'oggetto cDocument
            oDoc.IDEsercizio = Rw.Columns("IDEsercizio").Value
        Else
            oDoc.IDEsercizio = 0
        End If
        Set Rw = Nothing
    End If
End Function

'Funzionalita': Imposta l'IDSezionale in base alla filiale ed al tipo di documento
Function fnSetSezionale() As Long
    Dim Rw As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    
    'Effettua una Query sulla tabella Sezionale e RegistroIvaPerTipoOggetto cercando
    'il sezionale di default per il tipo di documento (TipoOggetto) e la filiale correnti
        
    sSQL = "SELECT Sezionale.IDSezionale, Sezionale.Sezionale, Sezionale.Prefisso "
    sSQL = sSQL & "FROM Sezionale INNER JOIN "
    sSQL = sSQL & "DefaultFilialePerTipoOggetto ON Sezionale.IDFiliale = DefaultFilialePerTipoOggetto.IDFiliale AND "
    sSQL = sSQL & "Sezionale.IDSezionale = DefaultFilialePerTipoOggetto.IDSezionale "
    sSQL = sSQL & "WHERE (DefaultFilialePerTipoOggetto.IDFiliale = " & oDoc.IDFiliale & ") AND (DefaultFilialePerTipoOggetto.IDTipoOggetto = " & oDoc.IDTipoOggetto & ")"
    Set Rw = Cn.OpenResultset(sSQL)
    If Not Rw.EOF Then
        'Imposta la proprietà IDSezionale dell'oggetto cDocument
        oDoc.IDSezionale = Rw.adoColumns("IDSezionale").Value
        'Imposta il valore dell'ID e del previsso nei campi della tabella di testata
        oDoc.Field "Link_Doc_sezionale", oDoc.IDSezionale, sTabellaTestata
        oDoc.Field "Doc_prefisso", Rw.adoColumns("Prefisso").Value, sTabellaTestata
        Me.cboSezionale.WriteOn oDoc.IDSezionale
    Else
        oDoc.IDSezionale = 0
        oDoc.Field "Link_Doc_sezionale", oDoc.IDSezionale, sTabellaTestata
        oDoc.Field "Doc_prefisso", Null, sTabellaTestata
        Me.cboSezionale.WriteOn oDoc.IDSezionale
    End If
    Set Rw = Nothing
    
    
End Function

'Funzionalita': Determina il numero documento.
Function fnDocumentNumber(Optional ByVal sDate As String) As Long
    Dim Rw As Rowset
    Dim lIDPeriodoIva As Long
    
    'Se non viene passata la data come parametro prendiamo la data del documento
    If Len(Trim(sDate)) = 0 Then
        sDate = oDoc.Field("Doc_data", , sTabellaTestata)
    End If
    If IsNumeric(sDate) Then
        If sDate = 0 Then
            sDate = Date
        End If
    End If
    If Val(sDate) = 0 Then
        sDate = Date
    End If
    If sDate = "" Then
        sDate = Date
    End If
    'If sDate <> "" Then
        'Effettua una Query sulla tabella PeriodoIva cercando
        'il periodo iva di riferimento per l'anno a cui fa riferimento la data
        Set Rw = TheApp.Database.CreateRowset
        Rw.Columns.Add "IDPeriodoIva"
        Rw.Tables.Add "PeriodoIva"
        Rw.Where = "IDAzienda = " & oDoc.IDAzienda
        Rw.Where = Rw.Where & " AND Anno = " & Year(sDate)
        Rw.Where = Rw.Where & " AND VirtualDelete = 0"
        Rw.Refresh
        If Rw.EOF = False Then
            'Legge l'ID del periodo IVA trovato
            lIDPeriodoIva = fnNotNullN(Rw.Columns("IDPeriodoIva").Value)
        Else
            MsgBox "Il periodo IVA a cui fa riferimento la data del documento non è esistente", vbCritical, "PERIODO I.V.A."
            lIDPeriodoIva = fnNotNullN(Rw.Columns("IDPeriodoIva").Value)
        End If
        Set Rw = Nothing
    'Else
    '    lIDPeriodoIva = 0
    'End If
    
    'Effettua una Query sulla tabella ProgressivoDisponibile cercando
    'il primo numero disponibile in base al periodo iva e al sezionale
    Set Rw = TheApp.Database.CreateRowset
    Rw.Columns.Add "ProgressivoDisponibile"
    Rw.Tables.Add "ProgressivoSezionale"
    Rw.Where = "IDPeriodoIva = " & lIDPeriodoIva
    Rw.Where = Rw.Where & " AND IDTipoModulo = " & oDoc.DBDefaults.IDTipoModulo
    Rw.Where = Rw.Where & " AND IDSezionale = " & oDoc.IDSezionale
    Rw.Where = Rw.Where & " AND VirtualDelete = 0"
    Rw.Refresh
    If Not Rw.EOF Then
        'Restituisce il progressivo disponible
        fnDocumentNumber = IIf(IsNull(Rw.Columns("ProgressivoDisponibile").Value), 1, Rw.Columns("ProgressivoDisponibile").Value)
    Else
        'Se il progressivo non esiste vuol dire che non è mai stato emesso
        'un documento per quel periodo IVA e sezionale
        fnDocumentNumber = 1
    End If
    Set Rw = Nothing
End Function

'Popola la listview degli artcoli in base ai dettagli presenti nel documento
Private Sub sbPopalaListaArticoli(RefreshDocumento As Boolean)
On Error GoTo ERR_sbPopalaListaArticoli
    Dim lRow As Long
    Dim oItem As MSComctlLib.ListItem
    Dim sSQL As String
    
    'Pulisce la listview
    lvwArticoli.ListItems.Clear
    
    With oDoc.Tables(sTabellaDettaglio)
        'Cicla per tutte le righe di dettaglio presenti nel documento
        If .NumRetails = 0 Then Exit Sub
        
        For lRow = 1 To .NumRetails
            Set oItem = lvwArticoli.ListItems.Add

                oItem.Text = fnNotNullN(.Fields("RV_POLinkRiga").Values(lRow).Value)
                oItem.SubItems(1) = fnNotNullN(.Fields("RV_POTipoRiga").Values(lRow).Value)
                oItem.SubItems(2) = GET_NUMERO_RETTIFICHE(oDoc.IDOggetto, oDoc.IDTipoOggetto, fnNotNullN(fnNotNullN(.Fields("RV_POLinkRiga").Values(lRow).Value)))
                oItem.SubItems(3) = fnNotNullN(fnNotNullN(.Fields("Link_Art_articolo").Values(lRow).Value))
                oItem.SubItems(4) = fnNotNull(.Fields("Art_Codice").Values(lRow).Value)
                oItem.SubItems(5) = fnNotNull(.Fields("Art_descrizione").Values(lRow).Value)
                oItem.SubItems(6) = FormatNumber(fnNotNullN(.Fields("Art_quantita_totale").Values(lRow).Value), 2)
                oItem.SubItems(7) = FormatNumber(fnNotNullN(.Fields("Art_prezzo_unitario_netto_IVA").Values(lRow).Value), 5)
                oItem.SubItems(8) = FormatNumber(fnNotNullN(.Fields("Art_sco_in_percentuale_1").Values(lRow).Value), 2)
                oItem.SubItems(9) = FormatNumber(fnNotNullN(.Fields("Art_sco_in_percentuale_2").Values(lRow).Value), 2)
                If fnNotNullN(.Fields("RV_POTipoRiga").Values(lRow).Value) = 1 Then
                    oItem.SubItems(10) = FormatNumber(fnNotNullN(.Fields("Art_pre_uni_net_sco_net_IVA").Values(lRow).Value), 4)
                Else
                    oItem.SubItems(10) = FormatNumber(fnNotNullN(.Fields("Art_prezzo_unitario_netto_IVA").Values(lRow).Value), 5)
                End If
                oItem.SubItems(11) = fnNotNullN(.Fields("RV_POIDImballoPrimario").Values(lRow).Value)
                oItem.SubItems(12) = fnNotNull(.Fields("RV_POCodiceImballoPrimario").Values(lRow).Value)
                oItem.SubItems(13) = fnNotNull(.Fields("RV_PODescrizioneImballoPrimario").Values(lRow).Value)
                
                oItem.SubItems(14) = fnNotNullN(.Fields("RV_POIDArticoloPedana").Values(lRow).Value)
                oItem.SubItems(15) = fnNotNull(.Fields("RV_POCodiceArticoloPedana").Values(lRow).Value)
                oItem.SubItems(16) = fnNotNull(.Fields("RV_PODescrizioneArticoloPedana").Values(lRow).Value)
                oItem.SubItems(17) = fnNotNullN(.Fields("RV_POQuantitaPedana").Values(lRow).Value)
                oItem.SubItems(18) = fnNotNull(.Fields("RV_POIDTipoUMOrdine").Values(lRow).Value)
                oItem.SubItems(19) = fnNotNull(.Fields("RV_POTipoUMOrdine").Values(lRow).Value)

                oItem.SubItems(20) = fnNotNullN(.Fields("Art_numero_colli").Values(lRow).Value)
                oItem.SubItems(21) = FormatNumber(fnNotNullN(.Fields("Art_quantita_pezzi").Values(lRow).Value), 2)
                oItem.SubItems(22) = FormatNumber(fnNotNullN(.Fields("Art_peso").Values(lRow).Value), 2)
                oItem.SubItems(23) = FormatNumber(fnNotNullN(.Fields("Art_tara").Values(lRow).Value), 2)
                oItem.SubItems(24) = fnNotNullN(.Fields("Art_aliquota_IVA").Values(lRow).Value)
                oItem.SubItems(25) = FormatNumber(fnNotNullN(.Fields("Art_importo_totale_netto_IVA").Values(lRow).Value), 2)
                oItem.SubItems(26) = FormatNumber(fnNotNullN(.Fields("Art_importo_totale_lordo_IVA").Values(lRow).Value), 2)
                oItem.SubItems(27) = fnNotNull(.Fields("RV_PONotaRigaOrdRaggr").Values(lRow).Value)
                oItem.SubItems(28) = GET_DESCRIZIONE_TABELLA(2, fnNotNullN(.Fields("RV_POIDTipoCategoria").Values(lRow).Value))
                oItem.SubItems(29) = GET_DESCRIZIONE_TABELLA(1, fnNotNullN(.Fields("RV_POIDCalibro").Values(lRow).Value))
                oItem.SubItems(30) = GET_DESCRIZIONE_TABELLA(3, fnNotNullN(.Fields("RV_POIDTipoLavorazione").Values(lRow).Value))

            'End If
        DoEvents
        Next lRow
        
        'Imposta la riga selezionata in base a quella che era già precedentemente selezionata
        If .ActiveRow > 0 Then
            If .ActiveRow <= lvwArticoli.ListItems.Count Then
                Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(.ActiveRow)
                lvwArticoli.SelectedItem.EnsureVisible
            ElseIf lvwArticoli.ListItems.Count > 0 Then
                Set lvwArticoli.SelectedItem = lvwArticoli.ListItems(lvwArticoli.ListItems.Count)
                lvwArticoli.SelectedItem.EnsureVisible
                sbVariazioneRiga
            End If
        End If
    End With
Exit Sub
ERR_sbPopalaListaArticoli:
MsgBox Err.Description, vbCritical, TheApp.FunctionName

End Sub

'Popola la listview delle scadenze in base con le scadenze presenti nel documento
Private Sub sbPopalaListaScadenze()
    Dim lRow As Long
    Dim oItem As MSComctlLib.ListItem
    
    'Pilisce la listview
    lvwScadenze.ListItems.Clear
    
    With oDoc.Tables(sTabellaScadenze)
        'Cicla per tutte le righe di scadenze presenti nel documento
        For lRow = 1 To .NumRetails
            Set oItem = lvwScadenze.ListItems.Add
            
            'Popola l'item della listview
            oItem.Text = .Fields("Sca_data_scadenza").Values(lRow).Value
            oItem.SubItems(1) = FormatNumber(fnNotNullN(.Fields("Sca_importo_scadenza").Values(lRow).Value), 2)
        Next lRow
    End With
End Sub

'Popola la listview del castelletto Iva in base ai dati presenti nel documento
Private Sub sbPopalaListaIva()
    Dim lRow As Long
    Dim oItem As MSComctlLib.ListItem
    
    'Pilisce la listview
    lvwIVA.ListItems.Clear
    
    With oDoc.Tables(sTabellaIVA)
        'Cicla per tutte le righe del castelletto Iva presenti nel documento
        For lRow = 1 To .NumRetails
            Set oItem = lvwIVA.ListItems.Add
            
            'Popola l'item della listview
            oItem.Text = fnNotNullN(.Fields("Cst_aliquota_IVA").Values(lRow).Value)
            oItem.SubItems(1) = fnNotNull(.Fields("Cst_descrizione_IVA").Values(lRow).Value)
            oItem.SubItems(2) = FormatNumber(fnNotNullN(.Fields("Cst_imponibile_IVA_corr").Values(lRow).Value), 2)
            oItem.SubItems(3) = FormatNumber(fnNotNullN(.Fields("Cst_imposta_IVA_corr").Values(lRow).Value), 2)
        Next lRow
    End With

End Sub

'Prende in variazione una riga di dettaglio del documento
'Questa procedura viene richiamata ogni volta che si click in un item della listview articoli
Private Sub sbVariazioneRiga()

    If Not lvwArticoli.SelectedItem Is Nothing Then
        'Imposta che ci stiamo predisponendo a prendere in modifica un dettaglio
        bVariazioneDettaglio = True
        A_Riga(0) = 0
        A_Riga(1) = 0

        'Si posiziona nella riga selezionata
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail lvwArticoli.SelectedItem.Index
        
        If oDoc.Field("RV_POTipoRiga", , sTabellaDettaglio) = 1 Then
            A_Riga(0) = lvwArticoli.SelectedItem.Index
            NumeroRecordPerModifica = oDoc.Field("IDValoriOggettoDettaglio", , sTabellaDettaglio)
            NumeroRigaSelezionata = oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio)
            Me.CDArticolo.Load fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))
            'Me.txtCodiceArticolo.Text = fnNotNull(oDoc.Field("Art_codice", , sTabellaDettaglio))
            Me.txtDescrizioneArticolo.Text = fnNotNull(oDoc.Field("Art_descrizione", , sTabellaDettaglio))
            Me.cboAliquotaArticolo.WriteOn fnNotNullN(oDoc.Field("Link_Art_IVA", , sTabellaDettaglio))
            Me.txtAliquotaArticolo.Value = fnNotNullN(oDoc.Field("Art_aliquota_IVA", , sTabellaDettaglio))
            Me.cboUnitaDiMisura.WriteOn fnNotNullN(oDoc.Field("Link_Art_unita_di_misura", , sTabellaDettaglio))
            Me.txtImportoUnitarioArticolo.Value = fnNotNullN(oDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglio))
            Me.txtImponibileUnitario.Value = fnNotNullN(oDoc.Field("Art_pre_uni_net_sco_net_IVA", , sTabellaDettaglio))
            Me.txtSconto1.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglio))
            Me.txtSconto2.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglio))
            
            Me.CDPedana.Load fnNotNullN(oDoc.Field("RV_POIDArticoloPedana", , sTabellaDettaglio))
            'Me.txtDescrizionePedana.Text = fnNotNull(oDoc.Field("RV_PODescrizioneArticoloPedana", , sTabellaDettaglio))
            Me.CDTipoPedana.Load fnNotNullN(oDoc.Field("RV_POIDTipoPedana", , sTabellaDettaglio))
            Me.txtQuantitaPedana.Value = fnNotNullN(oDoc.Field("RV_POQuantitaPedana", , sTabellaDettaglio))
            Me.cboUMRigaOrdine.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoUMOrdine", , sTabellaDettaglio))
            Me.txtQuantitaPerPedana.Value = fnNotNullN(oDoc.Field("RV_POColliPerPedana", , sTabellaDettaglio))
            
            

            'IMBALLO PRIMARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Me.CDImballoPrimario.Load fnNotNullN(oDoc.Field("RV_POIDImballoPrimario", , sTabellaDettaglio))
            Me.txtTaraConfImballo.Value = fnNotNullN(oDoc.Field("RV_POTaraImballoPrimario", , sTabellaDettaglio))
            Me.txtNumeroConfImballo.Value = fnNotNullN(oDoc.Field("RV_PONumeroConfezioniPerImballo", , sTabellaDettaglio))
            
            Me.txtQuantitaPerCollo.Value = fnNotNullN(oDoc.Field("RV_POQuantitaPerCollo", , sTabellaDettaglio))
            Me.txtPesoPerCollo.Value = fnNotNullN(oDoc.Field("RV_POPesoPerCollo", , sTabellaDettaglio))
            Me.txtMoltiplicatore.Value = fnNotNullN(oDoc.Field("RV_POMoltiplicatorePerCollo", , sTabellaDettaglio))
            
            Me.txtColli.Value = fnNotNullN(oDoc.Field("Art_numero_colli", , sTabellaDettaglio))
            Me.txtPesoLordo.Value = fnNotNullN(oDoc.Field("Art_peso", , sTabellaDettaglio))
            Me.txtTara.Value = fnNotNullN(oDoc.Field("Art_tara", , sTabellaDettaglio))
            Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
            Me.txtPezzi.Value = fnNotNullN(oDoc.Field("Art_quantita_pezzi", , sTabellaDettaglio))
            Me.txtQta_UM.Value = fnNotNullN(oDoc.Field("Art_quantita_totale", , sTabellaDettaglio))
            Me.txtImponibileArticolo.Value = fnNotNullN(oDoc.Field("Art_importo_totale_neutro", , sTabellaDettaglio))
            'Me.txtTotaleArticolo.Value = fnNotNullN(oDoc.Field("Art_importo_totale_lordo_IVA", , sTabellaDettaglio))
            
            Me.cboCalibro.WriteOn fnNotNullN(oDoc.Field("RV_POIDCalibro", , sTabellaDettaglio))
            Me.cboTipoCategoria.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoCategoria", , sTabellaDettaglio))
            Me.cboTipoLavorazione.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoLavorazione", , sTabellaDettaglio))
            
            Me.chkImportoImballoInArticolo.Value = Abs(CLng(Val((fnNotNullN(oDoc.Field("RV_POImportoImballoInArticolo", , sTabellaDettaglio))))))

            
            Me.txtAnnotazioniDiRiga.Text = fnNotNull(oDoc.Field("RV_POAnnotazioniRigaOrdine", , sTabellaDettaglio))
            Me.txtAnnotazioniDiRigaLav.Text = fnNotNull(oDoc.Field("RV_POAnnotazioniRigaLavorazione", , sTabellaDettaglio))
            
            'DETTAGLIO DELL'ORDINE VIVAIO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Me.CDArticoloPianale.Load fnNotNullN(oDoc.Field("RV_PO01_IDArticoloPianale", , sTabellaDettaglio))
            Me.txtQtaPianale.Value = fnNotNullN(oDoc.Field("RV_PO01_QuantitaPianale", , sTabellaDettaglio))
            
            Me.CDArticoloProlunga.Load fnNotNullN(oDoc.Field("RV_PO01_IDArticoloProlunga", , sTabellaDettaglio))
            Me.txtQtaProlunga.Value = fnNotNullN(oDoc.Field("RV_PO01_QuantitaProlunga", , sTabellaDettaglio))
            
            Me.ACSSocio.sbLoadCFByIDAnagrafica 7, fnNotNullN(oDoc.Field("RV_PO01_IDAnagraficaFornitore", , sTabellaDettaglio))
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            Me.txtImportoListinoArticolo.Value = fnNotNullN(oDoc.Field("RV_POImportoUnitarioListino", , sTabellaDettaglio))
            
            Me.txtAnnotazioniPerSocio.Text = fnNotNull(oDoc.Field("RV_PO01_AnnotazioniSocio", , sTabellaDettaglio))
            
            Me.txtRaggrRigaOrdine.Text = fnNotNull(oDoc.Field("RV_PONotaRigaOrdRaggr", , sTabellaDettaglio))
            Me.txtQuantitaPedanaEff.Value = fnNotNull(oDoc.Field("RV_POQuantitaPedanaEffettiva", , sTabellaDettaglio))
            Me.txtColliSfusi.Value = fnNotNull(oDoc.Field("RV_POColliSfusi", , sTabellaDettaglio))
            Me.cboReportNew.WriteOn fnNotNullN(oDoc.Field("RV_POIDConfigurazioneEtichettaLavorazione", , sTabellaDettaglio))
            Me.cboReportPedNew.WriteOn fnNotNullN(oDoc.Field("RV_POIDConfigurazioneEtichettaPedana", , sTabellaDettaglio))
            Me.txtIDRigaContratto.Value = oDoc.Field("RV_POIDValoriOggettoDettaglioContratto", , sTabellaDettaglio)
            
            LINK_LOTTO_PROD_LAV = fnNotNullN(oDoc.Field("RV_POIDLottoCampagnaLavorazione", , sTabellaDettaglio))
            If (LINK_LOTTO_PROD_LAV > 0) Then
                Me.txtRaggrRigaOrdine.Text = GET_DESCRIZIONE_LOTTO_PROD_LAV(LINK_LOTTO_PROD_LAV)
            Else
                Me.txtRaggrRigaOrdine.Text = oDoc.Field("RV_PONotaRigaOrdRaggr", , sTabellaDettaglio)
            End If
            
            If oDoc.Field("RV_PORigaCompleta", , sTabellaDettaglio) = 1 Then
                TrovaRigaAssociata 2, oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio)
            Else
                Me.CDImballo.Load 0
                Me.txtDescrizioneImballo.Text = ""
                Me.txtTaraUnitaria.Value = 0
                Me.CboAliquotaImballo.WriteOn 0
                Me.txtAliquotaArticolo.Value = 0
                Me.txtImportoUnitarioImballo.Value = 0
            End If
            
            CDImballoPrimario_ChangeElement
            
        Else
        
            A_Riga(1) = lvwArticoli.SelectedItem.Index
            Me.CDImballo.Load fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))
            Me.txtDescrizioneImballo.Text = fnNotNull(oDoc.Field("Art_descrizione", , sTabellaDettaglio))
            Me.txtTaraUnitaria.Value = fnNotNullN(oDoc.Field("Art_tara", , sTabellaDettaglio))
            Me.CboAliquotaImballo.WriteOn fnNotNullN(oDoc.Field("Link_Art_IVA", , sTabellaDettaglio))
            Me.txtAliquotaImballo.Value = fnNotNullN(oDoc.Field("Art_aliquota_IVA", , sTabellaDettaglio))
            Me.chkImportoImballoInArticolo.Value = Abs(CLng(Val((fnNotNullN(oDoc.Field("RV_POImportoImballoInArticolo", , sTabellaDettaglio))))))
            Me.txtImportoUnitarioImballo.Value = fnNotNullN(oDoc.Field("Art_prezzo_unitario_netto_IVA", , sTabellaDettaglio))
            Me.txtImponibileImballo.Value = fnNotNullN(oDoc.Field("Art_importo_totale_netto_IVA", , sTabellaDettaglio))
            Me.txtColli.Value = fnNotNullN(oDoc.Field("Art_quantita_totale", , sTabellaDettaglio))
            
            
            Me.cboCalibro.WriteOn fnNotNullN(oDoc.Field("RV_POIDCalibro", , sTabellaDettaglio))
            Me.cboTipoCategoria.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoCategoria", , sTabellaDettaglio))
            Me.cboTipoLavorazione.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoLavorazione", , sTabellaDettaglio))
            Me.CDPedana.Load fnNotNullN(oDoc.Field("RV_POIDArticoloPedana", , sTabellaDettaglio))
            Me.txtQuantitaPedana.Value = fnNotNullN(oDoc.Field("RV_POQuantitaPedana", , sTabellaDettaglio))
            Me.cboUMRigaOrdine.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoUMOrdine", , sTabellaDettaglio))
            Me.txtQuantitaPerPedana.Value = fnNotNullN(oDoc.Field("RV_POColliPerPedana", , sTabellaDettaglio))
            Me.CDTipoPedana.Load fnNotNullN(oDoc.Field("RV_POIDTipoPedana", , sTabellaDettaglio))
            Me.txtAnnotazioniDiRiga.Text = fnNotNull(oDoc.Field("RV_POAnnotazioniRigaOrdine", , sTabellaDettaglio))
            
            
            If oDoc.Field("RV_PORigaCompleta", , sTabellaDettaglio) = 1 Then
                TrovaRigaAssociata 1, oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio)
            Else
                
                
                Me.CDArticolo.Load 0
                Me.txtDescrizioneArticolo.Text = ""
                Me.txtPesoLordo.Value = 0
                Me.txtPesoNetto.Value = 0
                Me.txtPezzi.Value = 0
                Me.txtTara.Value = 0
                Me.cboAliquotaArticolo.WriteOn 0
                Me.txtAliquotaArticolo.Value = 0
                Me.txtImportoUnitarioArticolo.Value = 0
                Me.cboUnitaDiMisura.WriteOn 0
                Me.txtImponibileUnitario.Value = 0
                
            End If

        End If
        
        
        CalcolaTotaleRiga

    Else
        
        cmdNuovo_Click
    
    End If
End Sub
Private Function GetColliOriginari() As Long

End Function
'Riporta i dati del dettaglio articolo dal form alla riga attiva nel dettaglio
'dell'oggetto cDocumento ed effettua il calcolo della tabella di dettaglio
Private Sub sbImpostaECalcolaDatiRiga()
    'Riporta i dati dal form
    'oDoc.Field "Art_quantita_totale", lngQta.Value, sTabellaDettaglio
    'oDoc.Field "Art_prezzo_unitario_neutro", curPrezzo.Value, sTabellaDettaglio
    'oDoc.Field "Art_sco_in_percentuale_1", lngSconto.Value, sTabellaDettaglio
    
    'Effettua il calcolo solo della tabella di dettaglio per ottenerne i totali di riga
    oDoc.PerformTable sTabellaDettaglio, True
End Sub

'Riporta i dati del documento dal form alla all'oggetto cDocumento
Private Sub sbImpostaDatiDocumento()
    'Riporta i dati dal form
    If Len(dtData.Text) > 0 Then
        oDoc.DataEmissione = dtData.Text
    End If

    oDoc.Field "Doc_data", oDoc.DataEmissione, sTabellaTestata
    
    oDoc.Numero = lngNumero.Value
    oDoc.Field "Doc_numero", oDoc.Numero, sTabellaTestata
    oDoc.Field "Link_Nom_IVA", LINK_CLIENTE_IVA, sTabellaTestata
    oDoc.Field "Doc_prezzi_lordo_IVA", Me.chkLordoIVA.Value, sTabellaTestata
    'oDoc.Field "Nom_IVA_in_spospensione", Me.chkSpospensioneIva.Value, sTabellaTestata
    oDoc.Field "Link_Doc_pagamento", cboPagamento.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_listino", cboListino.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_listino_base", cboListinoAzienda.CurrentID, sTabellaTestata
    oDoc.Field "Spe_incasso_neutro", curSpeseIncasso.Value, sTabellaTestata
    oDoc.Field "Spe_trasporto_neutro", curSpeseTrasporto.Value, sTabellaTestata
    oDoc.Field "Sco_percentuale_fine_documento", lngScontoDocPer.Value, sTabellaTestata
    oDoc.Field "Sco_ad_importo_fine_documento", curScontoDocImp.Value, sTabellaTestata
    oDoc.Field "Link_Nom_porto", Me.cboPorto.CurrentID, sTabellaTestata
    oDoc.Field "Link_Vet_vettore", Me.cboVettore.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_spedizione", Me.cboTrasporto.CurrentID, sTabellaTestata

    oDoc.Field "Link_Doc_magazzino", Me.cboMagazzino.CurrentID, sTabellaTestata
    oDoc.Field "Nom_ult_sito_cap", Me.txtCapAltroSito.Text, sTabellaTestata
    oDoc.Field "Nom_ult_sito_comune", Me.txtComuneAltroSito.Text, sTabellaTestata
    'oDoc.Field "Nom_ult_sito_nazione", Me.txtNazioneAltroSito.Text, sTabellaTestata
    oDoc.Field "Nom_ult_sito_indirizzo", Me.txtIndirizzoAltroSito.Text, sTabellaTestata
    oDoc.Field "Nom_ult_sito_referente", Me.txtReferenteAltroSito.Text, sTabellaTestata
    oDoc.Field "Link_Nom_ult_sito", Me.cboAltroSito.CurrentID, sTabellaTestata
    oDoc.Field "Nom_ult_sito_provincia", Me.txtProvinciaAltroSito, sTabellaTestata
    oDoc.Field "Nom_raggruppamento_bolle", Me.chkRaggruppBolle.Value, sTabellaTestata
    oDoc.Field "Nom_raggruppamento_scadenze", Me.chkRaggruppaScadenze.Value, sTabellaTestata
    oDoc.Field "Doc_causale_trasporto", Me.txtCausaleDocumento.Text, sTabellaTestata
    oDoc.Field "Doc_causale_documento", Me.txtCausaleDocumentoEF.Text, sTabellaTestata
    
    oDoc.Field "Link_Nom_IVA", Me.cboIvaCliente.CurrentID, sTabellaTestata
    oDoc.Field "Link_Nom_lettera_intento", Me.txtIDLetteraIntento.Value, sTabellaTestata
    oDoc.Field "Link_Nom_accordi_commerciali", Me.cboAccordoCommerciale.CurrentID, sTabellaTestata
    oDoc.Field "Link_Nom_raggrup_fatturato", Me.cboRaggrFatturato.CurrentID, sTabellaTestata
    oDoc.Field "RV_POIDAnagraficaDestinazione", Me.ACSAnaDest.IDAnagrafica, sTabellaTestata
    
    oDoc.Field "Link_Val_valuta", Me.cboValuta.CurrentID, sTabellaTestata
    
'    If Me.cboCambioValuta.CurrentID = 0 Then
'        oDoc.Field "Link_Val_cambio", Null, sTabellaTestata
'        oDoc.Field "Val_valore_cambio", Null, sTabellaTestata
'        oDoc.Field "Val_data_cambio", Null, sTabellaTestata
'    Else
'        oDoc.Field "Link_Val_cambio", Me.cboCambioValuta.CurrentID, sTabellaTestata
'        oDoc.Field "Val_valore_cambio", Me.txtValoreCambioValuta.Value, sTabellaTestata
'        oDoc.Field "Val_data_cambio", Me.txtDataCambio.Text, sTabellaTestata
'    End If
    
    oDoc.Field "Tot_numero_colli", Me.txtColliTotali.Value, sTabellaTestata
    oDoc.Field "Tot_peso", Me.txtPesoTotale.Value, sTabellaTestata
    
    ''AGENTE
    oDoc.Field "Link_doc_agente", Me.CDAgenteTesta.KeyFieldID, sTabellaTestata
    oDoc.Field "Doc_age_ragione_sociale", Me.CDAgenteTesta.Code, sTabellaTestata
    oDoc.Field "Doc_age_nome", Me.CDAgenteTesta.Description, sTabellaTestata
    oDoc.Field "Doc_age_codice", GET_CODICE_AGENTE(Me.CDAgenteTesta.KeyFieldID), sTabellaTestata
    
    oDoc.Field "RV_POIstruzioniMittente", Me.txtIstruzioniMittente.Text, sTabellaTestata
    oDoc.Field "RV_POTargaAutomezzo", Me.txtTargaAutomezzo.Text, sTabellaTestata
    
    'ALTRI DATI
    oDoc.Field "RV_POIDLuogoPresaMerce", Me.cboLuogoPresaMerce.CurrentID, sTabellaTestata
    oDoc.Field "RV_POIDTrasportatoreSuccessivo", Me.cboVettoreSuccessivo.CurrentID, sTabellaTestata
    oDoc.Field "Doc_data_presso_nom", Me.txtDataOrdineCliente.Text, sTabellaTestata
    oDoc.Field "Doc_numero_presso_nom", Me.txtNumeroOrdineCliente.Text, sTabellaTestata
    
    oDoc.Field "RV_PODataArrivoMerce", Me.txtDataTrasporto.Text, sTabellaTestata
    oDoc.Field "RV_POOraArrivoMerce", Me.txtOraTrasporto.Text, sTabellaTestata
    
    oDoc.Field "RV_PODataArrivoMerceLuogo", Me.txtDataArrivoLuogo.Text, sTabellaTestata
    oDoc.Field "RV_POOraArrivoMerceLuogo", Me.txtOraArrivoLuogo.Text, sTabellaTestata
    
    oDoc.Field "Doc_data_prevista_evasione", Me.txtDataPartenza.Text, sTabellaTestata
    oDoc.Field "Doc_ordine_chiuso", Me.chkChiuso.Value, sTabellaTestata
    oDoc.Field "Doc_annotazioni_variazio", Mid(Me.txtAnnotazioni.Text, 1, 250), sTabellaTestata
    oDoc.Field "RV_POAnnotazioniInterna", Mid(Me.txtAnnotazioniInterna.Text, 1, 250), sTabellaTestata
    oDoc.Field "RV_PODescrizioneCorpoDocEv", Mid(Me.txtDescrizioneRigaDoc.Text, 1, 250), sTabellaTestata
    oDoc.Field "RV_POIDTipoOrdine", Me.cboTipoOrdine.CurrentID, sTabellaTestata

    oDoc.Field "Link_Doc_aspetto_esteriore", Me.cboAspettoEsteriore.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_spedizione", Me.cboTrasporto.CurrentID, sTabellaTestata

    'NUMERO ORDINE PADRE
    oDoc.Field "RV_PONumeroOrdinePadre", Me.txtNOrdinePadre.Value, sTabellaTestata
    oDoc.Field "RV_PODataOrdinePadre", Me.txtDataDocPadre.Value, sTabellaTestata
    oDoc.Field "RV_PONumeroListaPrelievo", Me.txtNListaPrelievo.Value, sTabellaTestata
    
    oDoc.Field "RV_POOrdineCompletato", Me.chkOrdineCompletato.Value, sTabellaTestata

    oDoc.Field "RV_PONumeroPedanePrelievo", Me.txtNPedaneTesta.Value, sTabellaTestata
    
    oDoc.Field "RV_POIDOggettoContratto", Me.txtIDContratto.Value, sTabellaTestata
    oDoc.Field "RV_POConfPresaVisContratto", Me.chkConfDaContratto.Value, sTabellaTestata
    oDoc.Field "RV_POFatturaProforma", Me.chkStampaFattProForma.Value, sTabellaTestata
    
End Sub

Private Sub fnRecuperaAnnotazioniPerDoc()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_PONoteDocumentiCoop.Annotazioni1, RV_PONoteDocumentiCoop.Annotazioni2, RV_PONoteDocumentiCoop.Annotazioni3 "
sSQL = sSQL & "FROM RV_PONoteDocumentiCoop INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_PONoteDocumentiCoop.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale =" & m_App.Branch & ") And (RV_POSchemaCoop.IDUtente = 0) "
sSQL = sSQL & " AND (RV_PONoteDocumentiCoop.IDRV_PODocumentiCoop = " & IDDocumento & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    oDoc.Field "RV_POAnnotazioni1", fnNotNull(rs!Annotazioni1), sTabellaTestata
    oDoc.Field "RV_POAnnotazioni2", fnNotNull(rs!Annotazioni2), sTabellaTestata
    oDoc.Field "RV_POAnnotazioni3", fnNotNull(rs!Annotazioni3), sTabellaTestata
Else
    oDoc.Field "RV_POAnnotazioni1", "", sTabellaTestata
    oDoc.Field "RV_POAnnotazioni2", "", sTabellaTestata
    oDoc.Field "RV_POAnnotazioni3", "", sTabellaTestata
End If

rs.CloseResultset
Set rs = Nothing

End Sub

'Effettua il calcolo completo del documento in base ai dati specificati
Private Sub sbCalcolaDocumento()
    'Effettua il calcolo solo se non siamo in fase di caricamento
    'di un documento esistente (variazione del documento)
    If Not bloading Then
        'Effettuta il calcolo del documento
        oDoc.PerformDocument Nothing, False
        'Aggiorna il contenuto delle listview delle scadenze e del castelletto Iva
        sbPopalaListaScadenze
        sbPopalaListaIva
    End If
End Sub

Public Sub ConnessioneDiamanteADO()
On Error GoTo ERR_ConnessioneDiamanteADO
    
    Set Cn = m_App.Database.Connection
    
Exit Sub
ERR_ConnessioneDiamanteADO:
    MsgBox Err.Description, vbCritical, "Connessione Diamante di tipo ADO"
End Sub


Private Sub PBOrdine_Click()
    cmdAnalizzaOrdine_Click
End Sub

Private Sub PBOrdineDettaglio_Click()
On Error GoTo ERR_cmdAnalizzaOrdine_Click
    If oDoc.IDOggetto > 0 Then
        
        LINK_ORDINE_SELEZIONATO = oDoc.IDOggetto
        LINK_ART_ORD_PADRE_SEL = Me.CDArticolo.KeyFieldID
        frmAnalizzaOrdine.Show vbModal
    End If
Exit Sub
ERR_cmdAnalizzaOrdine_Click:
    MsgBox Err.Description, vbCritical, "PBOrdineDettaglio_Click"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next

If Me.SSTab1.Tab = 1 Then
    cmdNuovo_Click
    
    Me.lblInfoTesta.Caption = ""
    If Me.cdAnagrafica.KeyFieldID > 0 Then
        Me.lblInfoTesta.Caption = "Cliente: " & Me.cdAnagrafica.Code
            If Me.cdAnagrafica.Description <> "" Then
                Me.lblInfoTesta.Caption = Me.lblInfoTesta.Caption & " " & Me.cdAnagrafica.Description
            End If
    Else
        Me.lblInfoTesta.Caption = "Cliente: NON IMPOSTATO"
        
    End If

    Me.lblInfoTesta.Caption = Me.lblInfoTesta.Caption & " - Riferimento documento: " & Me.dtData.Text & " n° " & Me.lngNumero.Text
    
    
End If

End Sub




Private Function GET_TOTALE_MERCE_DOCUMENTO() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoListinoImballo As Double


GET_TOTALE_MERCE_DOCUMENTO = 0

sSQL = "SELECT Art_importo_totale_neutro, RV_POIDImballo, RV_POImportoUnitarioImballo,RV_POImportoMerceNetta,Art_quantita_totale, "
sSQL = sSQL & "RV_POImportoImballoInArticolo, Art_numero_colli "
sSQL = sSQL & "FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE (IDOggetto = " & oDoc.IDOggetto & ") "
sSQL = sSQL & " AND (RV_POTipoRiga = 1)"


Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    
    GET_TOTALE_MERCE_DOCUMENTO = GET_TOTALE_MERCE_DOCUMENTO + (fnNotNullN(rs!RV_POImportoMerceNetta) * fnNotNullN(rs!Art_quantita_totale))
    
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
    
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_PREZZO_IMBALLO_PER_COMMMISSIONI = fnNotNullN(rs!PrezzoNettoIVA)
Else
    GET_PREZZO_IMBALLO_PER_COMMMISSIONI = 0
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub txtAnnotazioni_LostFocus()
    sbImpostaDatiDocumento
End Sub
Private Sub txtAnnotazioniInterna_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtCausaleDocumento_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtCausaleDocumentoEF_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtColli_Change()
If Link_UMCoop = 1 Then
    Me.txtQta_UM.Value = Me.txtColli.Value
End If
    
'Me.txtTara.Value = Me.txtColli.Value * Me.txtTaraUnitaria.Value

End Sub

Private Sub txtColli_LostFocus()
On Error Resume Next
Dim Quantita As Double

'If bVariazioneDettaglio = False Then
'    Me.txtPesoLordo.Value = Me.txtColli.Value * PESO_LORDO
    Me.txtTara.Value = Me.txtColli.Value * (Me.txtTaraUnitaria.Value + (Me.txtNumeroConfImballo.Value * Me.txtTaraConfImballo.Value))
    If (GESTIONE_ORDINE_VIVAIO = 0) Then
        If (Me.txtQuantitaPerPedana.Value > 0) Then
            Me.txtQuantitaPedana.Value = fnRoundDown(Me.txtColli.Value / Me.txtQuantitaPerPedana.Value)
        Else
            Me.txtQuantitaPedana.Value = Me.txtQuantitaPedana.Value
        End If
    Else
        Me.txtQuantitaPedana.Value = Me.txtQuantitaPedana.Value
    End If
    Me.txtColliSfusi.Value = Me.txtColli.Value - (Me.txtQuantitaPedana.Value * Me.txtQuantitaPerPedana.Value)
    Me.txtQuantitaPedanaEff.Value = Me.txtColli.Value / Me.txtQuantitaPerPedana.Value
    
    
'End If

If TIPO_PESO_ARTICOLO <= 1 Then
    If Me.txtPesoPerCollo.Value > 0 Then
        Me.txtPesoLordo.Value = Me.txtPesoPerCollo.Value * Me.txtColli.Value
        Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
    End If
Else
    If Me.txtPesoPerCollo.Value > 0 Then
        Me.txtPesoNetto.Value = Me.txtPesoPerCollo.Value * Me.txtColli.Value
        Me.txtPesoLordo.Value = Me.txtPesoNetto.Value + Me.txtTara.Value
    End If
End If


If txtNumeroConfImballo.Value > 0 Then
    If Me.txtQuantitaPerCollo.Value > 0 Then
        Me.txtPezzi.Value = Me.txtColli.Value * Me.txtQuantitaPerCollo.Value
    Else
        Me.txtPezzi.Value = Me.txtColli.Value * txtNumeroConfImballo.Value
    End If
Else
    If Me.txtQuantitaPerCollo.Value > 0 Then
        Me.txtPezzi.Value = Me.txtColli.Value * txtQuantitaPerCollo.Value
    End If
End If

If Link_UMCoop = 1 Then
    Me.txtQta_UM.Value = Me.txtColli.Value
End If

CalcoloPesoNetto

End Sub

Private Sub txtColliSfusi_LostFocus()
    If (GESTIONE_ORDINE_VIVAIO = 0) Then
        'If bVariazioneDettaglio = False Then
            'Me.txtColli.Value = (Me.txtQuantitaPedana.Value * GET_QUANTITA_IMBALLO_PER_TIPO_PEDANA(Me.CDTipoPedana.KeyFieldID, Me.CDImballo.KeyFieldID)) + Me.txtColliSfusi.Value
            Me.txtColli.Value = Me.txtQuantitaPedana.Value * Me.txtQuantitaPerPedana.Value + Me.txtColliSfusi.Value
            txtColli_LostFocus
        'End If
    Else
        Me.txtQuantitaPedanaEff.Value = Me.txtQuantitaPedana.Value
    End If
End Sub

Private Sub txtColliTotali_Change()
    sbImpostaDatiDocumento
End Sub

Private Sub ParametroTipoArrotondamento()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoArrotondamento FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_Arrotondamento = fnNotNullN(rs!IDTipoArrotondamento)
Else
    Link_Arrotondamento = 1
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroUMRigaOrdine()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoUMOrdine, AttivaCommissioniDaOrdine, ConfPresaVisAutOrdDaContr, "
sSQL = sSQL & " IvaArticoloDaDocumentoCollegato, LetteraIntentoDaDocumentoCollegato "
sSQL = sSQL & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    If fnNotNullN(rs!IDRV_POTipoUMOrdine) > 0 Then
        LINK_TIPO_UM_RIGA_ORDINE = fnNotNullN(rs!IDRV_POTipoUMOrdine)
    Else
        LINK_TIPO_UM_RIGA_ORDINE = 1
    End If
    ATTIVA_COMMISSIONI_DA_ORDINE = fnNotNullN(rs!AttivaCommissioniDaOrdine)
    CONF_AUT_PRESA_VIS_CONTR = fnNotNullN(rs!ConfPresaVisAutOrdDaContr)
    RIP_IVA_DA_DOC_COLL = fnNotNullN(rs!IvaArticoloDaDocumentoCollegato)
    RIP_LET_INT_DA_DOC_COLL = fnNotNullN(rs!LetteraIntentoDaDocumentoCollegato)
Else
    LINK_TIPO_UM_RIGA_ORDINE = 1
    ATTIVA_COMMISSIONI_DA_ORDINE = 0
    CONF_AUT_PRESA_VIS_CONTR = 0
    RIP_IVA_DA_DOC_COLL = 0
    RIP_LET_INT_DA_DOC_COLL = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub


Private Sub ParametroImballo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoImballo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoImballo = rs!IDTipoImballo
Else
    Link_TipoImballo = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Public Sub AGGIORNA_CONTENITORE_DATI_LOTTO()
End Sub

Private Sub ParametroSocio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoSocio = rs!IDCategoriaAnagrafica
Else
    Link_TipoSocio = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub txtDataArrivoLuogo_LostFocus()
sbImpostaDatiDocumento
End Sub

Private Sub txtDataDocPadre_Change()
    sbImpostaDatiDocumento
    
End Sub

Private Sub txtDataOrdineCliente_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtDataPartenza_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtDataTrasporto_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtDescrizioneArticolo_LostFocus()
    lblDocument_Click 24
End Sub

Private Sub txtDescrizioneImballo_LostFocus()
    lblDocument_Click 36
End Sub



'Private Sub txtDescrizionePedana_LostFocus()
'lblDocument_Click 9
'End Sub

Private Sub txtDescrizioneRigaDoc_Change()
sbImpostaDatiDocumento
End Sub

Private Sub txtIDContratto_Change()
On Error GoTo ERR_txtIDContratto_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtDescrContratto.Text = ""

sSQL = "SELECT IDOggetto, IDTipoOggetto, Doc_numero, Doc_data "
sSQL = sSQL & "FROM RV_POIEContratto "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDContratto.Value

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtDescrContratto.Text = "Contratto numero " & fnNotNull(rs!Doc_numero) & " del " & fnNotNull(rs!Doc_data)
End If

rs.CloseResultset
Set rs = Nothing

If (oDoc.IDOggetto = 0) Then
    If (CONF_AUT_PRESA_VIS_CONTR = 1) Then
        Me.chkConfDaContratto.Value = vbChecked
    End If
        
        
    
End If

sbImpostaDatiDocumento
Exit Sub
ERR_txtIDContratto_Change:
    MsgBox Err.Description, vbCritical, "txtIDContratto_Change"
End Sub

Private Sub txtIDLetteraIntento_Change()
On Error GoTo ERR_txtIDLetteraIntento_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If IsEmpty(Me.txtIDLetteraIntento.Value) Then Me.txtIDLetteraIntento.Value = 0

sSQL = "SELECT * FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & Me.txtIDLetteraIntento.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtNLetteraIntento.Text = ""
    Me.txtDataLetteraIntento.Value = 0
    Me.lblLetteraIntento.ToolTipText = ""
    
Else
    Me.txtNLetteraIntento.Text = fnNotNull(rs!Numero)
    Me.txtDataLetteraIntento.Value = fnNotNullN(rs!Data)
    Me.lblLetteraIntento.ToolTipText = "Prot. N° " & fnNotNull(rs!NumeroCliFor) & " del " & fnNotNull(rs!DataEmissione)
End If

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_txtIDLetteraIntento_Change:
    MsgBox Err.Description, vbCritical, "txtIDLetteraIntento_Change"
End Sub

Private Sub txtImponibileUnitario_Change()

    Me.txtScontoImpListino.Value = GET_SCONTO_IMPORTO_LISTINO
    
End Sub

Private Sub txtImportoUnitarioArticolo_LostFocus()
    CalcolaImportoScontato
    CalcolaTotaleRiga
    
    
End Sub
Private Sub CalcolaImportoScontato()
    Me.txtImponibileUnitario.Value = Me.txtImportoUnitarioArticolo.Value - ((Me.txtImportoUnitarioArticolo.Value / 100)) * (Me.txtSconto1.Value)
    Me.txtImponibileUnitario.Value = Me.txtImponibileUnitario.Value - ((Me.txtImponibileUnitario.Value / 100)) * (Me.txtSconto2.Value)
End Sub
Private Sub txtImportoUnitarioImballo_LostFocus()
    CalcolaTotaleRiga
End Sub
Private Sub CalcolaTotaleRiga()
Dim TOT_ARTICOLO As Double
Dim TOT_IMBALLO As Double
    If bVariazioneDettaglio = False Then
        'Calcolo della tara
        If Me.txtTaraUnitaria.Value > 0 Then
            If Me.txtTara.Value = 0 Then
                Me.txtTara.Value = Me.txtTaraUnitaria.Value * Me.txtColli.Value
            End If
        End If
    End If
    
    
    'Calcolo Articolo
    Me.txtImponibileArticolo.Value = Me.txtQta_UM.Value * Me.txtImponibileUnitario.Value
    TOT_ARTICOLO = Me.txtImponibileArticolo.Value + ((Me.txtImponibileArticolo.Value / 100) * Me.txtAliquotaArticolo.Value)
    'Calcolo Imballo
    Me.txtImponibileImballo.Value = Me.txtColli.Value * Me.txtImportoUnitarioImballo.Value
    TOT_IMBALLO = Me.txtImponibileImballo.Value + ((Me.txtImponibileImballo.Value / 100) * Me.txtAliquotaImballo.Value)
    
    'Totale
    Me.txtTotaleImponibile.Value = Me.txtImponibileArticolo.Value + Me.txtImponibileImballo.Value
    Me.txtTotaleRiga.Value = TOT_ARTICOLO + TOT_IMBALLO

End Sub








Private Sub txtIstruzioniMittente_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtNOrdinePadre_Change()
    sbImpostaDatiDocumento
End Sub

Private Sub txtNPedaneTesta_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtNumeroConfImballo_LostFocus()
    txtColli_LostFocus
End Sub

Private Sub txtNumeroOrdineCliente_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtOraArrivoLuogo_LostFocus()
sbImpostaDatiDocumento
End Sub

Private Sub txtOraTrasporto_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtPesoLordo_Change()
    If Link_UMCoop = 2 Then
        Me.txtQta_UM.Value = Me.txtPesoLordo.Value
    End If

End Sub
Private Sub txtPesoLordo_LostFocus()
    CalcoloPesoNetto
End Sub
Public Sub SalvataggioRiga()
Dim I As Long
If bVariazioneDettaglio = False Then
    NuovaRigaDocumento
Else
    ModificaRigaDocumento
End If
End Sub
Private Sub ControllaNumeroRiga()
On Error Resume Next
If oDoc.IDOggetto = 0 Then
    NumeroRiga = 0
    NumeroProgSingolaRiga = 0
Else
    NumeroRiga = fncNumeroRiga
    NumeroProgSingolaRiga = fncNumeroProgressivoSingolaRiga
End If

End Sub
Private Function fncNumeroRiga() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POLinkRiga FROM " & sTabellaDettaglio & " "
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto & " "
sSQL = sSQL & "ORDER BY RV_POLinkRiga DESC"

Set rs = Cn.OpenResultset(sSQL)


If rs.EOF Then
    fncNumeroRiga = 0
Else
    If IsNull(rs!RV_POLinkRiga) Then
        fncNumeroRiga = 0
    Else
        fncNumeroRiga = rs!RV_POLinkRiga
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function fncNumeroProgressivoSingolaRiga() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT ID_Art_dettaglio_prog FROM " & sTabellaDettaglio & " "
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto & " "
sSQL = sSQL & "ORDER BY ID_Art_dettaglio_prog DESC"

Set rs = Cn.OpenResultset(sSQL)


If rs.EOF Then
    fncNumeroProgressivoSingolaRiga = 0
Else
    If IsNull(rs!ID_Art_dettaglio_prog) Then
        fncNumeroProgressivoSingolaRiga = 0
    Else
        fncNumeroProgressivoSingolaRiga = rs!ID_Art_dettaglio_prog
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub TrovaRigaAssociata(TipoRiga As Integer, NumeroRiga As Long)
Dim lRow As Integer

If TipoRiga = 2 Then
    'With oDoc.Tables(sTabellaDettaglio)
        'Cicla per tutte le righe di dettaglio presenti nel documento
        For lRow = 1 To Me.lvwArticoli.ListItems.Count
            If Me.lvwArticoli.ListItems(lRow).Text = NumeroRiga Then
                If Me.lvwArticoli.ListItems(lRow).SubItems(1) = TipoRiga Then
                    A_Riga(1) = lRow
                    
                    oDoc.Tables(sTabellaDettaglio).SetActiveRetail lRow
                    
                    Me.CDImballo.Load fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))
                    Me.txtDescrizioneImballo.Text = fnNotNull(oDoc.Field("Art_descrizione", , sTabellaDettaglio))
                    Me.txtTaraUnitaria.Value = fnNotNullN(oDoc.Field("Art_tara", , sTabellaDettaglio))
                    Me.CboAliquotaImballo.WriteOn fnNotNullN(oDoc.Field("Link_Art_IVA", , sTabellaDettaglio))
                    Me.txtAliquotaImballo.Value = fnNotNullN(oDoc.Field("Art_aliquota_IVA", , sTabellaDettaglio))
                    Me.chkImportoImballoInArticolo.Value = Abs(CLng(Val((fnNotNullN(oDoc.Field("RV_POImportoImballoInArticolo", , sTabellaDettaglio))))))
                    Me.txtImportoUnitarioImballo.Value = fnNotNullN(oDoc.Field("Art_prezzo_unitario_netto_IVA", , sTabellaDettaglio))
                    Me.txtImponibileImballo.Value = fnNotNullN(oDoc.Field("Art_importo_totale_netto_IVA", , sTabellaDettaglio))

                    'Me.cboCalibro.WriteOn fnNotNullN(oDoc.Field("RV_POIDCalibro", , sTabellaDettaglio))
                    'Me.cboTipoCategoria.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoCategoria", , sTabellaDettaglio))
                    'Me.cboTipoLavorazione.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoLavorazione", , sTabellaDettaglio))
                    
                    'Me.txtQuantitaPerCollo.Value = fnNotNullN(oDoc.Field("RV_POColliPerPedana", , sTabellaDettaglio))
                    'Me.CDPedana.Load fnNotNullN(oDoc.Field("RV_POIDArticoloPedana", , sTabellaDettaglio))
                    'Me.txtDescrizionePedana.Text = fnNotNull(oDoc.Field("RV_PODescrizioneArticoloPedana", , sTabellaDettaglio))
                   ' Me.txtQuantitaPedana.Value = fnNotNullN(oDoc.Field("RV_POQuantitaPedana", , sTabellaDettaglio))
                    'Me.cboUMRigaOrdine.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoUMOrdine", , sTabellaDettaglio))
                    'Me.txtQuantitaPerCollo.Value = fnNotNullN(oDoc.Field("RV_POColliPerPedana", , sTabellaDettaglio))
                    'Me.CDTipoPedana.Load fnNotNullN(oDoc.Field("RV_POIDTipoPedana", , sTabellaDettaglio))
                    'Me.txtAnnotazioniDiRiga.Text = fnNotNull(oDoc.Field("RV_POAnnotazioniRigaOrdine", , sTabellaDettaglio))

                End If
            End If
        Next lRow
Else
        'Cicla per tutte le righe di dettaglio presenti nel documento
        For lRow = 1 To Me.lvwArticoli.ListItems.Count
            If Me.lvwArticoli.ListItems(lRow).Text = NumeroRiga Then
                If Me.lvwArticoli.ListItems(lRow).SubItems(1) = TipoRiga Then
        
                    A_Riga(0) = lRow
                    NumeroRigaSelezionata = fnNotNullN(oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio))
                    oDoc.Tables(sTabellaDettaglio).SetActiveRetail lRow
                    Me.CDArticolo.Load fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))
                    'Me.txtCodiceArticolo.Text = fnNotNull(oDoc.Field("Art_codice", , sTabellaDettaglio))
                    Me.txtDescrizioneArticolo.Text = fnNotNull(oDoc.Field("Art_descrizione", , sTabellaDettaglio))
                    Me.cboAliquotaArticolo.WriteOn fnNotNullN(oDoc.Field("Link_Art_IVA", , sTabellaDettaglio))
                    Me.txtAliquotaArticolo.Value = fnNotNullN(oDoc.Field("Art_aliquota_IVA", , sTabellaDettaglio))
                    Me.cboUnitaDiMisura.WriteOn fnNotNullN(oDoc.Field("Link_Art_unita_di_misura", , sTabellaDettaglio))
                    Me.txtImportoUnitarioArticolo.Value = fnNotNullN(oDoc.Field("Art_prezzo_unitario_netto_IVA", , sTabellaDettaglio))
                    Me.txtImponibileUnitario.Value = fnNotNullN(oDoc.Field("Art_pre_uni_net_sco_net_iva", , sTabellaDettaglio))
                    Me.txtSconto1.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglio))
                    Me.txtSconto2.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglio))
                    
                    'IMBALLO PRIMARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Me.CDImballoPrimario.Load fnNotNullN(oDoc.Field("RV_POIDImballoPrimario", , sTabellaDettaglio))
                    Me.txtTaraConfImballo.Value = fnNotNullN(oDoc.Field("RV_POTaraImballoPrimario", , sTabellaDettaglio))
                    Me.txtNumeroConfImballo.Value = fnNotNullN(oDoc.Field("RV_PONumeroConfezioniPerImballo", , sTabellaDettaglio))
                    Me.txtQuantitaPerCollo.Value = fnNotNullN(oDoc.Field("RV_POQuantitaPerCollo", , sTabellaDettaglio))
                    Me.txtPesoPerCollo.Value = fnNotNullN(oDoc.Field("RV_POPesoPerCollo", , sTabellaDettaglio))
                    Me.txtMoltiplicatore.Value = fnNotNullN(oDoc.Field("RV_POMoltiplicatorePerCollo", , sTabellaDettaglio))
                    
                    Me.txtColli.Value = fnNotNullN(oDoc.Field("Art_numero_colli", , sTabellaDettaglio))
                    Me.txtPesoLordo.Value = fnNotNullN(oDoc.Field("Art_peso", , sTabellaDettaglio))
                    Me.txtTara.Value = fnNotNullN(oDoc.Field("Art_tara", , sTabellaDettaglio))
                    Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
                    Me.txtPezzi.Value = fnNotNullN(oDoc.Field("Art_quantita_pezzi", , sTabellaDettaglio))
                    Me.txtQta_UM.Value = fnNotNullN(oDoc.Field("Art_quantita_totale", , sTabellaDettaglio))
                    Me.txtImponibileArticolo.Value = fnNotNullN(oDoc.Field("Art_importo_totale_netto_IVA", , sTabellaDettaglio))
                    
                    Me.cboCalibro.WriteOn fnNotNullN(oDoc.Field("RV_POIDCalibro", , sTabellaDettaglio))
                    Me.cboTipoCategoria.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoCategoria", , sTabellaDettaglio))
                    Me.cboTipoLavorazione.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoLavorazione", , sTabellaDettaglio))
                    Me.chkImportoImballoInArticolo.Value = Abs(CLng(Val((fnNotNullN(oDoc.Field("RV_POImportoImballoInArticolo", , sTabellaDettaglio))))))
                    Me.CDPedana.Load fnNotNullN(oDoc.Field("RV_POIDArticoloPedana", , sTabellaDettaglio))
                    'Me.txtDescrizionePedana.Text = fnNotNull(oDoc.Field("RV_PODescrizioneArticoloPedana", , sTabellaDettaglio))
                    Me.txtQuantitaPedana.Value = fnNotNullN(oDoc.Field("RV_POQuantitaPedana", , sTabellaDettaglio))
                    Me.cboUMRigaOrdine.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoUMOrdine", , sTabellaDettaglio))
                    Me.txtQuantitaPerPedana.Value = fnNotNullN(oDoc.Field("RV_POColliPerPedana", , sTabellaDettaglio))
                    Me.CDTipoPedana.Load fnNotNullN(oDoc.Field("RV_POIDTipoPedana", , sTabellaDettaglio))
                    Me.txtAnnotazioniDiRiga.Text = fnNotNull(oDoc.Field("RV_POAnnotazioniRigaOrdine", , sTabellaDettaglio))
                                        
                    Me.cboReportNew.WriteOn fnNotNullN(oDoc.Field("RV_POIDConfigurazioneEtichettaLavorazione", , sTabellaDettaglio))
                    Me.cboReportPedNew.WriteOn fnNotNullN(oDoc.Field("RV_POIDConfigurazioneEtichettaPedana", , sTabellaDettaglio))
                                        
                    'DETTAGLIO DELL'ORDINE VIVAIO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Me.CDArticoloPianale.Load fnNotNullN(oDoc.Field("RV_PO01_IDArticoloPianale", , sTabellaDettaglio))
                    Me.txtQtaPianale.Value = fnNotNullN(oDoc.Field("RV_PO01_QuantitaPianale", , sTabellaDettaglio))
                    
                    Me.CDArticoloProlunga.Load fnNotNullN(oDoc.Field("RV_PO01_IDArticoloProlunga", , sTabellaDettaglio))
                    Me.txtQtaProlunga.Value = fnNotNullN(oDoc.Field("RV_PO01_QuantitaProlunga", , sTabellaDettaglio))
                    
                    Me.ACSSocio.sbLoadCFByIDAnagrafica 7, fnNotNullN(oDoc.Field("RV_PO01_IDAnagraficaFornitore", , sTabellaDettaglio))
                    '''''''
                    Me.txtAnnotazioniPerSocio.Text = fnNotNull(oDoc.Field("RV_PO01_AnnotazioniSocio", , sTabellaDettaglio))
                    Me.txtImportoListinoArticolo.Value = fnNotNullN(oDoc.Field("RV_POImportoUnitarioListino", , sTabellaDettaglio))
                    
                    Me.txtRaggrRigaOrdine.Text = fnNotNull(oDoc.Field("RV_PONotaRigaOrdRaggr", , sTabellaDettaglio))
                    Me.txtQuantitaPedanaEff.Value = fnNotNull(oDoc.Field("RV_POQuantitaPedanaEffettiva", , sTabellaDettaglio))
                    Me.txtColliSfusi.Value = fnNotNull(oDoc.Field("RV_POColliSfusi", , sTabellaDettaglio))
                    Me.txtAnnotazioniDiRigaLav.Text = fnNotNull(oDoc.Field("RV_POAnnotazioniRigaLavorazione", , sTabellaDettaglio))
                    Me.txtIDRigaContratto.Value = oDoc.Field("RV_POIDValoriOggettoDettaglioContratto", , sTabellaDettaglio)
                    
                    LINK_LOTTO_PROD_LAV = fnNotNullN(oDoc.Field("RV_POIDLottoCampagnaLavorazione", , sTabellaDettaglio))
                    If (LINK_LOTTO_PROD_LAV > 0) Then
                        Me.txtRaggrRigaOrdine.Text = GET_DESCRIZIONE_LOTTO_PROD_LAV(LINK_LOTTO_PROD_LAV)
                    Else
                        Me.txtRaggrRigaOrdine.Text = oDoc.Field("RV_PONotaRigaOrdRaggr", , sTabellaDettaglio)
                    End If
                End If
            End If
        Next lRow
 End If
End Sub
Private Sub NuovaRigaDocumento()
Dim I As Integer

If Me.lvwArticoli.ListItems.Count = 0 Then
    
    For I = 1 To oDoc.Tables(sTabellaDettaglio).NumRetails
        oDoc.Tables(sTabellaDettaglio).Delete oDoc.Tables(sTabellaDettaglio).NumRetails
    Next
    
    oDoc.PerformTable sTabellaDettaglio, True
End If

    If Me.CDArticolo.KeyFieldID > 0 Then
            If oDoc.Tables(sTabellaDettaglio).NumRetails = 0 Then
                oDoc.Tables(sTabellaDettaglio).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglio).NumRetails
            Else
                oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
            End If
            
            'oDoc.ReadDataFromArticle Me.CDArticolo.KeyFieldID, sTabellaDettaglio
            
            oDoc.Field "Link_Art_articolo", Me.CDArticolo.KeyFieldID, sTabellaDettaglio
            oDoc.Field "Art_codice", Me.CDArticolo.Code, sTabellaDettaglio
            oDoc.Field "Art_descrizione", Me.txtDescrizioneArticolo.Text, sTabellaDettaglio
            oDoc.Field "Art_quantita_totale", Me.txtQta_UM.Value, sTabellaDettaglio
            oDoc.Field "Art_sco_in_percentuale_1", Me.txtSconto1.Value, sTabellaDettaglio
            oDoc.Field "Art_sco_in_percentuale_2", Me.txtSconto2.Value, sTabellaDettaglio
            oDoc.Field "Art_importo_sconto_netto_IVA", ((Me.txtImponibileUnitario.Value / 100) * (Me.txtSconto1.Value + Me.txtSconto2.Value)), sTabellaDettaglio
            oDoc.Field "Art_importo_totale_lordo_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) + (((Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_importo_totale_netto_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value), sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_netto_IVA", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_lordo_IVA", Me.txtImportoUnitarioArticolo.Value + ((Me.txtImportoUnitarioArticolo.Value / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_pre_uni_net_sco_net_IVA", Me.txtImponibileUnitario.Value, sTabellaDettaglio
            oDoc.Field "Art_pre_uni_net_sco_lor_IVA", Me.txtImponibileUnitario.Value + ((Me.txtImponibileUnitario.Value / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_Importo_totale_neutro", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value), sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_neutro", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
            oDoc.Field "Art_Importo_netto_IVA", Me.txtImponibileArticolo.Value, sTabellaDettaglio
            oDoc.Field "Art_importo_net_sconto_lor_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) + (((Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_importo_net_sconto_net_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value), sTabellaDettaglio
            oDoc.Field "Art_importo_sconto_lordo_IVA", ((Me.txtImponibileUnitario.Value / 100) * (Me.txtSconto1.Value + Me.txtSconto2.Value)), sTabellaDettaglio
            
            oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
            oDoc.Field "Link_art_IVA", Me.cboAliquotaArticolo.CurrentID, sTabellaDettaglio
            oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaArticolo.Value, sTabellaDettaglio
            
            oDoc.Field "Art_numero_colli", Me.txtColli.Value, sTabellaDettaglio
            oDoc.Field "Art_Peso", Me.txtPesoLordo.Value, sTabellaDettaglio
            oDoc.Field "Art_tara", Me.txtTara.Value, sTabellaDettaglio
            oDoc.Field "Art_quantita_pezzi", Me.txtPezzi.Value, sTabellaDettaglio
            oDoc.Field "Link_Art_unita_di_misura", Me.cboUnitaDiMisura.CurrentID, sTabellaDettaglio
            
            oDoc.Field "RV_POLinkRiga", NumeroRiga, sTabellaDettaglio
            oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneDocumento", "Rif. Ordine n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
            
            oDoc.Field "RV_POIDCalibro", Me.cboCalibro.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoCategoria", Me.cboTipoCategoria.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoLavorazione", Me.cboTipoLavorazione.CurrentID, sTabellaDettaglio

            oDoc.Field "RV_POImportoImballoInArticolo", Me.chkImportoImballoInArticolo.Value, sTabellaDettaglio
            
            
            oDoc.Field "RV_POIDArticoloPedana", Me.CDPedana.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POCodiceArticoloPedana", Me.CDPedana.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneArticoloPedana", Me.CDPedana.Description, sTabellaDettaglio
            oDoc.Field "RV_POQuantitaPedana", Me.txtQuantitaPedana.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoUMOrdine", cboUMRigaOrdine.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POTipoUMOrdine", cboUMRigaOrdine.Text, sTabellaDettaglio
            oDoc.Field "RV_POColliPerPedana", txtQuantitaPerPedana.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoPedana", Me.CDTipoPedana.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POCodiceTipoPedana", Me.CDTipoPedana.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneTipoPedana", Me.CDTipoPedana.Description, sTabellaDettaglio
            oDoc.Field "RV_POAnnotazioniRigaOrdine", Me.txtAnnotazioniDiRiga.Text, sTabellaDettaglio
            oDoc.Field "RV_POImportoUnitarioListino", IMPORTO_UNITARIO_LISTINO, sTabellaDettaglio
            
            oDoc.Field "RV_POIDConfigurazioneEtichettaLavorazione", Me.cboReportNew.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDConfigurazioneEtichettaPedana", Me.cboReportPedNew.CurrentID, sTabellaDettaglio
            
            
            'DETTAGLIO DELL'ORDINE VIVAIO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            oDoc.Field "RV_PO01_IDArticoloPianale", Me.CDArticoloPianale.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_PO01_QuantitaPianale", Me.txtQtaPianale.Value, sTabellaDettaglio
            
            oDoc.Field "RV_PO01_IDArticoloProlunga", Me.CDArticoloProlunga.KeyFieldID, sTabellaDettaglio
             oDoc.Field "RV_PO01_QuantitaProlunga", Me.txtQtaProlunga.Value, sTabellaDettaglio
            
            oDoc.Field "RV_PO01_IDAnagraficaFornitore", Me.ACSSocio.IDAnagrafica, sTabellaDettaglio
            ''''''''''''''''''''''''''''''''''''''      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            oDoc.Field "RV_PO01_AnnotazioniSocio", Me.txtAnnotazioniPerSocio.Text, sTabellaDettaglio
            
            oDoc.Field "RV_PONotaRigaOrdRaggr", Me.txtRaggrRigaOrdine.Text, sTabellaDettaglio
            oDoc.Field "RV_POQuantitaPedanaEffettiva", Me.txtQuantitaPedanaEff.Value, sTabellaDettaglio
            oDoc.Field "RV_POColliSfusi", Me.txtColliSfusi.Value, sTabellaDettaglio
            
            'IMBALLO PRIMARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            oDoc.Field "RV_POIDImballoPrimario", Me.CDImballoPrimario.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POTaraImballoPrimario", Me.txtTaraConfImballo.Value, sTabellaDettaglio
            oDoc.Field "RV_PONumeroConfezioniPerImballo", Me.txtNumeroConfImballo.Value, sTabellaDettaglio
            oDoc.Field "RV_POCodiceImballoPrimario", Me.CDImballoPrimario.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneImballoPrimario", Me.CDImballoPrimario.Description, sTabellaDettaglio
            oDoc.Field "RV_POQuantitaPerCollo", Me.txtQuantitaPerCollo.Value, sTabellaDettaglio
            oDoc.Field "RV_POPesoPerCollo", Me.txtPesoPerCollo.Value, sTabellaDettaglio
            oDoc.Field "RV_POMoltiplicatorePerCollo", Me.txtMoltiplicatore.Value, sTabellaDettaglio
            oDoc.Field "RV_POAnnotazioniRigaLavorazione", Me.txtAnnotazioniDiRigaLav, sTabellaDettaglio
            oDoc.Field "RV_POIDValoriOggettoDettaglioContratto", Me.txtIDRigaContratto.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDLottoCampagnaLavorazione", LINK_LOTTO_PROD_LAV, sTabellaDettaglio
            
            NumeroProgSingolaRiga = NumeroProgSingolaRiga + 1
            oDoc.Field "ID_Art_dettaglio_prog", NumeroProgSingolaRiga, sTabellaDettaglio
                        
            If Me.CDImballo.KeyFieldID > 0 Then
                oDoc.Field "RV_PORigaCompleta", 1, sTabellaDettaglio
                oDoc.Field "RV_POIDImballo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
                oDoc.Field "RV_POCodiceImballo", Me.CDImballo.Code, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneImballo", Me.txtDescrizioneImballo.Text, sTabellaDettaglio
                oDoc.Field "RV_POImportoUnitarioImballo", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
                oDoc.Field "RV_POTaraUnitariaImballo", Me.txtTaraUnitaria.Value, sTabellaDettaglio
                oDoc.Field "RV_POImportoImballoInArticolo", Abs(Me.chkImportoImballoInArticolo.Value), sTabellaDettaglio
            Else
                oDoc.Field "RV_PORigaCompleta", 0, sTabellaDettaglio
            End If
            
            If Me.CDImballo.KeyFieldID > 0 Then
                
                oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
                
                oDoc.Field "Link_Art_articolo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
                oDoc.Field "Art_codice", Me.CDImballo.Code, sTabellaDettaglio
                oDoc.Field "Art_descrizione", Me.txtDescrizioneImballo.Text, sTabellaDettaglio
                oDoc.Field "Art_quantita_totale", Me.txtColli.Value, sTabellaDettaglio
                    
                oDoc.Field "Art_prezzo_unitario_netto_IVA", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
                oDoc.Field "Art_prezzo_unitario_lordo_IVA", Me.txtImportoUnitarioImballo.Value + ((Me.txtImportoUnitarioImballo.Value / 100) * Me.txtAliquotaImballo.Value), sTabellaDettaglio
                    
                oDoc.Field "Art_pre_uni_net_sco_net_IVA", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
                oDoc.Field "Art_pre_uni_net_sco_lor_IVA", Me.txtImportoUnitarioImballo.Value + ((Me.txtImportoUnitarioImballo.Value / 100) * Me.txtAliquotaImballo.Value), sTabellaDettaglio
                    
                oDoc.Field "Art_Importo_totale_neutro", (Me.txtImportoUnitarioImballo.Value * Me.txtColli.Value), sTabellaDettaglio
                oDoc.Field "Art_prezzo_unitario_neutro", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
                    
                oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
                oDoc.Field "Art_Importo_netto_IVA", Me.txtImponibileImballo.Value, sTabellaDettaglio
                    
                oDoc.Field "Link_art_IVA", Me.CboAliquotaImballo.CurrentID, sTabellaDettaglio
                oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaImballo.Value, sTabellaDettaglio
                oDoc.Field "Art_tara", Me.txtTaraUnitaria.Value, sTabellaDettaglio
                oDoc.Field "Art_importo_totale_netto_IVA", Me.txtImponibileImballo.Value, sTabellaDettaglio
                    
    
                    
                oDoc.Field "Link_Art_unita_di_misura", 8, sTabellaDettaglio
                oDoc.Field "RV_POLinkRiga", NumeroRiga, sTabellaDettaglio
                oDoc.Field "RV_POTipoRiga", 2, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneDocumento", "Rif. ordine n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
                oDoc.Field "RV_PORigaCompleta", 1, sTabellaDettaglio
                oDoc.Field "RV_POIDImballo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
                NumeroProgSingolaRiga = NumeroProgSingolaRiga + 1
                oDoc.Field "ID_Art_dettaglio_prog", NumeroProgSingolaRiga, sTabellaDettaglio
                
                oDoc.Field "RV_POIDCalibro", Me.cboCalibro.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoCategoria", Me.cboTipoCategoria.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoLavorazione", Me.cboTipoLavorazione.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POImportoImballoInArticolo", Me.chkImportoImballoInArticolo.Value, sTabellaDettaglio
                
                oDoc.Field "RV_POIDArticoloPedana", Me.CDPedana.KeyFieldID, sTabellaDettaglio
                oDoc.Field "RV_POCodiceArticoloPedana", Me.CDPedana.Code, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneArticoloPedana", Me.CDPedana.Description, sTabellaDettaglio
                oDoc.Field "RV_POQuantitaPedana", Me.txtQuantitaPedana.Value, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoUMOrdine", cboUMRigaOrdine.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POTipoUMOrdine", cboUMRigaOrdine.Text, sTabellaDettaglio
                oDoc.Field "RV_POColliPerPedana", txtQuantitaPerCollo.Value, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoPedana", Me.CDTipoPedana.KeyFieldID, sTabellaDettaglio
                oDoc.Field "RV_POCodiceTipoPedana", Me.CDTipoPedana.Code, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneTipoPedana", Me.CDTipoPedana.Description, sTabellaDettaglio
                oDoc.Field "RV_POAnnotazioniRigaOrdine", Me.txtAnnotazioniDiRiga.Text, sTabellaDettaglio
               
            End If
    Else
        If Me.CDImballo.KeyFieldID > 0 Then
            If oDoc.Tables(sTabellaDettaglio).NumRetails = 0 Then
                oDoc.Tables(sTabellaDettaglio).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglio).NumRetails
            Else
                oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
            End If

                
            oDoc.Field "Link_Art_articolo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
            oDoc.Field "Art_codice", Me.CDImballo.Code, sTabellaDettaglio
            oDoc.Field "Art_descrizione", Me.txtDescrizioneImballo.Text, sTabellaDettaglio
            oDoc.Field "Art_quantita_totale", Me.txtColli.Value, sTabellaDettaglio
                    
            oDoc.Field "Art_prezzo_unitario_netto_IVA", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_lordo_IVA", Me.txtImportoUnitarioImballo.Value + ((Me.txtImportoUnitarioImballo.Value / 100) * Me.txtAliquotaImballo.Value), sTabellaDettaglio
                    
            oDoc.Field "Art_pre_uni_net_sco_net_IVA", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
            oDoc.Field "Art_pre_uni_net_sco_lor_IVA", Me.txtImportoUnitarioImballo.Value + ((Me.txtImportoUnitarioImballo.Value / 100) * Me.txtAliquotaImballo.Value), sTabellaDettaglio
                    
            'oDoc.Field "Art_Importo_totale_neutro", Me.txtTotaleImballo.Value, sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_neutro", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
                    
            oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
            oDoc.Field "Art_Importo_netto_IVA", Me.txtImponibileImballo.Value, sTabellaDettaglio
            oDoc.Field "Art_tara", Me.txtTaraUnitaria.Value, sTabellaDettaglio
            oDoc.Field "Link_art_IVA", Me.CboAliquotaImballo.CurrentID, sTabellaDettaglio
            oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaImballo.Value, sTabellaDettaglio
                    
            oDoc.Field "Art_importo_totale_netto_IVA", Me.txtImponibileImballo.Value, sTabellaDettaglio
                    
            oDoc.Field "Link_Art_unita_di_misura", 8, sTabellaDettaglio
            
            
            oDoc.Field "RV_POLinkRiga", NumeroRiga, sTabellaDettaglio
            oDoc.Field "RV_POTipoRiga", 2, sTabellaDettaglio
            oDoc.Field "RV_PORigaCompleta", 1, sTabellaDettaglio
            oDoc.Field "RV_POIDImballo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneDocumento", "Rif. ordine n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
                
            oDoc.Field "RV_PORigaCompleta", 0, sTabellaDettaglio
            NumeroProgSingolaRiga = NumeroProgSingolaRiga + 1
            oDoc.Field "ID_Art_dettaglio_prog", NumeroProgSingolaRiga, sTabellaDettaglio
            
            oDoc.Field "RV_POIDCalibro", Me.cboCalibro.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoCategoria", Me.cboTipoCategoria.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoLavorazione", Me.cboTipoLavorazione.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POImportoImballoInArticolo", Me.chkImportoImballoInArticolo.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDArticoloPedana", Me.CDPedana.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POCodiceArticoloPedana", Me.CDPedana.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneArticoloPedana", Me.CDPedana.Description, sTabellaDettaglio
            oDoc.Field "RV_POQuantitaPedana", Me.txtQuantitaPedana.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoUMOrdine", cboUMRigaOrdine.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POTipoUMOrdine", cboUMRigaOrdine.Text, sTabellaDettaglio
            oDoc.Field "RV_POColliPerPedana", txtQuantitaPerCollo.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoPedana", Me.CDTipoPedana.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POCodiceTipoPedana", Me.CDTipoPedana.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneTipoPedana", Me.CDTipoPedana.Description, sTabellaDettaglio
            oDoc.Field "RV_POAnnotazioniRigaOrdine", Me.txtAnnotazioniDiRiga.Text, sTabellaDettaglio

        Else
            If (Me.txtDescrizioneArticolo.Text <> "") Then
                If oDoc.Tables(sTabellaDettaglio).NumRetails = 0 Then
                    oDoc.Tables(sTabellaDettaglio).SetActiveRetail 1 'oDoc.Tables(sTabellaDettaglio).NumRetails
                Else
                    oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
                End If
        
                oDoc.Field "Link_Art_articolo", 0, sTabellaDettaglio
                oDoc.Field "Art_codice", "", sTabellaDettaglio
                oDoc.Field "Art_descrizione", Me.txtDescrizioneArticolo.Text, sTabellaDettaglio
                oDoc.Field "Art_quantita_totale", Me.txtQta_UM.Value, sTabellaDettaglio
                oDoc.Field "Art_prezzo_unitario_netto_IVA", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
                oDoc.Field "Art_prezzo_unitario_lordo_IVA", Me.txtImportoUnitarioArticolo.Value + ((Me.txtImportoUnitarioArticolo.Value / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
                oDoc.Field "Art_pre_uni_net_sco_net_IVA", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
                oDoc.Field "Art_pre_uni_net_sco_lor_IVA", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
                'oDoc.Field "Art_Importo_totale_neutro", Me.txtTotaleArticolo.Value, sTabellaDettaglio
                oDoc.Field "Art_prezzo_unitario_neutro", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
                oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
                oDoc.Field "Art_Importo_netto_IVA", Me.txtImponibileArticolo.Value, sTabellaDettaglio
                oDoc.Field "Link_art_IVA", Me.cboAliquotaArticolo.CurrentID, sTabellaDettaglio
                oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaArticolo.Value, sTabellaDettaglio
                
                oDoc.Field "Art_importo_totale_netto_IVA", Me.txtImponibileArticolo.Value, sTabellaDettaglio
                oDoc.Field "Art_numero_colli", Me.txtColli.Value, sTabellaDettaglio
                oDoc.Field "Art_Peso", Me.txtPesoNetto.Value, sTabellaDettaglio
                oDoc.Field "Art_tara", Me.txtTara.Value, sTabellaDettaglio
                oDoc.Field "Art_quantita_pezzi", Me.txtPezzi.Value, sTabellaDettaglio
                oDoc.Field "Link_Art_unita_di_misura", Me.cboUnitaDiMisura.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POLinkRiga", NumeroRiga, sTabellaDettaglio
                oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
                oDoc.Field "RV_PORigaCompleta", 1, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneDocumento", "Rif. ordine n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
                
                oDoc.Field "RV_PORigaCompleta", 0, sTabellaDettaglio
                NumeroProgSingolaRiga = NumeroProgSingolaRiga + 1
                oDoc.Field "ID_Art_dettaglio_prog", NumeroProgSingolaRiga, sTabellaDettaglio
                
                oDoc.Field "RV_POIDCalibro", Me.cboCalibro.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoCategoria", Me.cboTipoCategoria.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoLavorazione", Me.cboTipoLavorazione.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POImportoImballoInArticolo", Me.chkImportoImballoInArticolo.Value, sTabellaDettaglio
                oDoc.Field "RV_POIDArticoloPedana", Me.CDPedana.KeyFieldID, sTabellaDettaglio
                oDoc.Field "RV_POCodiceArticoloPedana", Me.CDPedana.Code, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneArticoloPedana", Me.CDPedana.Description, sTabellaDettaglio
                oDoc.Field "RV_POQuantitaPedana", Me.txtQuantitaPedana.Value, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoUMOrdine", cboUMRigaOrdine.CurrentID, sTabellaDettaglio
                oDoc.Field "RV_POTipoUMOrdine", cboUMRigaOrdine.Text, sTabellaDettaglio
                oDoc.Field "RV_POColliPerPedana", txtQuantitaPerCollo.Value, sTabellaDettaglio
                oDoc.Field "RV_POIDTipoPedana", Me.CDTipoPedana.KeyFieldID, sTabellaDettaglio
                oDoc.Field "RV_POCodiceTipoPedana", Me.CDTipoPedana.Code, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneTipoPedana", Me.CDTipoPedana.Description, sTabellaDettaglio
                oDoc.Field "RV_POAnnotazioniRigaOrdine", Me.txtAnnotazioniDiRiga.Text, sTabellaDettaglio

            Else
                
            End If
            
        
        End If
    
    
    End If
    
    
    
    
End Sub
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
Private Sub ModificaRigaDocumento()
Dim ID As Long



    If A_Riga(0) > 0 Then
            
            oDoc.Tables(sTabellaDettaglio).SetActiveRetail A_Riga(0)
'            ID = fnNotNullL(oDoc.Field("IDValoriOggettoDettaglio", -ID, sTabellaDettaglio))
'            If ID > 0 Then
'                oDoc.Field "IDValoriOggettoDettaglio", -ID, sTabellaDettaglio
'            End If
            
            oDoc.Field "Link_Art_articolo", Me.CDArticolo.KeyFieldID, sTabellaDettaglio
            oDoc.Field "Art_codice", Me.CDArticolo.Code, sTabellaDettaglio
            oDoc.Field "Art_descrizione", Me.txtDescrizioneArticolo.Text, sTabellaDettaglio
            oDoc.Field "Art_quantita_totale", Me.txtQta_UM.Value, sTabellaDettaglio
            oDoc.Field "Art_sco_in_percentuale_1", Me.txtSconto1.Value, sTabellaDettaglio
            oDoc.Field "Art_sco_in_percentuale_2", Me.txtSconto2.Value, sTabellaDettaglio
            oDoc.Field "Art_importo_sconto_netto_IVA", ((Me.txtImponibileUnitario.Value / 100) * (Me.txtSconto1.Value + Me.txtSconto2.Value)), sTabellaDettaglio
            oDoc.Field "Art_importo_totale_lordo_IVA", (Me.txtImponibileUnitario * Me.txtQta_UM) + (((Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_importo_totale_netto_IVA", (Me.txtImponibileUnitario * Me.txtQta_UM), sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_netto_IVA", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_lordo_IVA", Me.txtImportoUnitarioArticolo.Value + ((Me.txtImportoUnitarioArticolo.Value / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_pre_uni_net_sco_net_IVA", Me.txtImponibileUnitario.Value, sTabellaDettaglio
            oDoc.Field "Art_pre_uni_net_sco_lor_IVA", Me.txtImponibileUnitario.Value + ((Me.txtImponibileUnitario.Value / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_Importo_totale_neutro", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value), sTabellaDettaglio
            oDoc.Field "Art_prezzo_unitario_neutro", Me.txtImportoUnitarioArticolo.Value, sTabellaDettaglio
            oDoc.Field "Art_Importo_netto_IVA", Me.txtImponibileArticolo.Value, sTabellaDettaglio
            oDoc.Field "Art_importo_net_sconto_lor_IVA", (Me.txtImponibileUnitario * Me.txtQta_UM) + (((Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
            oDoc.Field "Art_importo_net_sconto_net_IVA", (Me.txtImponibileUnitario * Me.txtQta_UM), sTabellaDettaglio
            oDoc.Field "Art_importo_sconto_lordo_IVA", ((Me.txtImponibileUnitario.Value / 100) * (Me.txtSconto1.Value + Me.txtSconto2.Value)), sTabellaDettaglio
            
            oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
            oDoc.Field "Link_art_IVA", Me.cboAliquotaArticolo.CurrentID, sTabellaDettaglio
            oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaArticolo.Value, sTabellaDettaglio

            oDoc.Field "Art_numero_colli", Me.txtColli.Value, sTabellaDettaglio
            oDoc.Field "Art_Peso", Me.txtPesoLordo.Value, sTabellaDettaglio
            oDoc.Field "Art_tara", Me.txtTara.Value, sTabellaDettaglio
            oDoc.Field "Art_quantita_pezzi", Me.txtPezzi.Value, sTabellaDettaglio
            oDoc.Field "Link_Art_unita_di_misura", Me.cboUnitaDiMisura.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneDocumento", "Rif. ordine n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
            
            oDoc.Field "RV_POIDCalibro", Me.cboCalibro.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoCategoria", Me.cboTipoCategoria.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoLavorazione", Me.cboTipoLavorazione.CurrentID, sTabellaDettaglio

            oDoc.Field "RV_POImportoImballoInArticolo", Me.chkImportoImballoInArticolo.Value, sTabellaDettaglio

            oDoc.Field "RV_POIDArticoloPedana", Me.CDPedana.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POCodiceArticoloPedana", Me.CDPedana.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneArticoloPedana", Me.CDPedana.Description, sTabellaDettaglio
            oDoc.Field "RV_POQuantitaPedana", Me.txtQuantitaPedana.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoUMOrdine", cboUMRigaOrdine.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POColliPerPedana", Me.txtQuantitaPerPedana.Value, sTabellaDettaglio
            oDoc.Field "RV_POIDTipoPedana", Me.CDTipoPedana.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POCodiceTipoPedana", Me.CDTipoPedana.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneTipoPedana", Me.CDTipoPedana.Description, sTabellaDettaglio
            oDoc.Field "RV_POAnnotazioniRigaOrdine", Me.txtAnnotazioniDiRiga.Text, sTabellaDettaglio

            oDoc.Field "RV_POIDConfigurazioneEtichettaLavorazione", Me.cboReportNew.CurrentID, sTabellaDettaglio
            oDoc.Field "RV_POIDConfigurazioneEtichettaPedana", Me.cboReportPedNew.CurrentID, sTabellaDettaglio

            'DETTAGLIO DELL'ORDINE VIVAIO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            oDoc.Field "RV_PO01_IDArticoloPianale", Me.CDArticoloPianale.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_PO01_QuantitaPianale", Me.txtQtaPianale.Value, sTabellaDettaglio
            
            oDoc.Field "RV_PO01_IDArticoloProlunga", Me.CDArticoloProlunga.KeyFieldID, sTabellaDettaglio
             oDoc.Field "RV_PO01_QuantitaProlunga", Me.txtQtaProlunga.Value, sTabellaDettaglio
            
            oDoc.Field "RV_PO01_IDAnagraficaFornitore", Me.ACSSocio.IDAnagrafica, sTabellaDettaglio
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            oDoc.Field "RV_PO01_AnnotazioniSocio", Me.txtAnnotazioniPerSocio.Text, sTabellaDettaglio
            
            'IMBALLO PRIMARIO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            oDoc.Field "RV_POIDImballoPrimario", Me.CDImballoPrimario.KeyFieldID, sTabellaDettaglio
            oDoc.Field "RV_POTaraImballoPrimario", Me.txtTaraConfImballo.Value, sTabellaDettaglio
            oDoc.Field "RV_PONumeroConfezioniPerImballo", Me.txtNumeroConfImballo.Value, sTabellaDettaglio
            oDoc.Field "RV_POCodiceImballoPrimario", Me.CDImballoPrimario.Code, sTabellaDettaglio
            oDoc.Field "RV_PODescrizioneImballoPrimario", Me.CDImballoPrimario.Description, sTabellaDettaglio
            oDoc.Field "RV_POQuantitaPerCollo", Me.txtQuantitaPerCollo.Value, sTabellaDettaglio
            oDoc.Field "RV_POPesoPerCollo", Me.txtPesoPerCollo.Value, sTabellaDettaglio
            oDoc.Field "RV_POMoltiplicatorePerCollo", Me.txtMoltiplicatore.Value, sTabellaDettaglio
            oDoc.Field "RV_POAnnotazioniRigaLavorazione", Me.txtAnnotazioniDiRigaLav, sTabellaDettaglio
            oDoc.Field "RV_POIDValoriOggettoDettaglioContratto", Me.txtIDRigaContratto.Value, sTabellaDettaglio
            
            oDoc.Field "RV_PONotaRigaOrdRaggr", Me.txtRaggrRigaOrdine.Text, sTabellaDettaglio
            oDoc.Field "RV_POQuantitaPedanaEffettiva", Me.txtQuantitaPedanaEff.Value, sTabellaDettaglio
            oDoc.Field "RV_POColliSfusi", Me.txtColliSfusi.Value, sTabellaDettaglio
            
            oDoc.Field "TipoUMOrdine", cboUMRigaOrdine.Text, sTabellaDettaglio
            oDoc.Field "RV_POIDLottoCampagnaLavorazione", LINK_LOTTO_PROD_LAV, sTabellaDettaglio
            
            If A_Riga(1) > 0 Then
                oDoc.Field "RV_PORigaCompleta", 1, sTabellaDettaglio
                oDoc.Field "RV_POIDImballo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
                oDoc.Field "RV_POCodiceImballo", Me.CDImballo.Code, sTabellaDettaglio
                oDoc.Field "RV_PODescrizioneImballo", Me.txtDescrizioneImballo.Text, sTabellaDettaglio
                oDoc.Field "RV_POImportoUnitarioImballo", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
                oDoc.Field "RV_POTaraUnitariaImballo", Me.txtTaraUnitaria.Value, sTabellaDettaglio
                oDoc.Field "RV_POImportoImballoInArticolo", Abs(Me.chkImportoImballoInArticolo.Value), sTabellaDettaglio
            End If
            
    End If
    If A_Riga(1) > 0 Then
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail A_Riga(1)
'        ID = fnNotNullL(oDoc.Field("IDValoriOggettoDettaglio", -ID, sTabellaDettaglio))
'        If ID > 0 Then
'            oDoc.Field "IDValoriOggettoDettaglio", -ID, sTabellaDettaglio
'        End If
                
        oDoc.Field "Link_Art_articolo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
        oDoc.Field "Art_codice", Me.CDImballo.Code, sTabellaDettaglio
        oDoc.Field "Art_descrizione", Me.txtDescrizioneImballo.Text, sTabellaDettaglio
        oDoc.Field "Art_quantita_totale", Me.txtColli.Value, sTabellaDettaglio
        oDoc.Field "Art_sco_in_percentuale_1", 0, sTabellaDettaglio
        oDoc.Field "Art_sco_in_percentuale_2", 0, sTabellaDettaglio
        oDoc.Field "Art_importo_sconto_netto_IVA", 0, sTabellaDettaglio
        oDoc.Field "Art_importo_totale_lordo_IVA", (Me.txtImportoUnitarioImballo.Value * Me.txtColli) + (((Me.txtImportoUnitarioImballo.Value * Me.txtColli.Value) / 100) * Me.txtAliquotaImballo.Value), sTabellaDettaglio
        oDoc.Field "Art_importo_totale_netto_IVA", (Me.txtImportoUnitarioImballo.Value * Me.txtColli), sTabellaDettaglio
                    
        oDoc.Field "Art_prezzo_unitario_netto_IVA", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
        oDoc.Field "Art_prezzo_unitario_lordo_IVA", Me.txtImportoUnitarioImballo.Value + ((Me.txtImportoUnitarioImballo.Value / 100) * Me.txtAliquotaImballo.Value), sTabellaDettaglio
        
        oDoc.Field "Art_pre_uni_net_sco_net_IVA", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
        oDoc.Field "Art_pre_uni_net_sco_lor_IVA", Me.txtImportoUnitarioImballo.Value + ((Me.txtImportoUnitarioImballo.Value / 100) * Me.txtAliquotaImballo.Value), sTabellaDettaglio
                    
        oDoc.Field "Art_prezzo_unitario_neutro", Me.txtImportoUnitarioImballo.Value, sTabellaDettaglio
        oDoc.Field "Art_importo_totale_neutro", (Me.txtImportoUnitarioImballo.Value * Me.txtColli.Value), sTabellaDettaglio
        
        oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
        oDoc.Field "Art_Importo_netto_IVA", Me.txtImponibileImballo.Value, sTabellaDettaglio
                    
        oDoc.Field "Link_art_IVA", Me.CboAliquotaImballo.CurrentID, sTabellaDettaglio
        oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaImballo.Value, sTabellaDettaglio
        oDoc.Field "Art_importo_totale_netto_IVA", Me.txtImponibileImballo.Value, sTabellaDettaglio
        oDoc.Field "RV_PODescrizioneDocumento", "Rif. D.d.t. n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
        oDoc.Field "RV_POIDImballo", Me.CDImballo.KeyFieldID, sTabellaDettaglio
        oDoc.Field "Link_Art_unita_di_misura", 8, sTabellaDettaglio
        
        oDoc.Field "RV_POIDCalibro", Me.cboCalibro.CurrentID, sTabellaDettaglio
        oDoc.Field "RV_POIDTipoCategoria", Me.cboTipoCategoria.CurrentID, sTabellaDettaglio
        oDoc.Field "RV_POIDTipoLavorazione", Me.cboTipoLavorazione.CurrentID, sTabellaDettaglio
        oDoc.Field "RV_POImportoImballoInArticolo", Me.chkImportoImballoInArticolo.Value, sTabellaDettaglio
        oDoc.Field "RV_POIDArticoloPedana", Me.CDPedana.KeyFieldID, sTabellaDettaglio
        oDoc.Field "RV_POCodiceArticoloPedana", Me.CDPedana.Code, sTabellaDettaglio
        oDoc.Field "RV_PODescrizioneArticoloPedana", Me.CDPedana.Description, sTabellaDettaglio
        oDoc.Field "RV_POQuantitaPedana", Me.txtQuantitaPedana.Value, sTabellaDettaglio
        oDoc.Field "RV_POIDTipoUMOrdine", cboUMRigaOrdine.CurrentID, sTabellaDettaglio
        oDoc.Field "TipoUMOrdine", cboUMRigaOrdine.Text, sTabellaDettaglio
        oDoc.Field "RV_POColliPerPedana", Me.txtQuantitaPerCollo.Value, sTabellaDettaglio
        oDoc.Field "RV_POIDTipoPedana", Me.CDTipoPedana.KeyFieldID, sTabellaDettaglio
        oDoc.Field "RV_POCodiceTipoPedana", Me.CDTipoPedana.Code, sTabellaDettaglio
        oDoc.Field "RV_PODescrizioneTipoPedana", Me.CDTipoPedana.Description, sTabellaDettaglio
        oDoc.Field "RV_POAnnotazioniRigaOrdine", Me.txtAnnotazioniDiRiga.Text, sTabellaDettaglio

    End If
End Sub

Private Sub txtPesoNetto_Change()
    If Link_UMCoop = 3 Then
        Me.txtQta_UM.Value = Me.txtPesoNetto.Value
    End If

End Sub

Private Sub txtPesoNetto_LostFocus()
    Me.txtPesoLordo.Value = Me.txtPesoNetto.Value + Me.txtTara.Value
    CalcoloPesoNetto
End Sub

Private Sub txtPesoPerCollo_LostFocus()
    txtColli_LostFocus
End Sub

Private Sub txtPesoTotale_Change()
    sbImpostaDatiDocumento
End Sub

Private Sub fnGetDefaultPerSezionale()
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT IDSezionale FROM DefaultFilialePerTipoOggetto "
    sSQL = sSQL & "WHERE ((IDFiliale=" & oDoc.IDFiliale & ") "
    sSQL = sSQL & "AND (IDTipoOggetto=" & oDoc.IDTipoOggetto & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        Me.cboSezionale.WriteOn 0
    Else
        Me.cboSezionale.WriteOn rs!IDSezionale
    End If
    
    
    rs.CloseResultset
    Set rs = Nothing
    
    
End Sub
Public Function fnGetParametriMagazzino(NomeCampo As String) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetParametriMagazzino = rsEse.adoColumns(NomeCampo).Value
    Else
        sSQL = "SELECT " & NomeCampo & " FROM RV_POSchemaCoop "
        sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
        
        Set rsEse = Cn.OpenResultset(sSQL)
        
        If rsEse.EOF = False Then
            fnGetParametriMagazzino = rsEse.adoColumns(NomeCampo).Value
        Else
            fnGetParametriMagazzino = 0
        End If
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
Private Sub AggiornaAltreDestinazioni()
    With Me.cboAltroSito
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDSitoPerAnagrafica"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "SitoPerAnagrafica"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDSitoPerAnagrafica, SitoPerAnagrafica FROM SitoPerAnagrafica"
        .SQL = .SQL & " WHERE IDAnagrafica = " & Me.cdAnagrafica.KeyFieldID
        .SQL = .SQL & " ORDER BY SitoPerAnagrafica"
    End With

End Sub
Private Sub AggiornaContrattiBancariCliente()
    With Me.cboBancaCliente
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDBancaPerAnagrafica"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "BancaPerAnagrafica"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDBancaPerAnagrafica, BancaPerAnagrafica FROM BancaPerAnagrafica "
        .SQL = .SQL & " WHERE IDAnagrafica = " & Me.cdAnagrafica.KeyFieldID
        .SQL = .SQL & " ORDER BY BancaPerAnagrafica"
    End With

End Sub

Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POIDUnitaDiMisuraCoop FROM UnitaDiMisura WHERE "
    sSQL = sSQL & "IDUnitaDiMisura = " & Link_UMAcq
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetUMCoop = rs!RV_POIDUnitaDiMisuraCoop
    Else
        fnGetUMCoop = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub txtPezzi_Change()
    If Link_UMCoop = 5 Then
        Me.txtQta_UM.Value = CLng(Me.txtPezzi.Value)
    End If
    
End Sub



Private Sub txtQta_UM_Change()
    
    If GESTIONE_ORDINE_VIVAIO = 1 Then
        GET_CALCOLO_VIVAIO txtQta_UM.Value
    End If
    
    If bVariazioneDettaglio = True Then Exit Sub
    If DATI_DA_CONTRATTO = True Then Exit Sub
    
    GET_CONFIGURAZIONE_IMPORTI Me.cdAnagrafica.KeyFieldID, Me.CDArticolo.KeyFieldID, fnNotNullN(oDoc.Field("Link_Doc_listino", , sTabellaTestata)), fnNotNullN(oDoc.Field("Link_Doc_listino_base", , sTabellaTestata)), Me.txtQta_UM.Value

End Sub

Private Sub txtQuantitaPedana_LostFocus()
    If (GESTIONE_ORDINE_VIVAIO = 0) Then
        'If bVariazioneDettaglio = False Then
            'Me.txtColli.Value = (Me.txtQuantitaPedana.Value * GET_QUANTITA_IMBALLO_PER_TIPO_PEDANA(Me.CDTipoPedana.KeyFieldID, Me.CDImballo.KeyFieldID)) + Me.txtColliSfusi.Value
            Me.txtColli.Value = Me.txtQuantitaPedana.Value * Me.txtQuantitaPerPedana.Value + Me.txtColliSfusi.Value
            txtColli_LostFocus
        'End If
    Else
        Me.txtQuantitaPedanaEff.Value = Me.txtQuantitaPedana.Value
    End If
End Sub

Private Sub txtQuantitaPerCollo_LostFocus()
    txtColli_LostFocus
End Sub

Private Sub txtQuantitaPerPedana_LostFocus()
    Me.txtColli.Value = (Me.txtQuantitaPedana.Value * Me.txtQuantitaPerPedana.Value) + Me.txtColliSfusi.Value
    txtColli_LostFocus
End Sub

Private Sub txtRaggrRigaOrdine_DblClick()
    If (ATTIVA_SEL_LOTTO_PROD_IN_LAV = 1) Then
        frmSelezionaLottoDiCampagna.Show vbModal
    End If
End Sub

Private Sub txtRaggrRigaOrdine_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_txtRaggrRigaOrdine_KeyDown
Dim Testo As String
    If LINK_LOTTO_PROD_LAV > 0 Then
        If ((KeyCode = vbKeyDelete) Or (KeyCode = vbKeyBack)) Then
            If ((ATTIVA_SEL_LOTTO_PROD_IN_LAV = 1)) Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Sei sicuro di voler eliminare il riferimento del sublotto?"
                If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento sublotto") = vbNo Then Exit Sub
                LINK_LOTTO_PROD_LAV = 0
                frmMain.txtRaggrRigaOrdine.Text = frmMain.GET_DESCRIZIONE_LOTTO_PROD_LAV(LINK_LOTTO_PROD_LAV)
            End If
        End If
    End If
Exit Sub
ERR_txtRaggrRigaOrdine_KeyDown:
    MsgBox Err.Description, vbCritical, "txtRaggrRigaOrdine_KeyDown"
End Sub

Private Sub txtSconto1_Change()
    CalcolaImportoScontato
    CalcolaTotaleRiga
End Sub
Private Sub txtSconto2_Change()
    CalcolaImportoScontato
    CalcolaTotaleRiga
End Sub
Private Sub txtTara_Change()
    If Link_UMCoop = 4 Then
        Me.txtQta_UM.Value = Me.txtTara.Value
    End If
End Sub
Private Function DispLotto() As Double

End Function



Private Sub fnEliminaDatiTemporanei()
End Sub
Private Sub ControlloChiusuraLotto(IDLotto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POCollegamento "
sSQL = sSQL & "WHERE IDlottoArticolo=" & IDLotto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    If rs!QtaColliCaricati - rs!QtaColliVenduti <= 0 Then
        sSQL = "UPDATE RV_POLavorazione SET "
        sSQL = sSQL & "Chiuso=" & fnNormBoolean(1) & " "
        sSQL = sSQL & "WHERE IDCodiceLotto_Vendita=" & IDLotto
    Else
        sSQL = "UPDATE RV_POLavorazione SET "
        sSQL = sSQL & "Chiuso=" & fnNormBoolean(0) & " "
        sSQL = sSQL & "WHERE IDCodiceLotto_Vendita=" & IDLotto
    End If
    
    Cn.Execute sSQL
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub fnSetCausaleDocumento()
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    sSQL = "SELECT CausaleTrasporto FROM CausaleTrasportoPerFunzione "
    sSQL = sSQL & "WHERE IDFunzione=" & oDoc.IDFunzione
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        Me.txtCausaleDocumento.Text = ""
    Else
        Me.txtCausaleDocumento.Text = fnNotNull(rs!CausaleTrasporto)
    End If
    
    
    rs.CloseResultset
    Set rs = Nothing
    
End Sub

Private Sub fnSetProtocolloICE()
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset

End Sub
Private Sub AggiornaProtocolloICE(OperazioneDocumento As Integer)
End Sub

Private Sub txtTara_LostFocus()
    CalcoloPesoNetto
End Sub
Private Sub fnDeleteTabellaRicorsione(IDUtente As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fnDeleteTabellaRicorsione
Dim sSQL As String
    
    sSQL = "DELETE FROM TabellaRicorsione "
    sSQL = sSQL & "WHERE IDUtente=" & IDUtente
    'sSQL = sSQL & " WHERE IDFunzione=" & IDTipoOggetto
    Cn.Execute sSQL
    
    sSQL = "DELETE FROM TabellaRicorsione2 "
    'sSQL = sSQL & "WHERE IDPath1=" & IDTipoOggetto
    sSQL = sSQL & "WHERE IDUtente=" & IDUtente
    Cn.Execute sSQL

Exit Sub
ERR_fnDeleteTabellaRicorsione:
    MsgBox Err.Description, vbCritical, "Cancellazione tabella ricorsione"
End Sub
Private Sub BrwMain_ConditionEdit(ByVal Name As String, Value As Variant)
    Dim oSearch As dmtFind.Find
    Dim sSQL As String
    Dim oRes As DmtOleDbLib.adoResultset

    'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = TheApp.Database.Connection

If Name = "Anagrafica" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Anagrafica"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Cliente.IDAzienda "
    sSQL = sSQL & "FROM Anagrafica INNER JOIN "
    sSQL = sSQL & "Cliente ON Anagrafica.IDAnagrafica = Cliente.IDAnagrafica "
    sSQL = sSQL & "WHERE (Cliente.IDAzienda =" & TheApp.IDFirm & ") "
    sSQL = sSQL & "ORDER BY Anagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Anagrafica")
                
    End If
    oRes.CloseResultset
    Set oRes = Nothing
End If
If Name = "Destinazione diversa" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Destinazioni diverse"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Destinazione diversa", "SitoPerAnagrafica", 1   'STRINGTYPE
    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1   'STRINGTYPE
    oSearch.AddDisplayField "Codice", "Codice", 1   'STRINGTYPE
        
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT SitoPerAnagrafica.IDSitoPerAnagrafica, SitoPerAnagrafica.SitoPerAnagrafica, Cliente.Codice, Anagrafica.Anagrafica, Anagrafica.Nome, "
    sSQL = sSQL & "Cliente.IDAzienda "
    sSQL = sSQL & "FROM SitoPerAnagrafica INNER JOIN "
    sSQL = sSQL & "Cliente ON SitoPerAnagrafica.IDAnagrafica = Cliente.IDAnagrafica INNER JOIN "
    sSQL = sSQL & "Anagrafica ON Cliente.IDAnagrafica = Anagrafica.IDAnagrafica "
    sSQL = sSQL & "WHERE Cliente.IDAzienda =" & TheApp.IDFirm
    If Len(BrwMain.Conditions("Nom_ragione_sociale_o_cognome").FromValue) > 0 Then
        sSQL = sSQL & " AND Anagrafica LIKE" & fnNormString(BrwMain.Conditions("Nom_ragione_sociale_o_cognome").FromValue & "%")
    End If
    
    sSQL = sSQL & " ORDER BY Anagrafica, SitoPerAnagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("SitoPerAnagrafica")
                
    End If
    oRes.CloseResultset
    Set oRes = Nothing
End If
End Sub
Private Function DeleteAll() As Boolean
On Error GoTo ERR_DeleteAll
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCol As DmtOleDbLib.adoResultset

sSQL = "SELECT IDLottoArticolo, SUM(QtaColliPrec) AS QtaColliPrec "
sSQL = sSQL & "FROM RV_POTMPQtaLottoDocumento "
sSQL = sSQL & "GROUP BY IDLottoArticolo "


Set rs = Cn.OpenResultset(sSQL)

'Cn.BeginTrans
If rs.EOF Then
    DeleteAll = True
Else
    While Not rs.EOF
        sSQL = "SELECT QtaColliVenduti FROM RV_POCollegamento "
        sSQL = sSQL & "WHERE IDLottoArticolo=" & fnNotNullN(rs!IDLottoArticolo)
        
        Set rsCol = Cn.OpenResultset(sSQL)
        If rsCol.EOF = False Then
            sSQL = "UPDATE RV_POCollegamento SET "
            sSQL = sSQL & "QtaColliVenduti=" & fnNormNumber(fnNotNullN(rsCol!QtaColliVenduti) - fnNotNullN(rs!QtaColliPrec)) & " "
            sSQL = sSQL & "WHERE IDLottoArticolo=" & fnNotNullN(rs!IDLottoArticolo)
            
            Cn.Execute sSQL
        End If
        
        
        rsCol.CloseResultset
        Set rsCol = Nothing
        
        ControlloChiusuraLotto fnNotNullN(rs!IDLottoArticolo)
    rs.MoveNext
    Wend
    
    
    
'    Cn.CommitTrans
    
    DeleteAll = True
    
    
End If
rs.CloseResultset
Set rs = Nothing



Exit Function
ERR_DeleteAll:


    MsgBox Err.Description
    DeleteAll = False
End Function

Private Function GetAttivitaAzienda(IDAzienda As Long, IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivitaAzienda.IDAttivitaAzienda, Azienda.IDAzienda, Filiale.IDFiliale "
sSQL = sSQL & "FROM AttivitaAzienda INNER JOIN "
sSQL = sSQL & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda INNER JOIN "
sSQL = sSQL & "Filiale ON AttivitaAzienda.IDAttivitaAzienda = Filiale.IDAttivitaAzienda "
sSQL = sSQL & "Where (Azienda.IDAzienda =" & IDAzienda & ") And (Filiale.IDFiliale = " & IDFiliale & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GetAttivitaAzienda = 0
Else
    GetAttivitaAzienda = fnNotNullN(rs!IDAttivitaAzienda)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_STRINGA_CONFERIMENTO(IDRigaConf As Long, IDAssegnazioneMerce As Long, IDProcessoIVGamma As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Testo As String
    'PROCESSO NORMALE
    If ((IDRigaConf > 0) And (IDAssegnazioneMerce = 0)) Then
        sSQL = "SELECT RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe, RV_POCaricoMerceTesta.DataDocumento, RV_POCaricoMerceTesta.NumeroDocumento, "
        sSQL = sSQL & "RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceTesta.Anagrafica, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.Nome, RV_POCaricoMerceRighe.CodiceLotto "
        sSQL = sSQL & "FROM RV_POCaricoMerceTesta INNER JOIN "
        sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta "
        sSQL = sSQL & "WHERE RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe=" & IDRigaConf
        
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF Then
            GET_STRINGA_CONFERIMENTO = ""
        Else
            Testo = fnNotNull(rs!CodiceLotto) & " - "
            Testo = Testo & fnNotNull(rs!CodiceArticolo) & " - " & fnNotNull(rs!Articolo) & " - "
            Testo = Testo & "Numero " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " - "
            Testo = Testo & "del socio " & fnNotNull(rs!Anagrafica) & "  " & fnNotNull(rs!Nome)
            GET_STRINGA_CONFERIMENTO = Testo
        End If
        
        rs.CloseResultset
        Set rs = Nothing
    End If
    'PROCESSO DI ASSEGNAZIONE MERCE
    If ((IDRigaConf > 0) And (IDAssegnazioneMerce > 0) And (IDProcessoIVGamma = 0)) Then
        sSQL = "SELECT RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe, RV_POCaricoMerceTesta.DataDocumento, RV_POCaricoMerceTesta.NumeroDocumento, "
        sSQL = sSQL & "RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.CodiceArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceTesta.Anagrafica, "
        sSQL = sSQL & "RV_POCaricoMerceTesta.Nome, RV_POCaricoMerceRighe.CodiceLotto, "
        sSQL = sSQL & "RV_POAssegnazioneMerce.IDRV_POAssegnazioneMerce, RV_POAssegnazioneMerce.DataDocumento AS DataLavorazione "
        sSQL = sSQL & "FROM RV_POAssegnazioneMerce INNER JOIN "
        sSQL = sSQL & "RV_POCaricoMerceRighe ON "
        sSQL = sSQL & "RV_POAssegnazioneMerce.IDRV_POCaricoMerceRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe INNER JOIN "
        sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
        sSQL = sSQL & "WHERE RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe=" & IDRigaConf
        sSQL = sSQL & " AND RV_POAssegnazioneMerce.IDRV_POAssegnazioneMerce=" & IDAssegnazioneMerce

        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF Then
            GET_STRINGA_CONFERIMENTO = ""
        Else
            Testo = "Lavorazione del " & fnNotNull(rs!DataLavorazione) & " - "
            Testo = Testo & fnNotNull(rs!CodiceLotto) & " - "
            Testo = Testo & fnNotNull(rs!CodiceArticolo) & " - " & fnNotNull(rs!Articolo) & " - "
            Testo = Testo & "Numero " & fnNotNull(rs!NumeroDocumento) & " del " & fnNotNull(rs!DataDocumento) & " - "
            Testo = Testo & "del socio " & fnNotNull(rs!Anagrafica) & "  " & fnNotNull(rs!Nome)
            GET_STRINGA_CONFERIMENTO = Testo
        End If
        
        rs.CloseResultset
        Set rs = Nothing

    End If
    'PRCESSO DI LAVORAZIONE DI IV GAMMA
    If ((IDRigaConf = 0) And (IDAssegnazioneMerce > 0) And (IDProcessoIVGamma > 0)) Then
    
    End If
    

End Function
Private Function GET_CODICE(IDAnagrafica As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Codice FROM Fornitore "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & m_App.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE = ""
Else
    GET_CODICE = Trim(fnNotNull(rs!Codice))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub fnGrigliaCommissioni()

End Sub
Private Function GET_TOTALE_COMMISSIONI() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Sum(ImportoRiga) as TotaleRiga "
sSQL = sSQL & "FROM RV_POCommissioniPerDoc "
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_TOTALE_COMMISSIONI = 0
Else
    GET_TOTALE_COMMISSIONI = fnNotNullN(rs!TotaleRiga)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub ParametroGrezzo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoGrezzo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoGrezzo = fnNotNullN(rs!IDTipoGrezzo)
Else
    Link_TipoGrezzo = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroLavorato()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoLavorato FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoLavorato = fnNotNullN(rs!IDTipoLavorato)
Else
    Link_TipoLavorato = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroObbligatorio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Obbligatorio FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Par_OBBLIGATORIO = fnNotNullN(rs!Obbligatorio)
Else
    Par_OBBLIGATORIO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function fnGestioneArticoli() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoGestioneArticoliVendita "
sSQL = sSQL & "FROM RV_POSchemaCoop"
sSQL = sSQL & " WHERE IDFiliale=" & m_App.Branch

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    fnGestioneArticoli = 1
Else
    fnGestioneArticoli = fnNotNullN(rs!IDTipoGestioneArticoliVendita)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ARTICOLO_CONFERITO() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT IDArticolo FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ARTICOLO_CONFERITO = 0
Else
    GET_ARTICOLO_CONFERITO = fnNotNullN(rs!IDArticolo)
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Function fnGetCodiceLotto(TipoLotto As Integer, TipoStringaLotto As Integer, StringaLotto As String, Link_LottoArticolo As Long) As String

    
    
End Function
Private Function GET_STRINGALOTTO(IDRV_POLottoComp As Long, Lunghezza As Integer, Stringa As String, TipoStringa As Integer, SXDX As Integer, Optional TestoPersonalizzato As String) As String
Dim Parziale As String
Dim I As Integer
Parziale = ""

        If Len(Stringa) >= Lunghezza Then
            If SXDX = 2 Then 'Da destra verso sinistra
                GET_STRINGALOTTO = Right(Stringa, Lunghezza)
            Else
                'Da sinistra verso destra
                GET_STRINGALOTTO = Mid(Stringa, 1, Lunghezza)
            End If
        Else
            If SXDX <= 1 Then 'Da sinistra verso destra
                For I = Len(Stringa) To Lunghezza - 1
                    If TipoStringa = 0 Then
                        Parziale = "0" & Parziale
                    Else
                        Parziale = "" & Parziale
                    End If
                Next
                GET_STRINGALOTTO = Parziale & Stringa
            Else
                'Da destra verso sinistra
                GET_STRINGALOTTO = Right(Stringa, Lunghezza)
            End If
        End If
End Function


Private Function GET_NUMERO_LOTTO() As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT NumerazioneLotto FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    GET_NUMERO_LOTTO = 0
Else
    GET_NUMERO_LOTTO = fnNotNullN(rs!NumerazioneLotto) + 1
    rs!NumerazioneLotto = fnNotNullN(rs!NumerazioneLotto) + 1
    rs.Update
End If

rs.Close
Set rs = Nothing
End Function

Private Function GET_RIEPILOGO_QUANTITA_LAVORAZIONE(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



sSQL = "SELECT Sum(Qta_UM) as QtaTotale "
sSQL = sSQL & "FROM RV_POLavorazione "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = 0
Else
    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing



End Function
Private Function GET_RIEPILOGO_QUANTITA_VENDUTO(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_RIEPILOGO_QUANTITA_VENDUTO = 0

'DOCUMENTO DI TRASPORTO
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

'FATTURA ACCOMPAGNATORIA
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

'CORRISPETTIVI
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub GET_RIEPILOGO_CONFERIMENTO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT Colli, Qta_UM FROM RV_POCaricoMerceRighe "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & Link_RigaConferimento
    
    Set rs = Cn.OpenResultset(sSQL)
   
    If rs.EOF Then
        Me.txtQtaConferita.Value = 0
        Me.txtColliConferiti.Value = 0
    Else
        Me.txtQtaConferita.Value = fnNotNullN(rs!Qta_UM)
        Me.txtColliConferiti.Value = fnNotNullN(rs!Colli)

    End If
rs.CloseResultset
Set rs = Nothing

Me.txtColliVenduti_Conferimento.Value = GET_RIEPILOGO_COLLI_VENDUTO(Link_RigaConferimento)
Me.txtQtaVenduta_conferimento.Value = GET_RIEPILOGO_QUANTITA_VENDUTO(Link_RigaConferimento)
Me.txtQtaQuadrata_Conferimento.Value = GET_RIEPILOGO_QUANTITA_LAVORAZIONE(Link_RigaConferimento)
Me.txtDifferenzaConferimento.Value = Me.txtQtaConferita.Value - (Me.txtQtaQuadrata_Conferimento.Value + Me.txtQtaVenduta_conferimento.Value)
    

End Sub
Private Function GET_RIEPILOGO_COLLI_VENDUTO(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


GET_RIEPILOGO_COLLI_VENDUTO = 0

'DOCUMENTO DI TRASPORTO
sSQL = "SELECT Sum(Art_numero_colli) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_COLLI_VENDUTO = GET_RIEPILOGO_COLLI_VENDUTO
Else
    GET_RIEPILOGO_COLLI_VENDUTO = GET_RIEPILOGO_COLLI_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

'FATTURA ACCOMPAGNATORIA
sSQL = "SELECT Sum(Art_numero_colli) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_COLLI_VENDUTO = GET_RIEPILOGO_COLLI_VENDUTO
Else
    GET_RIEPILOGO_COLLI_VENDUTO = GET_RIEPILOGO_COLLI_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

'CORRISPETTIVI
sSQL = "SELECT Sum(Art_numero_colli) as QtaTotale "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_RIEPILOGO_COLLI_VENDUTO = GET_RIEPILOGO_COLLI_VENDUTO
Else
    GET_RIEPILOGO_COLLI_VENDUTO = GET_RIEPILOGO_COLLI_VENDUTO + fnNotNullN(rs!QtaTotale)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_PARAMETRO_GESTIONE_CHIUSURA_LOTTI() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT GestioneConferimento FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_GESTIONE_CHIUSURA_LOTTI = False
Else
    If IsNull(rs!GestioneConferimento) Then
        GET_PARAMETRO_GESTIONE_CHIUSURA_LOTTI = False
    Else
        GET_PARAMETRO_GESTIONE_CHIUSURA_LOTTI = fnNormBoolean(rs!GestioneConferimento)
    End If
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub fnAggiornaDescrizioneDocumento()
Dim sSQL As String

sSQL = "UPDATE " & sTabellaDettaglio & " SET "
sSQL = sSQL & "RV_PODescrizioneDocumento="
Select Case oDoc.IDTipoOggetto

    Case 2
        sSQL = sSQL & fnNormString("Rif. D.d.t. n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text)
    Case 114
        sSQL = sSQL & fnNormString("Rif. f.a. n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text)
    Case 8
        sSQL = sSQL & fnNormString("Rif. s.n.f. n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text)
    Case 15
        sSQL = sSQL & fnNormString("Rif. Ordine n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text)
End Select
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto

Cn.Execute sSQL


End Sub
Public Sub AzzerraArray()
Dim I As Integer

For I = 0 To 50
    ArrayConfMod(I) = 0
Next
End Sub



Private Sub txtTaraConfImballo_LostFocus()
    txtColli_LostFocus
End Sub

Private Sub txtTaraUnitaria_LostFocus()
    txtColli_LostFocus
    'CalcoloPesoNetto
End Sub

Private Sub sbCalcolaImportoLiquidazione(ImpImballo As Double)
End Sub
Private Function GET_PREZZO_IMBALLO(ImportoImballo As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If ImportoImballo = 0 Then
    sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
    sSQL = sSQL & "WHERE ("
    sSQL = sSQL & "(IDListino=" & Me.cboListino.CurrentID & ") "
    sSQL = sSQL & "AND (IDArticolo=" & Me.CDImballo.KeyFieldID & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        GET_PREZZO_IMBALLO = fnNotNullN(rs!PrezzoNettoIVA)
    Else
        GET_PREZZO_IMBALLO = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Else
    GET_PREZZO_IMBALLO = ImportoImballo
End If
End Function
Public Sub fnAggiornaCommissioniPerCliente()
On Error GoTo ERR_fnAggiornaCommissioniPerCliente
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rscomm As ADODB.Recordset
Dim Totale_merce_lavorato As Double
Dim Totale_merce_lavorato_lordo As Double
Dim Totale_Documento_merce_netto_iva As Double
Dim Totale_Documento_merce_lordo_iva As Double
Dim Totale_Importo As Double
Dim Link_Tipo_Valore_Comm As Long
Dim rsCli As ADODB.Recordset
Dim rsNew As ADODB.Recordset
Dim SpesaTrasporto As Double

Totale_merce_lavorato = 0 ' GET_TOTALE_MERCE_DOCUMENTO 'PRELEVA IL TOTALE MERCE NETTO DEGLI IMBALLI
Totale_merce_lavorato_lordo = 0 'GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO
Totale_Documento_merce_netto_iva = curTotImponibile.Value
Totale_Documento_merce_lordo_iva = curTotDocumento.Value

'RICALCOLO DELLE PERCENTUALI IN BASE ALL'IMPORTO FISSO
 sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
 sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
 sSQL = sSQL & " AND ((IDRV_POTipoPedana=0) OR (IDRV_POTipoPedana IS NULL))"
 Set rscomm = New ADODB.Recordset
 
 rscomm.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
 
' While Not rscomm.EOF
'     'If fnNotNullN(rsComm!ImportoRiga) > 0 Then
'     Link_Tipo_Valore_Comm = GET_TIPO_VALORE_COMMISSIONE(fnNotNullN(rscomm!IDRV_POTipoCommissione))
'     If Link_Tipo_Valore_Comm = 1 Then
'         If GET_TIPO_RICALCOLO_COMMISSIONE(fnNotNullN(rscomm!IDRV_POTipoCommissione)) = 1 Then
'             If Totale_merce_lavorato > 0 Then
'                 rscomm!ImportoRiga = (Totale_merce_lavorato / 100) * rscomm!Percentuale
'             Else
'                 rscomm!ImportoRiga = 0
'             End If
'         Else
'             If Totale_merce_lavorato > 0 Then
'                 rscomm!Percentuale = (fnNotNullN(rscomm!ImportoRiga) / Totale_merce_lavorato) * 100
'                 rscomm!ImportoRiga = (Totale_merce_lavorato / 100) * rscomm!Percentuale
'             Else
'                 rscomm!Percentuale = 0
'             End If
'         End If
'     Else
'         Select Case Link_Tipo_Valore_Comm
'             Case 1
'                 Totale_Importo = Totale_merce_lavorato
'             Case 2
'                 Totale_Importo = Totale_merce_lavorato_lordo
'             Case 3
'                 Totale_Importo = Totale_Documento_merce_netto_iva
'             Case 4
'                 Totale_Importo = Totale_Documento_merce_lordo_iva
'             Case Else
'                 Totale_Importo = Totale_merce_lavorato
'         End Select
'
'         If Totale_Importo = 0 Then
'             rscomm!Percentuale = 0
'             rscomm!ImportoRiga = 0
'         Else
'             rscomm!ImportoRiga = (Totale_Importo / 100) * fnNotNullN(rscomm!PercentualeDaCommissione)
'             If (Totale_merce_lavorato > 0) Then
'                 rscomm!Percentuale = (rscomm!ImportoRiga / Totale_merce_lavorato) * 100
'             Else
'                 rscomm!Percentuale = 0
'             End If
'
'         End If
'     End If
'     rscomm.Update
'     'End If
' rscomm.MoveNext
' Wend
'
' rscomm.Close
' Set rscomm = Nothing

'AGGIUNTA DELLE COMMISSIONI PER CLIENTE
sSQL = "SELECT * FROM RV_POCommissioniPerCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & Me.cdAnagrafica.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)
Set rscomm = New ADODB.Recordset
rscomm.Open "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & oDoc.IDOggetto, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
While Not rs.EOF

    If GET_ESISTENZA_COMMISSIONE_DOC(fnNotNullN(rs!IDRV_POTipoCommissione)) = False Then
        Link_Tipo_Valore_Comm = GET_TIPO_VALORE_COMMISSIONE(fnNotNullN(rs!IDRV_POTipoCommissione))
        
        rscomm.AddNew
            rscomm!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
            rscomm!IDOggetto = oDoc.IDOggetto
            rscomm!IDRV_POTipoCommissione = fnNotNullN(rs!IDRV_POTipoCommissione)
            rscomm!PercentualeDaCommissione = fnNotNullN(rs!Percentuale)
            rscomm!Quantita = 1
            rscomm!ImportoTotale = 0
            If Link_Tipo_Valore_Comm <= 1 Then
                rscomm!Percentuale = fnNotNullN(rs!Percentuale)
                rscomm!Importo = 0
                rscomm!ImportoRiga = (Totale_merce_lavorato / 100) * rscomm!Percentuale
            Else
                Select Case Link_Tipo_Valore_Comm
                    Case 1
                        Totale_Importo = Totale_merce_lavorato
                    Case 2
                        Totale_Importo = Totale_merce_lavorato_lordo
                    Case 3
                        Totale_Importo = Totale_Documento_merce_netto_iva
                    Case 4
                        Totale_Importo = Totale_Documento_merce_lordo_iva
                    Case Else
                        Totale_Importo = Totale_merce_lavorato
                End Select
                
                If Totale_Importo = 0 Then
                    rscomm!Percentuale = 0
                    rscomm!ImportoRiga = 0
                Else
                    rscomm!ImportoRiga = (Totale_Importo / 100) * fnNotNullN(rscomm!PercentualeDaCommissione)
                    If (Totale_merce_lavorato > 0) Then
                        rscomm!Percentuale = (rscomm!ImportoRiga / Totale_merce_lavorato) * 100
                    Else
                        rscomm!Percentuale = 0
                    End If
                End If
            End If
        rscomm.Update
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

'''''''''''''''''''''''''''''''COMMISSIONI PER TIPO PEDANA/IMBALLO''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT Link_Art_articolo AS IDArticoloImballo, SUM(Art_quantita_totale) AS QuantitaPedana "
sSQL = sSQL & "FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE IDOggetto =" & oDoc.IDOggetto
sSQL = sSQL & " AND RV_POTipoRiga = 2 "
sSQL = sSQL & "GROUP BY Link_Art_articolo"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    
    sSQL = "SELECT * FROM RV_POIEConfigurazioneClienteTrasporto "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDAnagrafica=" & Me.cdAnagrafica.KeyFieldID
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & Me.cboAltroSito.CurrentID
    sSQL = sSQL & " AND IDArticolo=" & fnNotNullN(rs!IDArticoloImballo)
    sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
    sSQL = sSQL & " AND ((CommissionePerPedana=0) OR (CommissionePerPedana IS NULL))"
    
    Set rsCli = New ADODB.Recordset
    
    rsCli.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    
    While Not rsCli.EOF
        If GET_ESISTENZA_COMMISSIONE_DOC_TIPO_IMB(fnNotNullN(rsCli!IDRV_POTipoCommissione), fnNotNullN(rs!IDArticoloImballo)) = False Then
            rscomm.AddNew
                rscomm!IDRV_POCommissioniPerDoc = fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc")
                rscomm!IDOggetto = oDoc.IDOggetto
                rscomm!IDRV_POTipoCommissione = fnNotNullN(rsCli!IDRV_POTipoCommissione)
                'SpesaTrasporto = (fnNotNullN(rsCli!PrezzoTrasporto) * (fnNotNullN(rs!QuantitaPedana) / fnNotNullN(rsCli!Quantita)))
                rscomm!Percentuale = 0 ' (SpesaTrasporto / Totale_merce_lavorato) * 100
                rscomm!Importo = 0
                If fnNotNullN(rsCli!Quantita) = 0 Then
                    rscomm!ImportoRiga = fnNotNullN(rsCli!PrezzoTrasporto)
                Else
                    rscomm!ImportoRiga = fnNotNullN(rsCli!PrezzoTrasporto) / fnNotNullN(rsCli!Quantita)
                End If
                rscomm!Quantita = 1
                rscomm!ImportoTotale = 0
                rscomm!IDArticoloImballo = fnNotNullN(rs!IDArticoloImballo)
            rscomm.Update
        End If
    rsCli.MoveNext
    Wend
    
    rsCli.Close
    Set rsCli = Nothing
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing


rscomm.Close
Set rscomm = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_fnAggiornaCommissioniPerCliente:
    MsgBox Err.Description, vbCritical, "fnAggiornaCommissioniPerCliente"
    
End Sub
Public Function GET_ESISTENZA_COMMISSIONE_DOC(IDTipoCommissione As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

    sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
    sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione
    sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_ESISTENZA_COMMISSIONE_DOC = False
    Else
        GET_ESISTENZA_COMMISSIONE_DOC = True
    End If
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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT = 0
Else
    GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDefault)
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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = 0
Else
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = fnNotNullN(rs!IDListinoImballiDefault)
    
End If
rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_ESISTENZA_NUMERO_DOCUMENTO(IDTipoOggetto As Long, IDOggetto As Long) As Boolean
Dim sSQL  As String
Dim rs As DmtOleDbLib.adoResultset

If IDOggetto = 0 Then
    sSQL = "SELECT IDOggetto FROM Oggetto "
    sSQL = sSQL & "WHERE IDSezionale=" & Me.cboSezionale.CurrentID
    sSQL = sSQL & " AND IDAzienda=" & oDoc.IDAzienda
    sSQL = sSQL & " AND Numero=" & fnNormString(Me.lngNumero.Value)
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate("01/01/" & Year(Me.dtData.Text))
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate("31/12/" & Year(Me.dtData.Text))
    
    
Else
    sSQL = "SELECT IDOggetto FROM Oggetto "
    sSQL = sSQL & "WHERE IDSezionale=" & Me.cboSezionale.CurrentID
    sSQL = sSQL & " AND IDAzienda=" & oDoc.IDAzienda
    sSQL = sSQL & " AND Numero=" & fnNormString(Me.lngNumero.Value)
    sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
    sSQL = sSQL & " AND DataEmissione>=" & fnNormDate("01/01/" & Year(Me.dtData.Text))
    sSQL = sSQL & " AND DataEmissione<=" & fnNormDate("31/12/" & Year(Me.dtData.Text))
    sSQL = sSQL & " AND IDOggetto<>" & IDOggetto
    
End If


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_NUMERO_DOCUMENTO = False
Else
    GET_ESISTENZA_NUMERO_DOCUMENTO = True
End If
rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CalcoloPesoNetto()
On Error GoTo ERR_CalcoloPesoNetto


Me.txtTara.Value = Me.txtColli.Value * (Me.txtTaraUnitaria.Value + (Me.txtNumeroConfImballo.Value * Me.txtTaraConfImballo.Value))


If TIPO_PESO_ARTICOLO <= 1 Then
    Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
Else
    Me.txtPesoLordo.Value = Me.txtTara.Value + Me.txtPesoNetto.Value
End If

Select Case Link_Arrotondamento
    Case 1 'Nessun arrotondamento
        Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
    Case 2 'Matematico
        Me.txtPesoNetto.Value = fnRoundChange(Me.txtPesoNetto.Value, 1, 3)
        Me.txtTara.Value = Me.txtPesoLordo.Value - Me.txtPesoNetto.Value
    Case 3 'Difetto
        Me.txtPesoNetto.Value = fnRoundDown(Me.txtPesoNetto.Value)
        Me.txtTara.Value = Me.txtPesoLordo.Value - Me.txtPesoNetto.Value
    Case 4 'Eccesso
        Me.txtPesoNetto.Value = fnRoundUp(Me.txtPesoNetto.Value)
        Me.txtTara.Value = Me.txtPesoLordo.Value - Me.txtPesoNetto.Value
End Select


Exit Sub
ERR_CalcoloPesoNetto:
    MsgBox Err.Description, vbCritical, "CalcoloPesoNetto"

End Sub
Private Function GET_NUMERO_COPIE() As Long

End Function
Private Function GET_ORIENTAMENTO() As Long
End Function

Private Function GET_CERTIFICAZIONE_SOCIO(IDSocio As Long, NomeCampo As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo & " "
sSQL = sSQL & "FROM RV_PO01_CertificazioneSocio INNER JOIN "
sSQL = sSQL & "RV_PO01_Certificazione ON RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDSocio
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CERTIFICAZIONE_SOCIO = ""
Else
    GET_CERTIFICAZIONE_SOCIO = fnNotNull(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_CERTIFICAZIONE_LOTTO_CAMPAGNA(IDRigaConferimento As Long, NomeCampo As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsConf As DmtOleDbLib.adoResultset
Dim IDLottoDiCampagna As Long

'RECUPERO IDENTIFICATIVO DEL LOTTO DI CAMPAGNA
sSQL = "SELECT IDLottoDiCampagna FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rsConf = Cn.OpenResultset(sSQL)

If rsConf.EOF Then
    IDLottoDiCampagna = 0
Else
    IDLottoDiCampagna = fnNotNullN(rsConf!IDLottoDiCampagna)
End If


rsConf.CloseResultset
Set rsConf = Nothing
'''''''''''''''''''FINE''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_PO01_LottoCampagna INNER JOIN "
sSQL = sSQL & "RV_PO01_CertificazioneSocio ON "
sSQL = sSQL & "RV_PO01_LottoCampagna.IDRV_PO01_CertificazioneSocio = RV_PO01_CertificazioneSocio.IDRV_PO01_CertificazioneSocio INNER JOIN "
sSQL = sSQL & "RV_PO01_Certificazione ON RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoDiCampagna

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CERTIFICAZIONE_LOTTO_CAMPAGNA = ""
Else
    GET_CERTIFICAZIONE_LOTTO_CAMPAGNA = fnNotNull(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TIPO_SCONTO(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTipoScontoPerCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

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
Private Sub BarMenu_BandClose(ByVal Band As ActiveBar3LibraryCtl.Band)
     'Se la banda è una Toolbar allora viene registrata la chiusura.
    If Band.Type = ddBTNormal And Band.Name <> BAND_CLOSE_PREVIEW Then
        
        'Salva nel registry l'impostazione sulla visibilità della toolbar
        AppOptions.ToolbarVisibility(Band.Name) = False
        
    End If
End Sub

Private Sub BarMenu_BandMove(ByVal Band As ActiveBar3LibraryCtl.Band)
    Form_Resize
End Sub

Private Sub BarMenu_BandOpen(ByVal Band As ActiveBar3LibraryCtl.Band, ByVal Cancel As ActiveBar3LibraryCtl.ReturnBool)
     'Se la banda è una Toolbar allora viene registrata l'apertura.
    If Band.Type = ddBTNormal And Band.Name <> BAND_CLOSE_PREVIEW Then
        AppOptions.ToolbarVisibility(Band.Name) = True
    End If
End Sub

Private Sub BarMenu_MenuItemEnter(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar Tool.Description
End Sub

Private Sub BarMenu_MenuItemExit(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar ""
End Sub

Private Sub BarMenu_MouseEnter(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar Tool.Description
End Sub

Private Sub BarMenu_MouseExit(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    WriteStatusBar ""
End Sub
Private Sub BarMenu_QueryUnload(Cancel As Integer)
    Cancel = True
End Sub

Private Sub BarMenu_Resize(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Form_Resize
End Sub

Private Sub BarMenu_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    Dim iKeyCode As Integer
    Dim iShift As Integer
    Dim bContinue As Boolean
    
    On Error GoTo BarMenu_ClickError
        
    'Forza il lostfocus ed attende l'esecuzione di eventuali eventi associati
    AutoLostFocus
        
    bContinue = True
    iShift = GetShift(Tool)
    iKeyCode = GetKeyCode(Tool)
    
    If iKeyCode <> 0 Or iShift <> 0 Then
        bContinue = Not ShortCut(iKeyCode, iShift)
        If bContinue Then
            SendKeys GetSendKeys(Tool) & "(" & GetKey(Tool) & ")"      '"^(R)"
        End If
    Else
        ExecuteMenuCommand Tool.Name
    End If
    
    Exit Sub
    
BarMenu_ClickError:
    If Err.Number = ERR_NDELFILTER Then
        'In seguito a particolari sequenze di eventi può risultare abilitato il cancella filtro sul
        'filtro di default. Se si esegue la cancellazione viene sollevata una eccezione.
        sbMsgError "Non è possibile eliminare il filtro di default.", m_App.FunctionName
    Else
        sbMsgError Err.Description, m_App.FunctionName
    End If
    
    Resume Next
End Sub
'**+
'Nome: SendDocument
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Esegue l'esportazione del documento con controllo di errore
'**/
Private Sub SendDocument(ByVal Appl As Long)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    m_Document.SendMail m_Report, Appl
    Screen.MousePointer = OLDCursor
    
    Exit Sub
errHandler:
    Screen.MousePointer = OLDCursor
    
    If Err.Number = 20507 Then
        'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
        sbMsgInfo "File di report non trovato", m_App.FunctionName
    Else
        sbMsgInfo Err.Description, m_App.FunctionName
    End If
End Sub
'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : InitSemaphore
'
'Parametri                  :
'
'Funzionalità               : Attiva/disattiva le attività del Riquadro attività
'
'**/
Private Sub EnableDOMActivitiesItems()
    oFiltersActivity.EnableItems (BrwMain.GuiMode = dgNormal And BrwMain.Visible)
    oTableViewsActivity.EnableItems (BrwMain.GuiMode = dgNormal And BrwMain.Visible)
    
    ActivityBox.Redraw = True
End Sub
'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : ActivityBox_CloseButtonPressed
'
'Parametri                  :
'
'Funzionalità               : Gestione della chiusura del Riquadro attività
'
'**/
Private Sub ActivityBox_CloseButtonPressed()
    OnFolders
End Sub

'**+
'Autore                     : Diamante s.p.a
'Data creazione             :
'Nome                       : ActivityBox_ItemSelected
'
'Parametri                  :
'
'Funzionalità               : Gestione della selezione delle voci del Riquadro attività
'
'**/
Private Sub ActivityBox_ItemSelected(ByVal Item As DmtActBoxTlb.Item, NeedRedraw As Boolean)
    Dim oFilter As Filter
    Dim oTableView As TableView
    
    If BrwMain.Visible And BrwMain.GuiMode = dgNormal Then
        Select Case ActivityBox.CurrentActivity.Caption
            Case "Filtri"
                For Each oFilter In m_DocType.Filters
                    If oFilter.ID = Val(Item.Tag) Then
                        Set m_ActiveFilter = m_DocType.Filters(oFilter.Name)
                        Exit For
                    End If
                Next
                'Flag usato per specificare che deve essere eseguito un filtro permanente.
                m_FilterSelected = True
                
                '---Modalità Filtro---
                'Ripulisce i campi di immissione delle condizioni di ricerca.
                BrwMain.Conditions.ClearValues
                
                'Se attivo, viene disabilitato il pulsante Salva Filtro del DocTypeExplorer.
                oFiltersActivity.AbortNewFilter
                ActivityBox.Redraw = True
                
                'Viene eseguita la ricerca basata sul nuovo filtro.
                ExecuteSearch

            Case "Viste tabellari"
                For Each oTableView In m_DocType.TableViews
                    If oTableView.ID = Val(Item.Tag) Then
                        Set m_ActiveTableView = m_DocType.TableViews(oTableView.Name)
                        Exit For
                    End If
                Next
                BrwMain.LoadColumns m_ActiveTableView
                SetVisibilityIDFields
        End Select
    End If
    If ActivityBox.CurrentActivity.Caption = "Esportazioni" Then
        If Item.Hyperlink Then
            Select Case Item.Name
                Case "E" & ExportConstants.PDF
                    ExecuteMenuCommand "ExportPDF"
                Case "E" & ExportConstants.Word
                    ExecuteMenuCommand "ExportWord"
                Case "E" & ExportConstants.Excel
                    ExecuteMenuCommand "ExportExcel"
                Case "E" & ExportConstants.HTML
                    ExecuteMenuCommand "ExportHtml"
                Case "S" & ExportConstants.PDF
                    ExecuteMenuCommand "MailPDF"
                Case "S" & ExportConstants.Word
                    ExecuteMenuCommand "MailWord"
                Case "S" & ExportConstants.Excel
                    ExecuteMenuCommand "MailExcel"
                Case "S" & ExportConstants.HTML
                    ExecuteMenuCommand "MailHtml"
            End Select
        End If
    End If
End Sub
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width, ActivityBox.Height
        picSplitter.AutoRedraw = True
    End With
    picSplitter.Visible = True
    m_SplitterMoving = True
    picSplitter.ZOrder
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single

    If m_SplitterMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < SPLITLIMIT Then
            picSplitter.Left = SPLITLIMIT
        ElseIf sglPos > BarMenu.ClientAreaWidth - SPLITLIMIT Then
            picSplitter.Left = BarMenu.ClientAreaWidth - SPLITLIMIT
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActivityBox.Width = picSplitter.Left - ActivityBox.Left
    FormRecalcLayout
    picSplitter.Visible = False
    m_SplitterMoving = False
End Sub





Private Sub GET_CONTROLLO_LICENZA()
Dim sSQL As String
Dim Codice_Diamante As String
Dim Codice_Prodotto_Calcolato  As String
Dim Codice_Attivazione As String

Dim Partita_Iva_Licenza As String

Codice_Diamante = GET_CODICE_DIAMANTE
Partita_Iva_Licenza = GET_PARTITA_IVA
Codice_Attivazione = GET_CODICE_SBLOCCO_ATTIVAZIONE


Codice_Prodotto_Calcolato = GET_CODICE_SBLOCCO(Codice_Diamante, Partita_Iva_Licenza, "05")

If Codice_Attivazione = Codice_Prodotto_Calcolato Then
    DEMO = False
Else
    DEMO = True
End If

    
End Sub
Private Function GET_CODICE_DIAMANTE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione FROM ComponenteSwAbilitata "
sSQL = sSQL & "WHERE NomeCompSW=" & fnNormString("*IDSW___")

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_DIAMANTE = ""
Else
    GET_CODICE_DIAMANTE = Trim(fnNotNull(rs!Descrizione))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PARTITA_IVA() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PartitaIva FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=5"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARTITA_IVA = ""
Else
    GET_PARTITA_IVA = fnCryptString(Trim(fnNotNull(rs!PartitaIVA)))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CODICE_SBLOCCO_ATTIVAZIONE() As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CodiceSblocco FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=5"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_SBLOCCO_ATTIVAZIONE = ""
Else
    GET_CODICE_SBLOCCO_ATTIVAZIONE = fnCryptString(Trim(fnNotNull(rs!CodiceSblocco)))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_NUMERO_INSERIMENTI() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDOggetto) As NumeroInserimenti "
sSQL = sSQL & "FROM Oggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND IDAzienda=" & m_App.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_INSERIMENTI = 0
Else
    GET_NUMERO_INSERIMENTI = fnNotNullN(rs!NumeroInserimenti)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub txtTargaAutomezzo_LostFocus()
    sbImpostaDatiDocumento
End Sub
Private Sub AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI(IDTipoOggetto As Long, IDOggetto As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim QuantitaMovimento As Double
Dim Link_Unita_di_Misura_Conferimeto As Long




Screen.MousePointer = 11
Me.Caption = "AGGIORNAMENTO MOVIMENTI....................."
'Me.lblInfoTesta.Font.Bold = True
'Me.lblInfoTesta.ForeColor = vbBlue
DoEvents


Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection

DELETE_MOV_CONFERIMENTO

sSQL = "SELECT " & sTabellaDettaglio & ".IDValoriOggettoDettaglio, "
sSQL = sSQL & sTabellaDettaglio & ".Art_pre_uni_net_sco_net_IVA, "
sSQL = sSQL & sTabellaDettaglio & ".Link_art_articolo, "
sSQL = sSQL & sTabellaDettaglio & ".Art_quantita_pezzi, "
sSQL = sSQL & sTabellaDettaglio & ".Art_numero_colli, "
sSQL = sSQL & sTabellaDettaglio & ".Art_tara, "
sSQL = sSQL & sTabellaDettaglio & ".Art_peso, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODataConferimento, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDConferimentoRighe, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCodiceLotto, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDSocio, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POLottoCampagna, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoDaLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POQuantitaLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDAssegnazioneMerce, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDProcessoIVGamma, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POPrezzoMedioInLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoUnitarioImballo, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceTesta.NumeroDocumento, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDUnitaDiMisura, RV_POCaricoMerceRighe.CodiceLotto, "
sSQL = sSQL & "RV_POCaricoMerceTesta.IDMagazzinoConferimento, RV_POCaricoMerceRighe.IDUnitaDiMisuraDiamante "
sSQL = sSQL & "FROM RV_POCaricoMerceTesta INNER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta = RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta INNER JOIN "
sSQL = sSQL & sTabellaDettaglio & " ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe = " & sTabellaDettaglio & ".RV_POIDConferimentoRighe "
sSQL = sSQL & "WHERE " & sTabellaDettaglio & ".IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND " & sTabellaDettaglio & ".RV_POTipoRiga=1"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection
    
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
    
    If ((fnNotNullN(rs!RV_POIDAssegnazioneMerce) = 0) And (fnNotNullN(rs!RV_POIDProcessoIVGamma) = 0)) Then
        Aggiorna_Movimento_Documento fnNotNullN(rs!Link_Art_articolo), fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNullN(rs!NumeroDocumento), _
        fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
        fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), QuantitaMovimento, _
        fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)), _
        fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POPrezzoMedioInLiq), fnNotNullN(rs!RV_POIDImballo), fnNotNullN(rs!RV_POImportoUnitarioImballo)

        DoEvents
                         
        GeneraMovimentoDiScarico fnNotNullN(rs!IDMagazzinoConferimento), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!IDArticolo), fnNotNull(rs!Articolo), fnNotNullN(rs!IDUnitaDiMisuraDiamante), QuantitaMovimento
    
    Else
        
        Aggiorna_Movimento_Documento fnNotNullN(rs!Link_Art_articolo), fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNull(rs!NumeroDocumento), _
        fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
        fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), QuantitaMovimento, _
        fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), (fnNotNullN(rs!Art_peso) - fnNotNullN(rs!Art_tara)), _
        fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POPrezzoMedioInLiq), fnNotNullN(rs!RV_POIDImballo), fnNotNullN(rs!RV_POImportoUnitarioImballo)
    
    End If
    
    DoEvents

rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set Mov = Nothing

Me.lblInfoTesta.Font.Bold = False
Me.lblInfoTesta.ForeColor = vbBlack



Screen.MousePointer = 0
End Sub
Private Function Aggiorna_Movimento_Documento(IDArticolo As Long, ImportoUnitarioArticolo As Double, IDRiga As Long, IDRigaConferimento, IDAssegnazione As Long, IDProcessoIVGamma As Long, IDSocio As Long, _
DataConferimento As String, NumeroConferimento As Long, _
CodiceLottoEntrata As String, CodiceLottoCampagna As String, CodiceLottoVendita As String, _
QuantitaLiquidazione As Double, ImportoInclusoImballo As Double, ImportoLiquidazione As Double, _
QuantitaMovimentata As Double, Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double, _
PrezzoMedioLiq As Double, IDArticoloImballo As Long, ImportoUnitarioImballo) As Long

Dim Prezzo As Double
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim Moltiplicatore As Double


    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(IDArticolo)
    
    
    
    sSQL = "SELECT * FROM Movimento "
    sSQL = sSQL & "WHERE IDTipoOggetto=" & oDoc.IDTipoOggetto
    sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDRiga
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    If Not rs.EOF Then
        rs("RV_POTipoRiga").Value = 1
        rs("RV_POIDCaricoMerceRighe").Value = IDRigaConferimento
        rs("RV_POIDAssegnazioneMerce").Value = IDAssegnazione
        rs("RV_POIDProcessoIVGamma") = IDProcessoIVGamma
        rs("RV_POIDAnagraficaSocio") = IDSocio
        rs("RV_PODataConferimento") = DataConferimento
        rs("RV_PONumeroConferimento") = NumeroConferimento
        rs("RV_POCodiceLotto") = CodiceLottoEntrata
        rs("RV_POCodiceLottoCampagna") = CodiceLottoCampagna
        rs("RV_POCodiceLottoVendita") = CodiceLottoVendita
        rs("RV_POQuantitaLiquidazione") = QuantitaLiquidazione
        rs("RV_POImportoInclusoImballo") = ImportoInclusoImballo
        
        Prezzo = ImportoUnitarioArticolo
        
        Prezzo = Prezzo / Moltiplicatore
        
        If ImportoInclusoImballo > 0 Then
            Prezzo = Prezzo + ImportoInclusoImballo
        Else
            Prezzo = Prezzo - Abs(ImportoInclusoImballo)
        End If
        
        'CALCOLO DEL PREZZO DI LIQUIDAZIONE'''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, oDoc.IDOggetto, oDoc.IDTipoOggetto, Moltiplicatore)
        
        rs("RV_POImportoLiquidazione") = Prezzo
        rs("RV_POQuantitaMovimentata") = QuantitaMovimentata
        rs("RV_PONumeroColli") = Colli
        rs("RV_POPesoLordo") = PesoLordo
        rs("RV_POPesoNetto") = PesoLordo - Tara
        rs("RV_POTara") = Tara
        rs("RV_POQuantitaPezzi") = Pezzi
        
        rs("RV_POPrezzoMedioInLiq").Value = PrezzoMedioLiq
        rs("RV_POIDImballo").Value = IDArticoloImballo
        rs("RV_POImportoUnitarioImballo").Value = ImportoUnitarioImballo
        
        rs("Oggetto").Value = GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto)
        rs.Update
        
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function
Private Function GET_COMMISSIONI_DOCUMENTO(PrezzoLiquidazione As Double, IDOggetto As Long, IDTipoOggetto As Long, Moltiplicatore As Double) As Double
Dim sSQL As String
Dim rscomm As DmtOleDbLib.adoResultset

GET_COMMISSIONI_DOCUMENTO = 0

sSQL = "SELECT * FROM RV_POCommissioniPerDoc WHERE IDOggetto=" & IDOggetto
Set rscomm = Cn.OpenResultset(sSQL)

While Not rscomm.EOF
    GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + ((PrezzoLiquidazione / 100) * fnNotNullN(rscomm!Percentuale))
    GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + (fnNotNullN(rscomm!Importo) / Moltiplicatore)
rscomm.MoveNext
Wend

rscomm.CloseResultset
Set rscomm = Nothing

GET_COMMISSIONI_DOCUMENTO = PrezzoLiquidazione - GET_COMMISSIONI_DOCUMENTO
End Function
Private Function GET_MOLTIPLICATORE_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POMoltiplicatore FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

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
Private Function GeneraMovimentoDiScarico(IDMagazzino As Long, IDRiga As Long, IDRigaConferimento As Long, IDArticoloConferito As Long, DescrizioneArticolo As String, IDUnitaDiMisura As Long, Quantita As Double) As Boolean

Mov.DataMovimento = Date
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = oDoc.IDEsercizio
Mov.IDTipoOggetto = oDoc.IDTipoOggetto
Mov.IDOggetto = oDoc.IDOggetto
Mov.IDFunzione = LINK_FUNZIONE_SCARICO_CONFERIMENTO
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoUscita = IDMagazzino
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Me.cdAnagrafica.KeyFieldID
Mov.Field "IDTipoAnagrafica", 2
Mov.Field "IDArticolo", IDArticoloConferito
Mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", DescrizioneArticolo
Mov.Field "QuantitaTotale", Quantita
Mov.Field "DataDocumento", Me.dtData.Text
Mov.Field "Oggetto", oDoc.Descrizione & " del " & Me.dtData.Text & " numero " & Me.lngNumero.Value
Mov.Field "IDTipoMovimento", 1
Mov.Field "TipoRiga", trcNessuno

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDRiga

GeneraMovimentoDiScarico = Mov.Insert

End Function
Private Function GET_DESCRIZIONE_TIPOOGGETTO(IDTipoOggetto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
                
    sSQL = "Select Oggetto "
    sSQL = sSQL & "FROM TipoOggetto "
    sSQL = sSQL & "WHERE IDTipoOggetto = " & IDTipoOggetto
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        GET_DESCRIZIONE_TIPOOGGETTO = fnNotNull(rs!Oggetto)
    Else
        GET_DESCRIZIONE_TIPOOGGETTO = ""
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub ParametroNuovoCalcolo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivazioneNuovoMetodoCalcolo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    ATTIVAZIONE_NUOVO_CALCOLO = fnNotNullN(rs!AttivazioneNuovoMetodoCalcolo)
Else
    ATTIVAZIONE_NUOVO_CALCOLO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_FUNZIONE_MAGAZZINO(IDTipoDocumentoCoop As Long, IDTipoProcesso As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POProcessiDocumentoCoop.IDFunzione "
sSQL = sSQL & "FROM RV_POProcessiDocumentoCoop INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE RV_POSchemaCoop.IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDDocumentoCoop=" & IDTipoDocumentoCoop
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDTipoProcessoCoop=" & IDTipoProcesso

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Select Case IDTipoProcesso
        Case 1 'Carico
            GET_FUNZIONE_MAGAZZINO = oDoc.IDFunzione
        Case 2 'Scarico
            GET_FUNZIONE_MAGAZZINO = oDoc.IDFunzione
    End Select
Else
    If fnNotNullN(rs!IDFunzione) = 0 Then
        Select Case IDTipoProcesso
            Case 1 'Carico
                GET_FUNZIONE_MAGAZZINO = oDoc.IDFunzione
            Case 2 'Scarico
                GET_FUNZIONE_MAGAZZINO = oDoc.IDFunzione
        End Select
        
    Else
        GET_FUNZIONE_MAGAZZINO = fnNotNullN(rs!IDFunzione)
    End If
End If


rs.CloseResultset
Set rs = Nothing

End Function
Private Sub ParametroTipoCaloPeso()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoCaloPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
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
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoAumentoPeso FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
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
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoScarto FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
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
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Private Sub DELETE_MOV_CONFERIMENTO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
'Dim Mov As DmtMovim.cMovimentazione
'Set Mov = New DmtMovim.cMovimentazione

'Set Mov.Connection = TheApp.Database.Connection



sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND IDFunzione=" & LINK_FUNZIONE_SCARICO_CONFERIMENTO

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    Mov.Delete rs!IDMovimento
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub


Private Sub CONTROLLA_BLOCCHI_INSERIMENTI()
Dim Abilitato As Boolean
Dim I As Integer

If oDoc.IsLocked = True Then
    Abilitato = False
Else
    Abilitato = True
End If

For I = 0 To Controls.Count - 1
    
    If TypeOf Controls(I) Is TextBox Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is DmtCodDesc Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is DMTCombo Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is dmtCurrency Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is dmtDate Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is dmtNumber Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is ComboBox Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is CheckBox Then
        Controls(I).Enabled = Abilitato
    End If
    If TypeOf Controls(I) Is CommandButton Then
        Controls(I).Enabled = Abilitato
    End If

Next

'''''''''CONTROLLI STANDARD SEMPRE ABILITATI O NON ABILITATI
'PIEDE DEL DOCUMENTO
Me.curTotArrotondamenti.Enabled = False
Me.curTotImposta.Enabled = False
Me.curTotImponibile.Enabled = False
Me.curTotDocumento.Enabled = False
Me.curNettoAPagare.Enabled = False
Me.curNettoAPagare_naz.Enabled = False



'CORPO DOCUMENTO
Me.txtImponibileUnitario.Enabled = False
Me.txtImponibileArticolo.Enabled = False
Me.txtImponibileImballo.Enabled = False
Me.txtTotaleImponibile.Enabled = False
Me.txtTotaleRiga.Enabled = False
Me.cmdSalva.Enabled = True






End Sub

Private Function GET_CONTROLLO_FIDO_CLIENTE() As Boolean
'Consideriamo di aver istanziato un oggetto oServices di tipo DmtCliFu.Services
Dim oServices As DMTCliFu.Services

Set oServices = New DMTCliFu.Services

'Passo una connessione al database di DMT (di tipo DmtOleDbLib.adoConnection)

Set oServices.Connection = Cn
'Identificativo dell'Azienda

oServices.IDAzienda = TheApp.IDFirm
'Identificativo dell'anagrafica del cliente

oServices.IDAnagrafica = Me.cdAnagrafica.KeyFieldID
'Totale del documento corrente

oServices.TotDocumento = Me.curNettoAPagare.Value
'IDPagamaneto del documento corrente

oServices.IDPagamento = Me.cboPagamento.CurrentID

'Identificativo univoco dell'oggetto del documento da cui leggere il totale netto a pagare nella valuta dell'azienda
oServices.IDDocumento = oDoc.IDOggetto

'Tabella di testata del tipo di documento che si sta prendendo in considerazione per la determinazione del netto a pagare
oServices.DocTableName = sTabellaTestata

GET_CONTROLLO_FIDO_CLIENTE = oServices.CheckFido

Set oServices = Nothing
End Function

Private Function GET_INFO_CLIENTE_FIDO_BLOCCO(IDAnagrafica As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


''''''''''FIDO PER CLIENTE
sSQL = "SELECT BloccoEmissioneDoc, IDTipoBloccoFido, PasswordSbloccoFido, DataFineSbloccoFido "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_BLOCCO_CLIENTE = 0
    LINK_TIPO_FIDO_CLIENTE = 0
    PASSWORD_FIDO_CLIENTE = ""
    DATA_SBLOCCO_FIDO_CLIENTE = ""
    
Else
    LINK_BLOCCO_CLIENTE = fnNotNullN(rs!BloccoEmissioneDoc)
    LINK_TIPO_FIDO_CLIENTE = fnNotNullN(rs!IDTipoBloccoFido)
    PASSWORD_FIDO_CLIENTE = fnCryptString(fnNotNull(rs!PasswordSbloccoFido))
    DATA_SBLOCCO_FIDO_CLIENTE = fnNotNull(rs!DataFineSbloccoFido)
End If

rs.CloseResultset
Set rs = Nothing

'''''''''''FIDO PER AZIENDA
sSQL = "SELECT IDTipoBloccoFido, PasswordSbloccoFido "
sSQL = sSQL & "FROM ConfigurazioneGenerale "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_FIDO_AZIENDA = 0
    PASSWORD_FIDO_AZIENDA = ""
Else
    LINK_TIPO_FIDO_AZIENDA = fnNotNullN(rs!IDTipoBloccoFido)
    PASSWORD_FIDO_AZIENDA = fnCryptString(fnNotNull(rs!PasswordSbloccoFido))
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_LINK_ANAGRAFICA_AZIENDA(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM Azienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANAGRAFICA_AZIENDA = 0
Else
    GET_LINK_ANAGRAFICA_AZIENDA = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_QUANTITA_IMBALLO_PER_TIPO_PEDANA(IDTipoPedana As Long, IDImballo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Quantita FROM RV_POTipoPedanaImballo "
sSQL = sSQL & "WHERE IDRV_POTipoPedana=" & IDTipoPedana
'sSQL = sSQL & "WHERE IDArticolo=" & IDTipoPedana
sSQL = sSQL & " AND IDArticolo=" & IDImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_QUANTITA_IMBALLO_PER_TIPO_PEDANA = 1
Else
    GET_QUANTITA_IMBALLO_PER_TIPO_PEDANA = fnNotNullN(rs!Quantita)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PREZZO_UNITARIO_ARTICOLO(IDArticolo As Long, IDListinoCliente As Long, IDListinoAzienda As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE ("
sSQL = sSQL & "(IDListino=" & IDListinoCliente & ") "
sSQL = sSQL & "AND (IDArticolo=" & IDArticolo & "))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    rs.CloseResultset
    Set rs = Nothing
    
    sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
    sSQL = sSQL & "WHERE ("
    sSQL = sSQL & "(IDListino=" & IDListinoAzienda & ") "
    sSQL = sSQL & "AND (IDArticolo=" & IDArticolo & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
        
    If rs.EOF Then
        GET_PREZZO_UNITARIO_ARTICOLO = 0
    Else
        GET_PREZZO_UNITARIO_ARTICOLO = fnNotNullN(rs!PrezzoNettoIVA)
    End If
    
    
    
        
Else
    If fnNotNullN(rs!PrezzoNettoIVA) = 0 Then
        rs.CloseResultset
        Set rs = Nothing
        
        sSQL = "SELECT PrezzoNettoIVA FROM ListinoPerArticolo "
        sSQL = sSQL & "WHERE ("
        sSQL = sSQL & "(IDListino=" & IDListinoAzienda & ") "
        sSQL = sSQL & "AND (IDArticolo=" & IDArticolo & "))"
        
        Set rs = Cn.OpenResultset(sSQL)
            
        If rs.EOF Then
            GET_PREZZO_UNITARIO_ARTICOLO = 0
        Else
            GET_PREZZO_UNITARIO_ARTICOLO = fnNotNullN(rs!PrezzoNettoIVA)
        End If
        
    Else
        GET_PREZZO_UNITARIO_ARTICOLO = fnNotNullN(rs!PrezzoNettoIVA)
    End If
    
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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE = 0
Else
    GET_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_BLOCCO_ORDINE_TOUR() As String
Dim LINK_TIPO_OGGETTO_TOUR As Long
Dim LINK_FUNZIONE_TOUR As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_BLOCCO_ORDINE_TOUR = ""

LINK_TIPO_OGGETTO_TOUR = fnGetTipoOggetto("RV_POTour")
LINK_FUNZIONE_TOUR = GET_FUNZIONE(LINK_TIPO_OGGETTO_TOUR)

sSQL = "SELECT * FROM Semaforo "
sSQL = sSQL & "WHERE IDOggetto=" & LINK_TOUR
sSQL = sSQL & " AND IDTipoOggetto=" & LINK_TIPO_OGGETTO_TOUR
sSQL = sSQL & " AND IDFunzione=" & LINK_FUNZIONE_TOUR

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_BLOCCO_ORDINE_TOUR = "Il tour risulta aperto dall'utente " & GET_UTENTE(fnNotNullN(rs!IDUtente))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_UTENTE(IDUtente As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Utente FROM Utente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_UTENTE = ""
Else
    GET_UTENTE = fnNotNull(rs!Utente)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CAMBIO_POSIZIONE(IDTour As Long, PosizioneOriginale As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * FROM RV_POTourRighe "
'sSQL = sSQL & "WHERE IDRV_POTourRighe<>" & IDRigaTour
sSQL = sSQL & " WHERE Posizione>" & PosizioneOriginale
sSQL = sSQL & " AND IDRV_POTour=" & IDTour
sSQL = sSQL & " ORDER BY Posizione"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
While Not rs.EOF
    rs!Posizione = rs!Posizione - 1
    rs.Update
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Function
Private Sub GET_DATI_TOUR(IDOggetto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT RV_POTour.IDRV_POTour, RV_POTour.Numero, RV_POTour.Anno, RV_POTourRighe.IDRV_POTourRighe, RV_POTourRighe.IDOggettoOrdine, "
sSQL = sSQL & "RV_POTourRighe.Posizione "
sSQL = sSQL & "FROM RV_POTour INNER JOIN "
sSQL = sSQL & "RV_POTourRighe ON RV_POTour.IDRV_POTour = RV_POTourRighe.IDRV_POTour "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TOUR = 0
    LINK_TOUR_RIGHE = 0
    POSIZIONE_TOUR = 0
Else
    LINK_TOUR = fnNotNullN(rs!IDRV_POTour)
    LINK_TOUR_RIGHE = fnNotNullN(rs!IDRV_POTourRighe)
    POSIZIONE_TOUR = fnNotNullN(rs!Posizione)
End If

rs.CloseResultset
Set rs = Nothing



End Sub

Private Function GET_CONTROLLO_LAVORAZIONI_ORDINE(IDOggettoOrdine As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_LAVORAZIONI_ORDINE = ""

sSQL = "SELECT IDRV_POAssegnazioneMerce FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_LAVORAZIONI_ORDINE = ""
Else
    GET_CONTROLLO_LAVORAZIONI_ORDINE = "L'ordine è collegato a delle lavorazioni pertanto è impossibile eliminare il documento"
    
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


Set rs = Cn.OpenResultset(sSQL)

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
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_MODALITA_PAGAMENTO = 0
Else
    GET_MODALITA_PAGAMENTO = fnNotNullN(rs!IDPagamento)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_UM_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraVendita "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo
        
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_UM_ARTICOLO = 8
Else
    GET_LINK_UM_ARTICOLO = fnNotNullN(rs!IDUnitaDiMisuraVendita)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function SCRIVE_NUOVE_RIGHE(IDOggetto As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IRiga As Long
Dim I As Long


sSQL = "SELECT * FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection

IRiga = 1
I = 0
While Not rs.EOF
    oDoc.Tables(sTabellaDettaglio).SetActiveRetail IRiga
    
    For I = 0 To rs.Fields.Count - 1
        If rs.Fields(I).Name <> "IDValoriOggettoDettaglio" Then
            oDoc.Field rs.Fields(I).Name, rs.Fields(I).Value, sTabellaDettaglio
        End If
    Next
    
    IRiga = IRiga + 1
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

oDoc.PerformTable sTabellaDettaglio, False
'Aggiorna il contenuto della listview degli articoli
sbPopalaListaArticoli False
'Ricalcola il documento
sbCalcolaDocumento


End Function
Private Function GET_LINK_PEDANA_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPedanaPerOrdini "
sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_PEDANA_CLIENTE = 0
Else
    GET_LINK_PEDANA_CLIENTE = fnNotNullN(rs!IDRV_POTipoPedanaPerOrdini)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_TOUR() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT RV_POTour.IDRV_POTour "
sSQL = sSQL & "FROM RV_POTour INNER JOIN "
sSQL = sSQL & "RV_POTourRighe ON RV_POTour.IDRV_POTour = RV_POTourRighe.IDRV_POTour "
sSQL = sSQL & "WHERE IDOggettoOrdine=" & oDoc.IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_TOUR = 0
Else
    GET_LINK_TOUR = fnNotNullN(rs!IDRV_POTour)
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Sub SCRIVI_RICERCA_RIGHE(IDTour As Long)
'On Error GoTo ERR_SCRIVI_RICERCA_RIGHE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'''''VARIABILI DI RICERCA'''''''''''''''''''''''''''''''''''''
Dim RIC_NUMERO_ORDINE As String
Dim RIC_DATA_ORDINE As String
Dim RIC_CLIENTE_ORDINE As String
Dim RIC_DEST_ORDINE As String
Dim RIC_DATA_ARRIVO_DEST_ORDINE As String
Dim RIC_ORA_ARRIVO_DEST_ORDINE As String
Dim RIC_LUOGO_ORDINE As String
Dim RIC_DATA_ARRIVO_LUOGO_ORDINE As String
Dim RIC_ORA_ARRIVO_LUOGO_ORDINE As String
Dim RIC_ARTICOLI_ORDINE As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


RIC_NUMERO_ORDINE = ""
RIC_DATA_ORDINE = ""
RIC_CLIENTE_ORDINE = ""
RIC_CLIENTE_ORDINE = ""
RIC_CLIENTE_ORDINE = ""
RIC_DEST_ORDINE = ""
RIC_DATA_ARRIVO_DEST_ORDINE = ""
RIC_ORA_ARRIVO_DEST_ORDINE = ""
RIC_LUOGO_ORDINE = ""
RIC_DATA_ARRIVO_LUOGO_ORDINE = ""
RIC_ORA_ARRIVO_LUOGO_ORDINE = ""

''''''''''''''''''''''RICERCA DI TESTA DELL'ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIETourRicerca "
sSQL = sSQL & "WHERE IDRV_POTour=" & IDTour

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    While Not rs.EOF
        RIC_NUMERO_ORDINE = RIC_NUMERO_ORDINE & fnNotNull(rs!Doc_numero) & "|"
        RIC_DATA_ORDINE = RIC_DATA_ORDINE & fnNotNull(rs!Doc_data) & "|"
        RIC_CLIENTE_ORDINE = RIC_CLIENTE_ORDINE & fnNotNull(rs!Nom_nome) & "|"
        RIC_CLIENTE_ORDINE = RIC_CLIENTE_ORDINE & fnNotNull(rs!Nom_ragione_sociale_o_cognome) & "|"
        RIC_CLIENTE_ORDINE = RIC_CLIENTE_ORDINE & fnNotNull(rs!Nom_codice) & "|"
        
        RIC_DEST_ORDINE = RIC_DEST_ORDINE & fnNotNull(rs!DestinazioneDiversa) & "|"
        RIC_DATA_ARRIVO_DEST_ORDINE = RIC_DATA_ARRIVO_DEST_ORDINE & fnNotNull(rs!RV_PODataArrivoMerce) & "|"
        RIC_ORA_ARRIVO_DEST_ORDINE = RIC_ORA_ARRIVO_DEST_ORDINE & fnNotNull(rs!RV_POOraArrivoMerce) & "|"
        
        RIC_LUOGO_ORDINE = RIC_LUOGO_ORDINE & fnNotNull(rs!LuogoPresaMerce)
        RIC_DATA_ARRIVO_LUOGO_ORDINE = RIC_DATA_ARRIVO_LUOGO_ORDINE & fnNotNull(rs!RV_PODataArrivoMerceLuogo)
        RIC_ORA_ARRIVO_LUOGO_ORDINE = RIC_ORA_ARRIVO_LUOGO_ORDINE & fnNotNull(rs!RV_POOraArrivoMerceLuogo)
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''RICERCA DI ARTICOLO DELL'ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_POIETourRicercaArt "
sSQL = sSQL & "WHERE IDRV_POTour=" & IDTour

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    While Not rs.EOF

        RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!Art_codice) & "|"
        RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!Art_descrizione) & "|"
        If fnNotNullN(rs!RV_POTipoRiga) = 1 Then
            RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!RV_POCodiceTipoPedana) & "|"
            RIC_ARTICOLI_ORDINE = RIC_ARTICOLI_ORDINE & fnNotNull(rs!RV_PODescrizioneTipoPedana) & "|"
        End If
        
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''AGGIORNAMENTO RICERCA PER TOUR''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "UPDATE RV_POTour SET "
sSQL = sSQL & "NumeroOrdineRic=" & fnNormString(RIC_NUMERO_ORDINE) & ", "
sSQL = sSQL & "DataOrdineRic=" & fnNormString(RIC_DATA_ORDINE) & ", "
sSQL = sSQL & "ClienteOrdineRic=" & fnNormString(RIC_CLIENTE_ORDINE) & ", "
sSQL = sSQL & "DestinazioneDiversaRic=" & fnNormString(RIC_DEST_ORDINE) & ", "
sSQL = sSQL & "DataArrivoMerceDestRic=" & fnNormString(RIC_DATA_ARRIVO_DEST_ORDINE) & ", "
sSQL = sSQL & "OraArrivoMerceDestRic=" & fnNormString(RIC_ORA_ARRIVO_DEST_ORDINE) & ", "
sSQL = sSQL & "PresaLuogoMerceRic=" & fnNormString(RIC_LUOGO_ORDINE) & ", "
sSQL = sSQL & "DataArrivoLuogoMerceRic=" & fnNormString(RIC_DATA_ARRIVO_LUOGO_ORDINE) & ", "
sSQL = sSQL & "OraArrivoLuogoMerceRic=" & fnNormString(RIC_ORA_ARRIVO_LUOGO_ORDINE) & ", "
sSQL = sSQL & "ArticoloOrdineRic=" & fnNormString(RIC_ARTICOLI_ORDINE)
sSQL = sSQL & " WHERE IDRV_POTour=" & IDTour

Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Exit Sub
ERR_SCRIVI_RICERCA_RIGHE:
    MsgBox Err.Description, vbCritical, "Funzionalità per ricerca righe"
    
End Sub
Private Function GET_ARTICOLO_ANNULLATO(IDArticolo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Annullato FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ARTICOLO_ANNULLATO = False
Else
    GET_ARTICOLO_ANNULLATO = fnNotNullN(rs!Annullato)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_ANDAMENTO_ORDINE_OLD(IDOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroPedaneOrdinate As Double
Dim NumeroPedaneLavorate As Double
Dim Andamento As Long

Me.PBOrdine.Visible = False
Me.txtAndamentoOrdine.Visible = False
Me.PBOrdine.Value = 0
Me.txtAndamentoOrdine.Text = "0%"

NumeroPedaneOrdinate = GET_TOTALI_ORDINE(IDOrdine)
If NumeroPedaneOrdinate = 0 Then Exit Sub

NumeroPedaneLavorate = GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE(IDOrdine)


Me.PBOrdine.Visible = True
Me.txtAndamentoOrdine.Visible = True
Me.PBOrdine.Value = 0
Me.PBOrdine.Max = NumeroPedaneOrdinate
If NumeroPedaneOrdinate < NumeroPedaneLavorate Then
    Me.PBOrdine.Value = NumeroPedaneOrdinate
Else
    Me.PBOrdine.Value = NumeroPedaneLavorate
End If

Andamento = FormatNumber(((NumeroPedaneLavorate / NumeroPedaneOrdinate) * 100), 0)

Me.txtAndamentoOrdine.Text = Andamento & "%"

If Andamento < 50 Then
    Me.txtAndamentoOrdine.BackColor = Me.BackColor
Else
    Me.txtAndamentoOrdine.BackColor = &H8000000D
End If

End Sub

Private Sub GET_ANDAMENTO_ORDINE(IDOrdine As Long)
On Error GoTo ERR_GET_ANDAMENTO_ORDINE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroPedaneOrdinate As Double
Dim NumeroColliOrdinati As Double

Dim NumeroPedaneLavorate As Double
Dim NumeroColliLavorati As Double

Dim AndamentoPedana As Double
Dim AndamentoColli As Double
Dim RigaAndamanentoOrdine As String


'Me.PBOrdine.Visible = False
Me.txtAndamentoOrdine.Visible = False
'Me.PBOrdine.Value = 0
Me.txtAndamentoOrdine.Text = ""
RigaAndamanentoOrdine = ""

NumeroPedaneOrdinate = GET_TOTALI_ORDINE(IDOrdine)
NumeroColliOrdinati = GET_TOTALI_COLLI_ORDINE(IDOrdine)

'If NumeroPedaneOrdinate = 0 Then Exit Sub

NumeroPedaneLavorate = GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE(IDOrdine)
NumeroColliLavorati = GET_NUMERO_COLLI_LAVORATE_PER_ORDINE(IDOrdine)


'Me.PBOrdine.Visible = True
Me.txtAndamentoOrdine.Visible = True
'Me.PBOrdine.Value = 0
'Me.PBOrdine.Max = NumeroPedaneOrdinate
'If NumeroPedaneOrdinate < NumeroPedaneLavorate Then
'    Me.PBOrdine.Value = NumeroPedaneOrdinate
'Else
'    Me.PBOrdine.Value = NumeroPedaneLavorate
'End If

AndamentoPedana = 100
If NumeroPedaneOrdinate > 0 Then
    AndamentoPedana = FormatNumber(((NumeroPedaneLavorate / NumeroPedaneOrdinate) * 100), 2)
End If

AndamentoColli = 100
If NumeroColliOrdinati > 0 Then
    AndamentoColli = FormatNumber(((NumeroColliLavorati / NumeroColliOrdinati) * 100), 2)
End If

RigaAndamanentoOrdine = RigaAndamanentoOrdine & FormatNumber(NumeroPedaneOrdinate, 2) & "/" & FormatNumber(NumeroPedaneLavorate, 2) & " [" & AndamentoPedana & "%]"
RigaAndamanentoOrdine = RigaAndamanentoOrdine & " - " & FormatNumber(NumeroColliOrdinati, 2) & "/" & FormatNumber(NumeroColliLavorati, 2) & " [" & AndamentoColli & "%]"

Me.txtAndamentoOrdine.Visible = True
Me.txtAndamentoOrdine.Text = RigaAndamanentoOrdine

'If Andamento < 50 Then
'    Me.txtAndamentoOrdine.BackColor = Me.BackColor
'Else
'    Me.txtAndamentoOrdine.BackColor = &H8000000D
'End If


Exit Sub
ERR_GET_ANDAMENTO_ORDINE:
    MsgBox Err.Description, vbCritical, "GET_ANDAMENTO_ORDINE"
End Sub

Private Sub GET_ANDAMENTO_ORDINE_LAVORAZIONE(IDOrdine As Long, IDArticoloLavorato As Long, raggrOrd As String)
On Error GoTo ERR_GET_ANDAMENTO_ORDINE_LAVORAZIONE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroPedaneOrdinate As Double
Dim NumeroPedaneLavorate As Double
Dim NumeroColliOrdinati As Double
Dim NumeroColliLavorati As Double

Dim AndamentoPedane As Double
Dim AndamentoColli As Double
Dim RigaAndamanentoOrdine As String

Dim IDArticoloPadreOrdine As Long


RigaAndamanentoOrdine = ""

CREA_RECORDSET_PEDANA_TMP


Me.txtAndamentoOrdineDett.Visible = False

Me.txtAndamentoOrdineDett.Text = ""

IDArticoloPadreOrdine = IDArticoloLavorato ' GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticoloLavorato)

If IDArticoloPadreOrdine = 0 Then Exit Sub

NumeroPedaneOrdinate = GET_TOTALI_ORDINE_DETTAGLIO(IDOrdine, IDArticoloPadreOrdine, raggrOrd)
NumeroColliOrdinati = GET_TOTALI_ORDINE_DETTAGLIO_COLLI(IDOrdine, IDArticoloPadreOrdine, raggrOrd)

'If NumeroPedaneOrdinate = 0 Then Exit Sub

NumeroPedaneLavorate = GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO(IDOrdine, IDArticoloPadreOrdine, raggrOrd)
NumeroColliLavorati = GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO(IDOrdine, IDArticoloPadreOrdine, raggrOrd)


'Me.PBOrdineDettaglio.Visible = False
'Me.txtAndamentoOrdineDett.Visible = False
'Me.PBOrdineDettaglio.Value = 0
'Me.PBOrdineDettaglio.Max = NumeroPedaneOrdinate

'If NumeroPedaneOrdinate < NumeroPedaneLavorate Then
'    Me.PBOrdineDettaglio.Value = NumeroPedaneOrdinate
'Else
'    Me.PBOrdineDettaglio.Value = NumeroPedaneLavorate
'End If

''RAGGRUPPAMENTO ORDINE
RigaAndamanentoOrdine = RigaAndamanentoOrdine & Me.txtRaggrRigaOrdine.Text

''ANDAMENTO PEDANE
AndamentoPedane = 100
If (NumeroPedaneOrdinate) > 0 Then
    AndamentoPedane = FormatNumber(((NumeroPedaneLavorate / NumeroPedaneOrdinate) * 100), 2)
End If
If Len(Trim(RigaAndamanentoOrdine)) > 0 Then
    RigaAndamanentoOrdine = RigaAndamanentoOrdine & " - " & FormatNumber(NumeroPedaneOrdinate, 2) & "/" & FormatNumber(NumeroPedaneLavorate, 2) & " [" & AndamentoPedane & "%]"
Else
    RigaAndamanentoOrdine = RigaAndamanentoOrdine & FormatNumber(NumeroPedaneOrdinate, 2) & "/" & FormatNumber(NumeroPedaneLavorate, 2) & " [" & AndamentoPedane & "%]"
End If

''ANDAMENTO COLLI
AndamentoColli = 100
If (NumeroColliOrdinati) > 0 Then
    AndamentoColli = FormatNumber(((NumeroColliLavorati / NumeroColliOrdinati) * 100), 2)
End If
RigaAndamanentoOrdine = RigaAndamanentoOrdine & " - " & FormatNumber(NumeroColliOrdinati, 2) & "/" & FormatNumber(NumeroColliLavorati, 2) & " [" & AndamentoColli & "%]"

'Me.txtAndamentoOrdineDett.Text = Andamento & "%"
Me.txtAndamentoOrdineDett.Visible = True
Me.txtAndamentoOrdineDett.Text = RigaAndamanentoOrdine

'If Andamento < 50 Then
'    Me.txtAndamentoOrdineDett.BackColor = Me.BackColor
'Else
'    Me.txtAndamentoOrdineDett.BackColor = &H8000000D
'End If

Exit Sub

ERR_GET_ANDAMENTO_ORDINE_LAVORAZIONE:
    MsgBox Err.Description, vbCritical, "GET_ANDAMENTO_ORDINE_LAVORAZIONE"
End Sub

Private Function GET_TOTALI_ORDINE(IDOggettoOrdine As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(RV_POQuantitaPedanaEffettiva) as NumeroPedane "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=1 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALI_ORDINE = 0
Else
    GET_TOTALI_ORDINE = fnNotNullN(rs!NumeroPedane)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE = 0

sSQL = "SELECT IDRV_POPedana "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
sSQL = sSQL & " GROUP BY IDRV_POPedana"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE = GET_NUMERO_PEDANE_LAVORATE_PER_ORDINE + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ParametroAndamentoOrdine()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT VisAndamentoOrdDaOrd FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    VISUALIZZA_ANDAMENTO_ORDINE = fnNotNullN(rs!VisAndamentoOrdDaOrd)
Else
    VISUALIZZA_ANDAMENTO_ORDINE = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroGestioneOrdineVivaio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivaGestioneOrdineVivaio FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GESTIONE_ORDINE_VIVAIO = fnNotNullN(rs!AttivaGestioneOrdineVivaio)
Else
    GESTIONE_ORDINE_VIVAIO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_ANDAMENTO_ORDINE_LAVORAZIONE_OLD(IDOrdine As Long, IDArticoloLavorato As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroPedaneOrdinate As Double
Dim NumeroPedaneLavorate As Double
Dim Andamento As Long
Dim IDArticoloPadreOrdine As Long

CREA_RECORDSET_PEDANA_TMP

'Me.PBOrdineDettaglio.Visible = False
Me.txtAndamentoOrdineDett.Visible = False
'Me.PBOrdineDettaglio.Value = 0
Me.txtAndamentoOrdineDett.Text = "0%"

IDArticoloPadreOrdine = IDArticoloLavorato
If IDArticoloPadreOrdine = 0 Then Exit Sub


NumeroPedaneOrdinate = GET_TOTALI_ORDINE_DETTAGLIO(IDOrdine, IDArticoloPadreOrdine, txtRaggrRigaOrdine.Text)

If NumeroPedaneOrdinate = 0 Then Exit Sub


NumeroPedaneLavorate = GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO(IDOrdine, IDArticoloPadreOrdine, txtRaggrRigaOrdine.Text)


'Me.PBOrdineDettaglio.Visible = True
Me.txtAndamentoOrdineDett.Visible = True
'Me.PBOrdineDettaglio.Value = 0
'Me.PBOrdineDettaglio.Max = NumeroPedaneOrdinate

If NumeroPedaneOrdinate < NumeroPedaneLavorate Then
    'Me.PBOrdineDettaglio.Value = NumeroPedaneOrdinate
Else
    'Me.PBOrdineDettaglio.Value = NumeroPedaneLavorate
End If

Andamento = FormatNumber(((NumeroPedaneLavorate / NumeroPedaneOrdinate) * 100), 0)

Me.txtAndamentoOrdineDett.Text = Andamento & "%"
If Andamento < 50 Then
    Me.txtAndamentoOrdineDett.BackColor = Me.BackColor
Else
    Me.txtAndamentoOrdineDett.BackColor = &H8000000D
End If

End Sub
Private Function GET_LINK_ARTICOLO_PADRE_ORDINATO(IDArticoloDerivato As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo FROM RV_POArticoloFiglioOrdine "
sSQL = sSQL & "WHERE IDArticoloFiglio=" & IDArticoloDerivato

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ARTICOLO_PADRE_ORDINATO = 0
Else
    GET_LINK_ARTICOLO_PADRE_ORDINATO = fnNotNullN(rs!IDArticolo)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TOTALI_ORDINE_DETTAGLIO(IDOggettoOrdine As Long, IDArticolo As Long, raggrOrd As String) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(RV_POQuantitaPedanaEffettiva) as NumeroPedane "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND Link_art_articolo=" & IDArticolo
sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(raggrOrd)
sSQL = sSQL & " AND RV_POTipoRiga=1 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALI_ORDINE_DETTAGLIO = 0
Else
    GET_TOTALI_ORDINE_DETTAGLIO = fnNotNullN(rs!NumeroPedane)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREA_RECORDSET_PEDANA_TMP()
If Not (rsTmpPed Is Nothing) Then
    If rsTmpPed.State > 0 Then
        rsTmpPed.Close
    End If
    Set rsTmpPed = Nothing
End If
Set rsTmpPed = New ADODB.Recordset

rsTmpPed.CursorLocation = adUseClient

rsTmpPed.Fields.Append "IDRV_POPedana", adInteger, , adFldIsNullable

rsTmpPed.Open , , adOpenKeyset, adLockPessimistic


End Sub
Private Function GET_CONTROLLO_ESISTENZA_PEDANA_CALCOLATA(IDPedana As Long) As Boolean
rsTmpPed.Filter = "IDRV_POPedana=" & IDPedana

If rsTmpPed.EOF Then
    GET_CONTROLLO_ESISTENZA_PEDANA_CALCOLATA = False
    rsTmpPed.Filter = ""
    rsTmpPed.AddNew
        rsTmpPed!IDRV_POPedana = IDPedana
    rsTmpPed.Update

Else
    GET_CONTROLLO_ESISTENZA_PEDANA_CALCOLATA = True
End If
End Function
Private Function GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO(IDOggettoOrdine As Long, IDArticoloPadre As Long, raggrOrd As String) As Double
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset
Dim rstmp As ADODB.Recordset


GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO = 0

'sSQL = "SELECT * FROM RV_POArticoloFiglioOrdine "
'sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloPadre
'
'Set rsArt = Cn.OpenResultset(sSQL)
'
'While Not rsArt.EOF

    sSQL = "SELECT IDRV_POPedana "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDOggettoOrdine=" & IDOggettoOrdine
    sSQL = sSQL & " AND IDArticolo=" & IDArticoloPadre 'fnNotNullN(rsArt!IDArticoloFiglio)
    sSQL = sSQL & " AND NotaRigaOrdRaggr=" & fnNormString(raggrOrd)
    sSQL = sSQL & " GROUP BY IDRV_POPedana"
    
    Set rs = Cn.OpenResultset(sSQL)
        
    While Not rs.EOF
        If GET_CONTROLLO_ESISTENZA_PEDANA_CALCOLATA(fnNotNullN(rs!IDRV_POPedana)) = False Then
            GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO = GET_NUMERO_PEDANE_LAVORATE_PER_ARTICOLO_ORDINATO + 1
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
'rsArt.MoveNext
'Wend

'rsArt.CloseResultset
'Set rsArt = Nothing
End Function

Private Function GET_PREZZO_IMBALLO_INCLUSO(IDArticoloImballo As Long, idcliente As Long) As Long
On Error GoTo ERR_GET_PREZZO_IMBALLO_INCLUSO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsCli As DmtOleDbLib.adoResultset


sSQL = "SELECT PrezzoInclusoImballo "
sSQL = sSQL & "FROM RV_POConfigurazioneClienteImb "
sSQL = sSQL & "WHERE IDAnagrafica=" & idcliente
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    sSQL = "SELECT PrezzoInclusoImballo "
    sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
    sSQL = sSQL & "WHERE IDAnagrafica=" & idcliente
    sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
    'sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    
    Set rsCli = Cn.OpenResultset(sSQL)
    
    If rsCli.EOF Then
        GET_PREZZO_IMBALLO_INCLUSO = 0
    Else
        GET_PREZZO_IMBALLO_INCLUSO = fnNotNullN(rsCli!PrezzoInclusoImballo)
    End If
    
    rsCli.CloseResultset
    Set rsCli = Nothing
    
Else
    GET_PREZZO_IMBALLO_INCLUSO = fnNotNullN(rs!PrezzoInclusoImballo)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_PREZZO_IMBALLO_INCLUSO:
MsgBox Err.Description, vbCritical, "ERR_GET_PREZZO_IMBALLO_INCLUSO"
End Function

Private Function GET_CONTROLLO_NUMERO_LETTERE_INTENTO(IDAnagrafica As Long, IDAzienda As Long, Anno As Long) As Long
On Error GoTo ERR_GET_CONTROLLO_NUMERO_LETTERE_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Count(IDLetteraIntento) AS NumeroRecord "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDAzienda_CF=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica_CF=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica_CF=2"
sSQL = sSQL & " AND ((Anno=" & Anno & ") OR (AnnoOperazione=" & Anno & "))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_NUMERO_LETTERE_INTENTO = 0
Else
    GET_CONTROLLO_NUMERO_LETTERE_INTENTO = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_NUMERO_LETTERE_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_NUMERO_LETTERE_INTENTO"
End Function
Private Function GET_LINK_LETTERA_INTENTO(IDAnagrafica As Long, IDAzienda As Long, Anno As Long) As Long
On Error GoTo ERR_GET_LINK_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDLetteraIntento "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDAzienda_CF=" & IDAzienda
sSQL = sSQL & " AND IDAnagrafica_CF=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica_CF=2"
sSQL = sSQL & " AND ((Anno=" & Anno & ") OR (AnnoOperazione=" & Anno & "))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LETTERA_INTENTO = 0
Else
    GET_LINK_LETTERA_INTENTO = fnNotNullN(rs!IDLetteraIntento)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_LETTERA_INTENTO"
End Function
Private Function GET_LINK_IVA_LETTERA_INTENTO(IDLetteraIntento As Long, IDIvaCliente As Long) As Long
On Error GoTo ERR_GET_LINK_IVA_LETTERA_INTENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
Else
    If fnNotNullN(rs!IDIva) > 0 Then
        GET_LINK_IVA_LETTERA_INTENTO = fnNotNullN(rs!IDIva)
    Else
        GET_LINK_IVA_LETTERA_INTENTO = IDIvaCliente
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_LINK_IVA_LETTERA_INTENTO:
    MsgBox Err.Description, vbCritical, "GET_LINK_IVA_LETTERA_INTENTO"
End Function

Private Function GET_LINK_IVA_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_CLIENTE = 0
Else

    GET_LINK_IVA_CLIENTE = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function

End Function
Private Sub AggiornaAccordiCommerciali()
    With Me.cboAccordoCommerciale
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDAccordiCommerciali"
        .DisplayField = "Descrizione"
        .SQL = "SELECT IDAccordiCommerciali, Descrizione FROM AccordiCommerciali"
        .SQL = .SQL & " WHERE IDAnagrafica = " & Me.cdAnagrafica.KeyFieldID
        .SQL = .SQL & " AND IDTipoAnagrafica=" & oDoc.IDTipoAnagrafica
        .SQL = .SQL & " ORDER BY Descrizione"
    End With
End Sub
Private Function GET_LINK_ACCORDO_COMMERCIALE_PREDEFINITO(IDAnagrafica As Long, IDTipoAnagrafica As Long, DataControllo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAccordiCommerciali FROM AccordiCommerciali "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDTipoAnagrafica=" & IDTipoAnagrafica
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(DataControllo)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(DataControllo)
sSQL = sSQL & " AND Chiuso=" & fnNormBoolean(0)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ACCORDO_COMMERCIALE_PREDEFINITO = 0
Else
    GET_LINK_ACCORDO_COMMERCIALE_PREDEFINITO = fnNotNullN(rs!IDAccordiCommerciali)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_ACCORDO_CHIUSO(IDAccordoCommerciale As Long) As Long
On Error GoTo ERR_GET_CONTROLLO_ACCORDO_CHIUSO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAccordiCommerciali, Chiuso FROM AccordiCommerciali "
sSQL = sSQL & "WHERE IDAccordiCommerciali=" & IDAccordoCommerciale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ACCORDO_CHIUSO = 1
Else
    GET_CONTROLLO_ACCORDO_CHIUSO = Abs(fnNotNullN(rs!Chiuso))
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_ACCORDO_CHIUSO:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_ACCORDO_CHIUSO"
    GET_CONTROLLO_ACCORDO_CHIUSO = 1
End Function
Private Function GET_CONTROLLO_ACCORDO_PER_DATA(IDAccordoCommerciale As Long, DataControllo As String) As Long
On Error GoTo ERR_GET_CONTROLLO_ACCORDO_PER_DATA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAccordiCommerciali FROM AccordiCommerciali "
sSQL = sSQL & "WHERE IDAccordiCommerciali=" & IDAccordoCommerciale
sSQL = sSQL & " AND DataInizio<=" & fnNormDate(DataControllo)
sSQL = sSQL & " AND DataFine>=" & fnNormDate(DataControllo)
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_ACCORDO_PER_DATA = 1
Else
    GET_CONTROLLO_ACCORDO_PER_DATA = 0
End If
rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CONTROLLO_ACCORDO_PER_DATA:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_ACCORDO_PER_DATA"
    GET_CONTROLLO_ACCORDO_PER_DATA = 0
End Function
Private Function GET_LINK_AGENTE_CLIENTE(IDAnagrafica As Long, IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagraficaAgente FROM ProvvClienteAgente "
sSQL = sSQL & " WHERE IDAziendaCliente=" & IDAzienda
sSQL = sSQL & " AND IDTipoAnagraficaCliente=2"
sSQL = sSQL & " AND IDAnagraficaCliente=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_AGENTE_CLIENTE = 0
Else
    GET_LINK_AGENTE_CLIENTE = fnNotNullN(rs!IDAnagraficaAgente)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_CODICE_AGENTE(IDAnagraficaAgente As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Codice FROM IERepAgente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaAgente

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CODICE_AGENTE = ""
Else
    GET_CODICE_AGENTE = fnNotNull(rs!Codice)
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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REGOLA_PROVV_AGE = 0
Else
    GET_LINK_REGOLA_PROVV_AGE = fnNotNullN(rs!IDRegolaProvv)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_DATI_TOUR(IDOggetto As Long, IDVettore As Long, TargaAutomezzo As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_DATI_TOUR = False

sSQL = "SELECT RV_POTour.IDRV_POTour, RV_POTour.IDAzienda, RV_POTour.IDFiliale, RV_POTour.IDVettore, RV_POTour.IDRV_POTipoMezzoTrasporto, "
sSQL = sSQL & "RV_POTour.TargaMezzoTrasporto, RV_POTour.Importo, RV_POTour.DataPartenza, RV_POTour.Anno, RV_POTour.Numero,"
sSQL = sSQL & "RV_POTour.DataInserimento, RV_POTour.IDUtenteInserimento, RV_POTour.DataUltimaModifica, RV_POTour.IDUtenteUltimaModifica,"
sSQL = sSQL & "RV_POTour.IDRV_POStatoTour, RV_POTour.TelefonoAutista, RV_POTour.AnnotazioniTour, RV_POTour.NumeroOrdineRic, RV_POTour.DataOrdineRic,"
sSQL = sSQL & "RV_POTour.ClienteOrdineRic, RV_POTour.DestinazioneDiversaRic, RV_POTour.DataArrivoMerceDestRic, RV_POTour.OraArrivoMerceDestRic,"
sSQL = sSQL & "RV_POTour.PresaLuogoMerceRic, RV_POTour.DataArrivoLuogoMerceRic, RV_POTour.OraArrivoLuogoMerceRic, RV_POTour.ArticoloOrdineRic,"
sSQL = sSQL & "RV_POTour.IDAnagraficaFornitore , RV_POTourRighe.IDOggettoOrdine "
sSQL = sSQL & "FROM RV_POTour INNER JOIN "
sSQL = sSQL & "RV_POTourRighe ON RV_POTour.IDRV_POTour = RV_POTourRighe.IDRV_POTour "
sSQL = sSQL & "WHERE RV_POTourRighe.IDOggettoOrdine=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If IDVettore <> fnNotNullN(rs!IDVettore) Then
        GET_CONTROLLO_DATI_TOUR = True
    End If
    
    If GET_CONTROLLO_DATI_TOUR = False Then
        If TargaAutomezzo <> fnNotNull(rs!TargaMezzoTrasporto) Then
            GET_CONTROLLO_DATI_TOUR = True
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub GET_CONFIGURAZIONE_IMPORTI(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double)
Dim ObjDoc As DmtDocs.cDocument
Dim sTabellaTestataLocal As String
Dim sTabellaDettaglioLocal As String
Dim sTabellaIVALocal As String
Dim sTabellaScadenzeLocal As String

    

    Set ObjDoc = New DmtDocs.cDocument
    
    Set ObjDoc.Connection = TheApp.Database.Connection
    ObjDoc.SetTipoOggetto oDoc.IDTipoOggetto
    ObjDoc.IDFunzione = oDoc.IDFunzione
    ObjDoc.TablesNames oDoc.IDTipoOggetto, sTabellaTestataLocal, sTabellaDettaglioLocal, sTabellaIVALocal, sTabellaScadenzeLocal
    ObjDoc.IDAzienda = oDoc.IDAzienda
    ObjDoc.IDFiliale = oDoc.IDFiliale
    ObjDoc.IDAttivitaAzienda = oDoc.IDAttivitaAzienda
    ObjDoc.IDTipoAnagrafica = oDoc.IDTipoAnagrafica
    ObjDoc.IDUtente = oDoc.IDUtente
    ObjDoc.DataEmissione = oDoc.DataEmissione
    ObjDoc.ReadDataFromCliFo IDAnagrafica
    ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
    ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal
    
    ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail oDoc.Tables(sTabellaDettaglioLocal).NumRetails
    ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
    ObjDoc.ReadDataFromArticle IDArticolo, sTabellaDettaglioLocal

    
    ObjDoc.ReadDataFromPriceList IDListino
    ObjDoc.ReadDataFromDiscountsList
    
'    Dim idlistinoapplicato As Long
'    idlistinoapplicato = ObjDoc.ReadDataFromDiscountsListWithReturnID
    
    If Quantita = 0 Then
        ObjDoc.Field "Art_quantita_totale", "001", sTabellaDettaglioLocal
    Else
        ObjDoc.Field "Art_quantita_totale", Quantita, sTabellaDettaglioLocal
       
    End If
    
    Me.txtImportoUnitarioArticolo.Value = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))
    Me.txtImportoListinoArticolo.Value = Me.txtImportoUnitarioArticolo.Value
    Me.txtScontoImpListino.Value = 0
    
    IMPORTO_UNITARIO_LISTINO = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))
    
    Me.txtSconto1.Value = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglioLocal))
    Me.txtSconto2.Value = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglioLocal))

    
    Set ObjDoc = Nothing
End Sub


Private Sub GET_CONFIGURAZIONE_IMPORTI_IMBALLI(IDAnagrafica As Long, IDArticolo As Long, IDListino As Long, IDListinoAzienda As Long, Quantita As Double)
On Error GoTo ERR_GET_CONFIGURAZIONE_IMPORTI_IMBALLI
Dim ObjDoc As DmtDocs.cDocument
Dim sTabellaTestataLocal As String
Dim sTabellaDettaglioLocal As String
Dim sTabellaIVALocal As String
Dim sTabellaScadenzeLocal As String

    Set ObjDoc = New DmtDocs.cDocument
    Set ObjDoc.Connection = TheApp.Database.Connection
    ObjDoc.SetTipoOggetto oDoc.IDTipoOggetto
    ObjDoc.IDFunzione = oDoc.IDFunzione
    ObjDoc.TablesNames oDoc.IDTipoOggetto, sTabellaTestataLocal, sTabellaDettaglioLocal, sTabellaIVALocal, sTabellaScadenzeLocal
    ObjDoc.IDAzienda = oDoc.IDAzienda
    ObjDoc.IDFiliale = oDoc.IDFiliale
    ObjDoc.IDAttivitaAzienda = oDoc.IDAttivitaAzienda
    ObjDoc.IDTipoAnagrafica = oDoc.IDTipoAnagrafica
    ObjDoc.IDUtente = oDoc.IDUtente
    ObjDoc.DataEmissione = Me.dtData.Text

    ObjDoc.ClearValues
    
    
    ObjDoc.Tables(sTabellaDettaglioLocal).SetActiveRetail oDoc.Tables(sTabellaDettaglioLocal).NumRetails
    ObjDoc.ReadDataFromCliFo IDAnagrafica
    ObjDoc.Field "Link_Val_valuta", 9, sTabellaTestataLocal
    ObjDoc.ReadDataFromArticle IDArticolo, sTabellaDettaglioLocal
    ObjDoc.Field "Link_Doc_listino", IDListino, sTabellaTestataLocal
    ObjDoc.Field "Link_Doc_listino_base", IDListinoAzienda, sTabellaTestataLocal


    
    ObjDoc.ReadDataFromPriceList IDListino
    ObjDoc.ReadDataFromDiscountsList
    
    If Quantita = 0 Then
        ObjDoc.Field "Art_quantita_totale", "0,01", sTabellaDettaglioLocal
    Else
        ObjDoc.Field "Art_quantita_totale", Quantita, sTabellaDettaglioLocal
    End If
    
    Me.txtImportoUnitarioImballo.Value = fnNotNullN(ObjDoc.Field("Art_prezzo_unitario_neutro", , sTabellaDettaglioLocal))
    
    'Me.txtSconto1.Value = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglioLocal))
    'Me.txtSconto2.Value = fnNotNullN(ObjDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglioLocal))

    
Set ObjDoc = Nothing
Exit Sub
ERR_GET_CONFIGURAZIONE_IMPORTI_IMBALLI:
    MsgBox Err.Description, vbCritical, "GET_CONFIGURAZIONE_IMPORTI_IMBALLI"
End Sub

Private Sub GET_MODULO_ATTIVATO(Codice As String, IdentificativoProgramma As Long)
On Error GoTo ERR_GET_MODULO_ATTIVATO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Attivato, DescrizioneModulo FROM RV_POProgrammaModulo "
sSQL = sSQL & "WHERE CodiceModulo=" & fnNormString(Codice)
sSQL = sSQL & " AND IdentificazioneProgramma=" & IdentificativoProgramma

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
Else
    MODULO_ATTIVATO = Abs(fnNotNullN(rs!Attivato))
    MODULO_DESCRIZIONE = fnNotNull(rs!DescrizioneModulo)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_MODULO_ATTIVATO:
    MODULO_ATTIVATO = 0
    MODULO_DESCRIZIONE = ""
End Sub
Private Function fncIDTipoOggettoPrg(Gestore As String) As Long
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto, Gestore.Gestore"
    sSQL = sSQL & " FROM Gestore INNER JOIN TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore"
    sSQL = sSQL & " WHERE (((Gestore.Gestore)=" & fnNormString(Gestore) & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
        
    If rs.EOF = False Then
        fncIDTipoOggettoPrg = rs!IDTipoOggetto
    Else
        fncIDTipoOggettoPrg = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Function fncImpostaDefaultReport(ByVal IDReportDefault As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fncImpostaDefaultReport
    Dim sSQL As String
    
    sSQL = "UPDATE DefaultFilialePerTipoOggetto SET "
    sSQL = sSQL & "IDReportTipoOggetto=" & IDReportDefault
    sSQL = sSQL & " WHERE IDTipoOggetto = " & IDTipoOggetto
    sSQL = sSQL & " AND IDFiliale = " & TheApp.Branch
    
    Cn.Execute sSQL
    
Exit Function
ERR_fncImpostaDefaultReport:
    MsgBox Err.Description, vbCritical, "Settaggio report di default"
End Function
Public Sub GET_CALCOLO_VIVAIO(QuantitaVenduta As Double)
Dim QuantitaPiantePerCarrello As Long
Dim NumeroCarrelli As Double
Dim NumeroPianali As Double
Dim NumeroProlunghe As Double

QuantitaPiantePerCarrello = N_PIANALI_PER_CARRELLO * N_PIANTE_PER_PIANALE

NumeroCarrelli = 1
NumeroPianali = 1
NumeroProlunghe = 0

If (QuantitaPiantePerCarrello > 0) Then
    NumeroCarrelli = QuantitaVenduta / QuantitaPiantePerCarrello
    NumeroPianali = fnRoundUp((QuantitaVenduta / N_PIANTE_PER_PIANALE) - fnRoundUp(NumeroCarrelli))
    NumeroProlunghe = fnRoundUp(fnRoundUp(NumeroCarrelli) * N_PROLUNGHE_PER_CARRELLO)
End If

Me.txtQuantitaPedana.Value = NumeroCarrelli
Me.txtQtaPianale.Value = NumeroPianali
Me.txtQtaProlunga.Value = NumeroProlunghe

End Sub
Private Sub DISEGNA_FORM()

If GESTIONE_ORDINE_VIVAIO = 1 Then
    fraImballo.Top = 800
    fraPedana.Top = Me.fraImballo.Top + Me.fraImballo.Height - 120
    Me.fraVivaio.Visible = True
    Me.fraAnnotazioniPerSocio.Top = Me.fraVivaio.Top + Me.fraVivaio.Height - 120
    Me.fraAnnotazioni.Top = Me.fraAnnotazioniPerSocio.Top + Me.fraAnnotazioniPerSocio.Height - 120
    fraOrdineRigaOri.Visible = True
    Me.fraAnnotazioni.ZOrder 0
    Me.fraTotaliPedane.Visible = False
Else
    
    Me.fraAnnotazioniPerSocio.Visible = False
    fraPedana.Top = 800
    fraImballo.Top = Me.fraPedana.Top + Me.fraPedana.Height - 120
    fraPedana.ZOrder 0
    
    Me.CDTipoPedana.TabIndex = 43
    Me.CDPedana.TabIndex = 44
    Me.txtQuantitaPedana.TabIndex = 45
    Me.txtColliSfusi.TabIndex = 46
    Me.cboUMRigaOrdine.TabIndex = 47
    
    Me.fraVivaio.Visible = False
    Me.fraAnnotazioni.Top = Me.fraArticoli.Top + Me.fraArticoli.Height - 20
    Me.Frame2.Top = Me.fraAnnotazioni.Top
    Me.fraAnnotazioni.ZOrder 0
    Me.fraTotaliPedane.Visible = True
    Me.fraOrdineRigaOri.Visible = False
    
    Me.fraTotaliPedane.Top = Frame10.Top + Frame10.Height - 60
End If

If VISUALIZZA_ANDAMENTO_ORDINE = 1 Then
    Me.txtAndamentoOrdineDett.Visible = True
Else
    txtAndamentoOrdineDett.Visible = False
End If

txtAndamentoOrdineDett.Top = (Me.fraAnnotazioni.Top + Me.fraAnnotazioni.Height) + 120

If VISUALIZZA_ANDAMENTO_ORDINE = 0 Then
    Me.lvwArticoli.Height = Me.FraTab(3).Height - (Me.fraAnnotazioni.Top + Me.fraAnnotazioni.Height) - 360
    Me.lvwArticoli.Top = (Me.fraAnnotazioni.Top + Me.fraAnnotazioni.Height) + 120
Else
    Me.lvwArticoli.Height = Me.FraTab(3).Height - (Me.txtAndamentoOrdineDett.Top + Me.txtAndamentoOrdineDett.Height) - 360
    Me.lvwArticoli.Top = (Me.txtAndamentoOrdineDett.Top + Me.txtAndamentoOrdineDett.Height) + 120
End If
End Sub
Private Function GET_LINK_FORNITORE(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica FROM FornitorePerArticolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND LivelloPreferenza=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_FORNITORE = 0
Else
    GET_LINK_FORNITORE = fnNotNullN(rs!IDAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub GET_PARAMETRI_IVA_IMBALLO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IvaBloccata, IvaImballoARendere "
sSQL = sSQL & "FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDUtente=" & 0

Set rs = Cn.OpenResultset(sSQL)

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

Private Function GET_SCONTO_IMPORTO_LISTINO()
Dim Value As Double
Value = 0

If (txtImportoListinoArticolo.Value > 0) Then
    Value = (1 - (Me.txtImponibileUnitario.Value / txtImportoListinoArticolo.Value)) * 100
End If

'Value = fnRoundChange(Value, 0.25, 1)

GET_SCONTO_IMPORTO_LISTINO = Value

End Function
Private Sub GET_TOTALI_PEDANE()
 Dim lRow As Long
 Dim TotaliPedane As Double
 Dim TotaliPedaneEff As Double
 
TotaliPedane = 0
TotaliPedaneEff = 0
    With oDoc.Tables(sTabellaDettaglio)
        'Cicla per tutte le righe di dettaglio presenti nel documento
        If .NumRetails = 0 Then Exit Sub
        
        For lRow = 1 To .NumRetails
            If (fnNotNullN(.Fields("RV_POTipoRiga").Values(lRow).Value)) = 1 Then
                TotaliPedane = TotaliPedane + fnNotNullN(.Fields("RV_POQuantitaPedana").Values(lRow).Value)
                TotaliPedaneEff = TotaliPedaneEff + fnNotNullN(.Fields("RV_POQuantitaPedanaEffettiva").Values(lRow).Value)
            End If
        Next
        
    End With
    
    Me.txtTotalePedane.Value = TotaliPedane
    Me.txtTotalePedaneEff.Value = TotaliPedaneEff
    
End Sub
Private Function GET_DESCRIZIONE_COLLEGAMENTO(IDRigaOrdine As Long) As String
On Error GoTo ERR_GET_DESCRIZIONE_COLLEGAMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_DESCRIZIONE_COLLEGAMENTO = ""

sSQL = "SELECT * FROM RV_POIERigheOrdineConferimento "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDRigaOrdine

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!IDValoriOggettoDettaglioOrdineConf) > 0 Then
        GET_DESCRIZIONE_COLLEGAMENTO = GET_DESCRIZIONE_COLLEGAMENTO & "Conferimento "
        GET_DESCRIZIONE_COLLEGAMENTO = GET_DESCRIZIONE_COLLEGAMENTO & " n° " & fnNotNull(rs!NumeroConferimento)
        GET_DESCRIZIONE_COLLEGAMENTO = GET_DESCRIZIONE_COLLEGAMENTO & " del " & fnNotNull(rs!DataConferimento)
    End If
End If

rs.CloseResultset
Set rs = Nothing

Exit Function

ERR_GET_DESCRIZIONE_COLLEGAMENTO:
    MsgBox Err.Description, vbCritical, "GET_DESCRIZIONE_COLLEGAMENTO"
End Function

Private Function GET_CONTROLLO_COLLEGAMENTO_CONF(IDRigaOrdine As Long) As Boolean
On Error GoTo ERR_GET_CONTROLLO_COLLEGAMENTO_CONF
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_COLLEGAMENTO_CONF = False

sSQL = "SELECT * FROM RV_POIERigheOrdineConferimento "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio=" & IDRigaOrdine

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!IDValoriOggettoDettaglioOrdineConf) > 0 Then
        GET_CONTROLLO_COLLEGAMENTO_CONF = True
    End If
End If

rs.CloseResultset
Set rs = Nothing

Exit Function

ERR_GET_CONTROLLO_COLLEGAMENTO_CONF:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_COLLEGAMENTO_CONF"
End Function
Private Sub ParametroSezionalePrelievo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionaleListaPrelievoOrdine FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_SEZIONALE_LISTA = fnNotNullN(rs!IDSezionaleListaPrelievoOrdine)
Else
    LINK_SEZIONALE_LISTA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_NUMERO_LISTA_PRELIEVO(IDOrdinePadre) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(RV_PONumeroListaPrelievo) AS NumeroMaxPrelievo "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE RV_POIDOrdinePadre=" & IDOrdinePadre

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_LISTA_PRELIEVO = 1
Else
    GET_NUMERO_LISTA_PRELIEVO = fnNotNullN(rs!NumeroMaxPrelievo) + 1
End If


End Function
Private Function AGGIORNA_LINK_ORDINE_PADRE(IDOggetto As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
sSQL = sSQL & " RV_POIDOrdinePadre=" & IDOggetto
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto

Cn.Execute sSQL

oDoc.Field "RV_POIDOrdinePadre", IDOggetto, sTabellaTestata


End Function

Private Sub ABILITA_CONTROLLI()
Dim bValue As Boolean

If Me.txtNListaPrelievo.Value = 1 Then
    bValue = True
Else
    bValue = False
End If

cboMagazzino.Enabled = bValue
cboTipoOrdine.Enabled = bValue
txtDataOrdineCliente.Enabled = bValue
txtNumeroOrdineCliente.Enabled = bValue
cboBancaAzienda.Enabled = bValue
cboListino.Enabled = bValue
cboListinoAzienda.Enabled = bValue
cboListinoAzienda.Enabled = bValue
cboAccordoCommerciale.Enabled = bValue
chkLordoIVA.Enabled = bValue
chkRaggruppBolle.Enabled = bValue
chkRaggruppaScadenze.Enabled = bValue
chkOrdineCompletato.Enabled = bValue
chkStampaFattProForma.Enabled = bValue
curSpeseIncasso.Enabled = bValue
lngScontoDocPer.Enabled = bValue
curSpeseTrasporto.Enabled = bValue
curScontoDocImp.Enabled = bValue

cdAnagrafica.Enabled = bValue
cboPagamento.Enabled = bValue
cboBancaCliente.Enabled = bValue
cmdEliminaRifLetInt.Enabled = bValue
cmdLetteraIntento.Enabled = bValue
cboIvaCliente.Enabled = bValue
cboValuta.Enabled = bValue
CDAgenteTesta.Enabled = bValue
cboSezionale.Enabled = bValue
ACSCliente.Enabled = bValue
ACSAnaDest.Enabled = bValue



End Sub
Private Sub IMPOSTA_DATI_LISTA_PRELIEVO(IDOggettoPadre As Long)
On Error GoTo ERR_IMPOSTA_DATI_LISTA_PRELIEVO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM " & sTabellaTestata
sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoPadre

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    cboMagazzino.WriteOn fnNotNullN(rs!Link_Doc_magazzino)
    cboTipoOrdine.WriteOn fnNotNullN(rs!RV_POIDTipoOrdine)
    txtDataOrdineCliente.Text = fnNotNull(rs!Doc_data_presso_nom)
    txtNumeroOrdineCliente.Text = fnNotNull(rs!Doc_numero_presso_nom)
    cboBancaAzienda.WriteOn fnNotNullN(rs!Link_Doc_contratto_bancario_az)
    
    cboListino.WriteOn fnNotNullN(rs!Link_Doc_listino)
    cboListinoAzienda.WriteOn fnNotNullN(rs!Link_Doc_listino_base)

    cboAccordoCommerciale.WriteOn fnNotNullN(rs!Link_Nom_accordi_commerciali)
    
    chkLordoIVA.Value = 0
    chkRaggruppBolle.Value = 0
    chkRaggruppaScadenze.Value = 0
    chkOrdineCompletato.Value = 0
    chkStampaFattProForma.Value = fnNotNullN(rs!RV_POFatturaProforma)
    
    curSpeseIncasso.Value = 0
    lngScontoDocPer.Value = 0
    curSpeseTrasporto.Value = 0
    curScontoDocImp.Value = 0
    
    cdAnagrafica.Load fnNotNullN(rs!Link_Nom_anagrafica)
    cboPagamento.WriteOn fnNotNullN(rs!Link_doc_pagamento)
    cboBancaCliente.WriteOn fnNotNullN(rs!Link_Nom_contratto_bancario)
    txtIDLetteraIntento.Value = fnNotNullN(rs!Link_Nom_lettera_intento)
    
    cboIvaCliente.WriteOn fnNotNullN(rs!Link_Nom_IVA)
    cboValuta.WriteOn fnNotNullN(rs!Link_Val_valuta)
    
    CDAgenteTesta.Load fnNotNullN(rs!Link_doc_agente)
    
    cboAltroSito.WriteOn fnNotNullN(rs!Link_Nom_ult_sito)
    cboLuogoPresaMerce.WriteOn fnNotNullN(rs!RV_POIDLuogoPresaMerce)
    
    Me.ACSAnaDest.sbLoadCFByIDAnagrafica 0, fnNotNullN(rs!RV_POIDAnagraficaDestinazione)

    
    sbImpostaDatiDocumento

End If
Exit Sub
ERR_IMPOSTA_DATI_LISTA_PRELIEVO:
    MsgBox Err.Description, vbCritical, "IMPOSTA_DATI_LISTA_PRELIEVO"
End Sub

Private Sub RIPRISTINA_GRIGLIA_CONDIZIONI()
Dim I As Long
Dim X As String

Me.BrwMain.Conditions.ClearValues
'Me.BrwMain.Conditions.Item("Link_Doc_sezionale").FromValue = GET_DESCRIZIONE_SEZIONALE(oDoc.IDSezionale)
'
Me.BrwMain.Conditions.Item("Doc_numero").FromValue = oDoc.Numero
Me.BrwMain.Conditions.Item("Doc_data").FromValue = oDoc.DataEmissione
Me.BrwMain.Conditions.Item("Doc_data").ToValue = oDoc.DataEmissione
'Me.BrwMain.Conditions.Item("IDOggetto").FromValue = oDoc.IDOggetto

BrwMain.ApplyFilter

If (BrwMain.Recordset.EOF = False) Then
    brwMain_DblClick
End If


End Sub


Private Function fncTrovaIDFunzione(Gestore As String, Optional Funzione As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione, Gestore.Gestore "
sSQL = sSQL & "FROM Gestore INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Gestore.IDGestore = TipoOggetto.IDGestore INNER JOIN "
sSQL = sSQL & "Funzione ON TipoOggetto.IDTipoOggetto = Funzione.IDTipoOggetto "
sSQL = sSQL & "WHERE Gestore.Gestore = " & fnNormString(Gestore)
sSQL = sSQL & " AND Funzione = " & fnNormString(Funzione)
sSQL = sSQL & " AND Funzione.IDFunzione >= 10000"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    fncTrovaIDFunzione = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzione = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_CLIENTE_DESTINAZIONE(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POConfigurazioneCliente, IDAnagrafica, IDAnagraficaDestinazione "
sSQL = sSQL & "FROM RV_POConfigurazioneCliente "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaCliente

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CLIENTE_DESTINAZIONE = 0
Else
    GET_LINK_CLIENTE_DESTINAZIONE = fnNotNullN(rs!IDAnagraficaDestinazione)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub ParametroPesoArticolo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPesoArticolo, SelezionaLottoCampagnaInLavorazione "
sSQL = sSQL & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & " (IDFiliale=" & TheApp.Branch & ") "
sSQL = sSQL & " AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    TIPO_PESO_ARTICOLO = fnNotNullN(rs!IDRV_POTipoPesoArticolo)
    ATTIVA_SEL_LOTTO_PROD_IN_LAV = fnNotNullN(rs!SelezionaLottoCampagnaInLavorazione)
Else
    TIPO_PESO_ARTICOLO = 0
    ATTIVA_SEL_LOTTO_PROD_IN_LAV = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_DESCRIZIONE_SEZIONALE(IDSezionale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDSezionale, Sezionale "
sSQL = sSQL & "FROM Sezionale "
sSQL = sSQL & "WHERE IDSezionale=" & IDSezionale


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_DESCRIZIONE_SEZIONALE = fnNotNull(rs!Sezionale)
Else
    GET_DESCRIZIONE_SEZIONALE = ""
End If

rs.CloseResultset
Set rs = Nothing
End Function
'**+
'Nome: OnPrint
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Print
'**/
'**+
'Nome: OnPrint
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Operazioni sul comando Print
'**/
Private Sub OnPrint(ByVal ToolName As String)
Dim lFlags As Long
Dim OLDCursor As Integer
Dim sStr As String
Dim Field As DmtDocManLib.Field
Dim IDReportDefault As Long
    
        
    
    OLDCursor = Screen.MousePointer
    
    IDReportDefault = oReportsActivity.DefaultReportID
    
    'Se il filtro attivo è "Nessun record" è possibile eseguire una stampa/esportazione soltanto se
    'si è in modalità form. In tal caso, infatti, verrà passato al Crystals Reports un filtro
    'creato ad hoc sull'ID del record attuale.
    If m_ActiveFilter.NothingSelected And BrwMain.Visible Then
        sStr = "Impossibile effettuare l'operazione richiesta." & vbCrLf
        sStr = sStr & "Prima di procedere occorre eseguire un filtro."
        sbMsgInfo sStr, m_App.FunctionName
        Screen.MousePointer = OLDCursor
        Exit Sub
    End If
    
    'Se non esiste un report attivo occorre annullare l'operazione.
   'Se non esiste un report attivo occorre annullare l'operazione.
    If Len(oReportsActivity.SelectedReportName) > 0 Then
        Set m_Report = m_DocType.Reports.Item(oReportsActivity.SelectedReportName)
    End If
    If m_Report Is Nothing Then
        sbMsgError "Impossibile eseguire - Nessun report predefinito.", m_App.FunctionName
        GoTo OnPrint_Exit
    End If
    m_iNumeroCopieDefault = m_Report.Copies
    m_OrientamentoDefault = m_Report.Orientation
    
    
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    'Se è attivo il pulsante Salva deve essere visualizzato un messaggio di avviso
    'con i pulsanti OK e annulla (occorre salvare PRIMA della stampa)
    If m_Changed Then
        Select Case ChooseAboutSavingOkCancel
            Case vbOK
                OnSave
                'Se la registrazione non è andata a buon fine esce
                If Not m_Saved Then
                    GoTo OnPrint_Exit
                End If
                
            Case vbCancel
                GoTo OnPrint_Exit
        End Select
    End If
    
    
        Set oReport = New dmtReportLib.dmtReport
        Set oReport.Connection = Cn
        oReport.Password = m_App.Password
        oReport.User = m_App.User
        oReport.Copies = m_Report.Copies
        oReport.Orientation = m_Report.Orientation
        
        
            
        'Viene inserita la condizione di ricerca basata sull'ID del record corrente.
        fnDeleteTabellaRicorsione m_App.IDUser, oDoc.IDTipoOggetto

        'm_DocType.Fields("IDOggetto").Value = m_Document.Fields("IDOggetto").Value
        ''m_DocType.Fields("IDOggetto").Value = oDoc.IDOggetto
        'Viene creato un filtro temporaneo per il Crystals Reports.
        ''m_DocType.RemoveFilter "Form"
            'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = oDoc.IDFiliale 'IDFiliale
            'Imposta l'identificativo del tipo di documento
            oReport.DocTypeID = fncIDTipoOggettoPrg(App.EXEName)
            
    
    If Not BrwMain.Visible Then
        'Modalità Form - deve stampare solo il record corrente
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        'For Each Field In m_DocType.Fields
        '    Field.Value = Empty
        'Next
            'oReport.Where = "IDOggetto = 873" '& Val(Me.Txt_Reg_IDRegistro)
            oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, oDoc.IDOggetto, oDoc.IDTipoOggetto
            oReport.Where = "IDOggetto = " & oDoc.IDOggetto
            oReport.Where = oReport.Where & " AND IDUtente = " & oDoc.IDUtente
            
        ''Set m_Report.Filter = m_DocType.AddFilterWithConditions("Form")
            
    Else
        'Modalità vista tabellare
            GET_OGGETTI_PER_STAMPE
            'oDoc.Prepare2Print oDoc.IDAzienda, oDoc.IDUtente, BrwMain.AllColumns("IDOggetto").Value, oDoc.IDTipoOggetto
            'oReport.Where = "IDOggetto = " & BrwMain.AllColumns("IDOggetto").Value
            oReport.Where = "IDUtente = " & oDoc.IDUtente
                
    End If

    fncImpostaDefaultReport oReportsActivity.SelectedReportID, fnGetTipoOggetto("RV_POOrdineL")
    Me.ActivityBox.Redraw = True
    DoEvents
    Select Case ToolName
    
        Case "PrePrint", "Mnu_PrePrint"
            On Error GoTo ErrorHandler
            
            Screen.MousePointer = vbHourglass
            
            
            
            'm_TabMode = BrwMain.Visible
            'PicForm.Visible = False
            'BrwMain.Visible = False
            'DocTypeExplorer.Visible = False
            
            'SetStatus4Modality Preview, OpenPrw
            'Refresh
            'm_Document.Preview m_Report, m_DocType.PrinterName, 0, 0, 0, CInt(ScaleWidth / Screen.TwipsPerPixelX), CInt(ScaleHeight / Screen.TwipsPerPixelY), False
            
            oReport.Preview 0, 0, 0
            
            'm_Document.Preview m_Report, "", 0, 0, 0, CInt(ScaleWidth / Screen.TwipsPerPixelX), CInt(ScaleHeight / Screen.TwipsPerPixelY), True
            'lFlags = SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOMOVE
            SetWindowPos m_PreviewWindowHandle, HWND_TOP, 0, 0, 0, 0, lFlags
            
        Case "Print", "Mnu_Print"
            'PrintDocument ToolName
            oReport.DoPrint oReport.PrinterName
            
        Case "ExportWord", "Mnu_ExportWord"
            oReport.Export recWord
            
            'ExportDocument ecWord
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Word, TheApp.Name

        Case "ExportExcel", "Mnu_ExportExcel"
            oReport.Export recExcel
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Excel, TheApp.Name
            
        Case "ExportHtml", "Mnu_ExportHtml"
            oReport.Export recHtml
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), HTML, TheApp.Name
        
        Case "ExportPDF", "Mnu_ExportPDF"
            oReport.Export recPDF
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), PDF, TheApp.Name
        
        Case "MailWord"
            oReport.SendMail recWord
            
        Case "MailExcel"
            oReport.SendMail recExcel
            
        Case "MailHtml"
            oReport.SendMail recHtml
        
        Case "MailPDF"
            oReport.SendMail recPDF
        
    End Select
    
    fncImpostaDefaultReport IDReportDefault, fnGetTipoOggetto("RV_POOrdineL")
    Me.ActivityBox.Redraw = True
    DoEvents
    
   
OnPrint_Exit:
    Set Field = Nothing
    Screen.MousePointer = OLDCursor
    Exit Sub
    
ErrorHandler:
    Const ERROR_PRINTING_ABORTED = 3
    Const ERROR_PRINTING_CANCELLED = 4
    Select Case Err.Number
        Case 20507
            'Errore "Invalid file Name" generato quando non è possibile trovare il file .rpt
            sbMsgInfo "File di report non trovato", m_App.FunctionName
        Case ERROR_PRINTING_ABORTED, ERROR_PRINTING_CANCELLED
            'non deve far niente, è stato già segnalato da CrystalReport
        Case Else
            If Len(Trim(Err.Description)) > 0 Then
                sbMsgInfo Err.Description, m_App.FunctionName
            End If
    End Select

    'Si è verificato un errore durante la procedura di anteprima.
    Screen.MousePointer = OLDCursor
    
    'Ripristina la situazione del form
    m_PreviewWindowHandle = 0
    PicForm.Visible = True

    BrwMain.Visible = m_TabMode
    ActivityBox.Visible = BarMenu.Bands("Band_View").Tools("Mnu_Folders").Checked
    FormRecalcLayout
    SetStatus4Modality Preview, ClosePrw
        
    Set Field = Nothing
End Sub

Private Sub GET_OGGETTI_PER_STAMPE()
On Error GoTo ERR_GET_OGGETTI_PER_STAMPE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL_WHERE As String
Dim Field As DmtDocManLib.Field
Dim Cond As DmtGridCtl.dgCondition
Dim OLD_CURSOR As Long


sSQL_WHERE = "WHERE IDAzienda=" & TheApp.IDFirm
sSQL_WHERE = sSQL_WHERE & " AND IDFiliale=" & TheApp.Branch
For Each Cond In BrwMain.Conditions

    Select Case Cond.ConditionType
        
        'Condizione boolean
        Case dgCondTypeBoolean
            If (Cond.FromValue = "SI") Then
                sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormBoolean(1)
            End If
            If (Cond.FromValue = "NO") Then
                sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormBoolean(0)
            End If
        'Condizione associata ad una combo box
        Case dgCondTypeComboDB
            If fnNotNullN(Cond.FromValueID) > 0 Then
                sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormString(Cond.FromValueID)
            End If 'm_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValueID
        
        'Condizione di tipo text, numeric, data, time
        Case dgCondTypeText
                If Len(Trim(fnNotNull(Cond.FromValue))) > 0 Then
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & " LIKE " & fnNormString(Cond.FromValue & "%")
                End If
        Case dgCondTypeNumber
            If Cond.RangeChecked = True Then
                If fnNotNullN(Cond.ToValue) = 0 Then
                    If fnNotNullN(Cond.FromValue) > 0 Then
                        sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormNumber(Cond.FromValue)
                    End If
                Else
                    If fnNotNullN(Cond.FromValue) > 0 Then
                        sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & ">=" & fnNormNumber(Cond.FromValue)
                    End If
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "<=" & fnNormNumber(Cond.ToValue)
                End If

            Else
                If fnNotNullN(Cond.FromValue) > 0 Then
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormNumber(Cond.FromValue)
                End If
            End If
        
        Case dgCondTypeDate
            If Cond.RangeChecked = True Then
                If Len(fnNotNull(Cond.ToValue)) = 0 Then
                    If Len(fnNotNull(Cond.FromValue)) > 0 Then
                        sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormDate(Cond.FromValue)
                    End If
                Else
                    If Len(fnNotNull(Cond.FromValue)) > 0 Then
                        sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & ">=" & fnNormDate(Cond.FromValue)
                    End If
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "<=" & fnNormDate(Cond.ToValue)
                End If
            Else
                If Len(fnNotNull(Cond.FromValue)) > 0 Then
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormDate(Cond.FromValue)
                End If
            End If
        
      
        'Altre condizioni
        Case Else
            sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & " LIKE " & fnNormString(Cond.FromValue)

    End Select
    
Next Cond

'sSQL = "SELECT Oggetto.IDOggetto "
'sSQL = sSQL & "FROM " & sTabellaTestata & " INNER JOIN "
'sSQL = sSQL & "Oggetto ON " & sTabellaTestata & ".IDOggetto = Oggetto.IDOggetto AND " & sTabellaTestata & ".IDTipoOggetto = Oggetto.IDTipoOggetto "
'sSQL = sSQL & sSQL_WHERE

sSQL = "SELECT * FROM RV_POIEOrdineCliente "
sSQL = sSQL & sSQL_WHERE

Set rs = Cn.OpenResultset(sSQL)

OLD_CURSOR = Cn.CursorLocation
Cn.CursorLocation = 1
 
While Not rs.EOF
    Screen.MousePointer = 11
    Me.Caption = "PREPARAZIONE STAMPA PER IL DOCUMENTO " & fnNotNullN(rs!IDOggetto)
    DoEvents
    
    oDoc.Prepare2Print TheApp.IDFirm, TheApp.IDUser, fnNotNullN(rs!IDOggetto), oDoc.IDTipoOggetto
    
    DoEvents
    Screen.MousePointer = 0
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Me.Caption = "PREPARAZIONE STAMPA CONCLUSA................"

Cn.CursorLocation = OLD_CURSOR

Me.Caption = Caption2Display(True)

Exit Sub
ERR_GET_OGGETTI_PER_STAMPE:
    MsgBox Err.Description, vbCritical, "GET_OGGETTI_PER_STAMPE"
End Sub

Private Sub GET_TABELLE_ARTICOLO()
On Error GoTo ERR_GET_TABELLE_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


Set rsTipoLav = New ADODB.Recordset
rsTipoLav.CursorLocation = adUseClient

rsTipoLav.Fields.Append "ID", adInteger, , adFldIsNullable
rsTipoLav.Fields.Append "Descrizione", adVarChar, 250, adFldIsNullable
rsTipoLav.Fields.Append "Tipo", adInteger, , adFldIsNullable


rsTipoLav.Open , , adOpenKeyset, adLockBatchOptimistic


sSQL = "SELECT * FROM RV_POCalibro"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rsTipoLav.AddNew
        rsTipoLav!ID = fnNotNullN(rs!IDRV_POCalibro)
        rsTipoLav!Descrizione = fnNotNull(rs!Calibro)
        rsTipoLav!TIpo = 1
    rsTipoLav.Update
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing


sSQL = "SELECT * FROM RV_POTipoCategoria"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rsTipoLav.AddNew
        rsTipoLav!ID = fnNotNullN(rs!IDRV_POTipoCategoria)
        rsTipoLav!Descrizione = fnNotNull(rs!TipoCategoria)
        rsTipoLav!TIpo = 2
    rsTipoLav.Update
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing


sSQL = "SELECT * FROM RV_POTipoLavorazione"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    rsTipoLav.AddNew
        rsTipoLav!ID = fnNotNullN(rs!IDRV_POTipoLavorazione)
        rsTipoLav!Descrizione = fnNotNull(rs!TipoLavorazione)
        rsTipoLav!TIpo = 3
    rsTipoLav.Update
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

Exit Sub

ERR_GET_TABELLE_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_TABELLE_ARTICOLO"
End Sub


Private Function GET_DESCRIZIONE_TABELLA(TIpo As Long, IDTabella As Long) As String
On Error GoTo ERR_GET_DESCRIZIONE_TABELLA
GET_DESCRIZIONE_TABELLA = ""
rsTipoLav.Filter = "Tipo=" & TIpo & " AND ID=" & IDTabella


If Not rsTipoLav.EOF Then
    GET_DESCRIZIONE_TABELLA = fnNotNull(rsTipoLav!Descrizione)
End If

rsTipoLav.Filter = vbNullString
Exit Function
ERR_GET_DESCRIZIONE_TABELLA:
    
End Function
Private Function GET_TOTALI_COLLI_ORDINE(IDOggettoOrdine As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(art_numero_colli) as NumeroPedane "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND RV_POTipoRiga=1 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALI_COLLI_ORDINE = 0
Else
    GET_TOTALI_COLLI_ORDINE = fnNotNullN(rs!NumeroPedane)
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_NUMERO_COLLI_LAVORATE_PER_ORDINE(IDOggettoOrdine As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_COLLI_LAVORATE_PER_ORDINE = 0

sSQL = "SELECT SUM(Colli) as NumeroColli "
sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
sSQL = sSQL & " GROUP BY IDRV_POPedana"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_COLLI_LAVORATE_PER_ORDINE = 0
Else
    GET_NUMERO_COLLI_LAVORATE_PER_ORDINE = fnNotNullN(rs!NumeroColli)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_TOTALI_ORDINE_DETTAGLIO_COLLI(IDOggettoOrdine As Long, IDArticolo As Long, raggrOrd As String) As Double
On Error GoTo ERR_GET_TOTALI_ORDINE_DETTAGLIO_COLLI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TOTALI_ORDINE_DETTAGLIO_COLLI = 0

sSQL = "SELECT SUM(art_numero_colli) as NumeroPedane "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoOrdine
sSQL = sSQL & " AND Link_art_articolo=" & IDArticolo
sSQL = sSQL & " AND RV_PONotaRigaOrdRaggr=" & fnNormString(raggrOrd)
sSQL = sSQL & " AND RV_POTipoRiga=1 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALI_ORDINE_DETTAGLIO_COLLI = 0
Else
    GET_TOTALI_ORDINE_DETTAGLIO_COLLI = fnNotNullN(rs!NumeroPedane)
End If

Exit Function
ERR_GET_TOTALI_ORDINE_DETTAGLIO_COLLI:
    MsgBox Err.Description, vbCritical, "GET_TOTALI_ORDINE_DETTAGLIO_COLLI"
rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO(IDOggettoOrdine As Long, IDArticoloPadre As Long, raggrOrd As String) As Double
Dim sSQL As String
Dim rsArt As DmtOleDbLib.adoResultset
Dim rs As DmtOleDbLib.adoResultset
Dim rstmp As ADODB.Recordset


GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO = 0


    sSQL = "SELECT SUM(Colli) as NumeroColli "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDOggettoOrdinePadre=" & IDOggettoOrdine
    sSQL = sSQL & " AND NotaRigaOrdRaggr=" & fnNormString(raggrOrd)
    sSQL = sSQL & " AND IDArticolo=" & IDArticoloPadre
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO = 0
    Else
        GET_NUMERO_COLLI_LAVORATI_PER_ARTICOLO_ORDINATO = fnNotNullN(rs!NumeroColli)
    End If
    
   
    
    rs.CloseResultset
    Set rs = Nothing
    
End Function
Private Sub AGGIORNA_NUMERAZIONE_ORDINE()
On Error GoTo ERR_AGGIORNA_NUMERAZIONE_ORDINE
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoPerTipo000F SET "
sSQL = sSQL & "RV_PONumeroOrdinePadre=Doc_numero, "
sSQL = sSQL & "RV_PODataOrdinePadre=Doc_data, "
sSQL = sSQL & "RV_PONumeroListaPrelievo=1, "
sSQL = sSQL & "RV_POIDOrdinePadre=IDOggetto, "
sSQL = sSQL & "RV_PONumeroPedanePrelievo=0, "
sSQL = sSQL & "RV_POOrdineCompletato=0, "
sSQL = sSQL & "RV_POIDAnagraficaDestinazione=0 "
sSQL = sSQL & "WHERE (RV_POIDOrdinePadre IS NULL)"

Cn.Execute sSQL

Exit Sub
ERR_AGGIORNA_NUMERAZIONE_ORDINE:
    MsgBox Err.Description, vbCritical, "AGGIORNA_NUMERAZIONE_ORDINE"
End Sub

Private Sub GET_STAMPA_DOCUMENTI_SEL()
On Error GoTo ERR_GET_STAMPA_DOCUMENTI_SEL
Dim Cond As DmtGridCtl.dgCondition
Dim sSQL As String
Dim IDReportDefault As Long
Dim NomeVistaOggetto As String

NomeVistaOggetto = GET_VISTA_TIPO_OGGETTO(m_DocType.ID)

If NomeVistaOggetto = "" Then
    MsgBox "Impossibile recuperare la vista dell'oggetto", vbCritical, "Recupero dati"
    Exit Sub
End If

sSQL = "SELECT * FROM " & NomeVistaOggetto
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch



For Each Cond In BrwMain.Conditions
    If Cond.IsHeader = False Then
        Select Case Cond.ConditionType
              
          'Condizione boolean
            Case dgCondTypeBoolean
                'm_DocType.Fields(Cond.FieldName).Value = IIf(IsEmpty(Cond.FromValue), Empty, Abs(CDbl(Cond.FromValue = "SI")))
                If (Cond.FromValue = "SI") Then
                    sSQL = sSQL & " AND " & Cond.FieldName & "=" & fnNormBoolean(1)
                End If
                If (Cond.FromValue = "NO") Then
                    sSQL = sSQL & " AND " & Cond.FieldName & "=" & fnNormBoolean(0)
                End If
          'Condizione associata ad una combo box
            Case dgCondTypeComboDB
                'm_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValueID
                If fnNotNullN(Cond.FromValueID) > 0 Then
                    sSQL = sSQL & " AND " & Cond.FieldName & "=" & Cond.FromValueID
                End If
          'Condizione di tipo text, numeric, data, time
            Case dgCondTypeText
                If Cond.RangeChecked = True Then
                    If Len(Cond.FromValue) > 0 Then
                        sSQL = sSQL & " AND " & Cond.FieldName & ">=" & fnNormString(Cond.FromValue)
                    End If
                    If Len(Cond.ToValue) > 0 Then
                        sSQL = sSQL & " AND " & Cond.FieldName & "<=" & fnNormString(Cond.ToValue)
                    End If

                    'm_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    'm_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                    If Len(Cond.FromValue) > 0 Then
                        sSQL = sSQL & " AND " & Cond.FieldName & " LIKE " & fnNormString("%" & Replace(Cond.FromValue, "%", "") & "%")
                    End If
                End If
            Case dgCondTypeNumber
              If Cond.RangeChecked = True Then
                    If Cond.FromValue > 0 Then
                        sSQL = sSQL & " AND " & Cond.FieldName & ">=" & fnNormNumber(Cond.FromValue)
                    End If
                    If Cond.ToValue > 0 Then
                        sSQL = sSQL & " AND " & Cond.FieldName & "<=" & fnNormNumber(Cond.ToValue)
                    End If
              Else
                If fnNotNullN(Cond.FromValue) > 0 Then
                    sSQL = sSQL & " AND " & Cond.FieldName & "=" & fnNormNumber(Cond.FromValue)
                End If
              End If
          
            Case dgCondTypeDate
              If Cond.RangeChecked = True Then
                  'm_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                    If Cond.FromValue > 0 Then
                        sSQL = sSQL & " AND " & Cond.FieldName & ">=" & fnNormDate(Cond.FromValue)
                    End If
                    If Cond.ToValue > 0 Then
                        sSQL = sSQL & " AND " & Cond.FieldName & "<=" & fnNormDate(Cond.ToValue)
                    End If

              Else
                If Cond.FromValue > 0 Then
                    sSQL = sSQL & " AND " & Cond.FieldName & "=" & fnNormDate(Cond.FromValue)
                End If
              End If

          'Altre condizioni
            Case Else
                'm_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                  
    
        End Select
    End If
Next Cond

sSQL = sSQL & " ORDER BY Doc_data DESC, Doc_Numero DESC"
Set rsGrigliaSelDoc = New ADODB.Recordset
rsGrigliaSelDoc.Open sSQL, Cn.InternalConnection

If Len(oReportsActivity.SelectedReportName) > 0 Then
    Set m_Report = m_DocType.Reports.Item(oReportsActivity.SelectedReportName)
End If

NUMERO_COPIE_SEL_DOC = m_Report.Copies
ORIENTAMENTO_SEL_DOC = m_Report.Orientation

IDReportDefault = oReportsActivity.DefaultReportID

fncImpostaDefaultReport oReportsActivity.SelectedReportID, fnGetTipoOggetto(App.EXEName)

frmSelDocStampa.Show vbModal


fncImpostaDefaultReport IDReportDefault, fnGetTipoOggetto(App.EXEName)
Me.ActivityBox.Redraw = True
DoEvents


Exit Sub
ERR_GET_STAMPA_DOCUMENTI_SEL:
    MsgBox Err.Description, vbCritical, "GET_STAMPA_DOCUMENTI_SEL"

End Sub
Private Function GET_VISTA_TIPO_OGGETTO(IDTipoOggetto As Long) As String
On Error GoTo ERR_GET_VISTA_TIPO_OGGETTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_VISTA_TIPO_OGGETTO = ""

sSQL = "SELECT * FROM ViewPerTipoOggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_VISTA_TIPO_OGGETTO = ""
Else
    GET_VISTA_TIPO_OGGETTO = fnNotNull(rs!Tabella)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_VISTA_TIPO_OGGETTO:

End Function
Private Sub ELIMINA_RIGHE_DOCUMENTO_VUOTE(IDOggetto As Long, Tabella As String)
On Error GoTo ERR_ELIMINA_RIGHE_DOCUMENTO_VUOTE
Dim sSQL As String

sSQL = "DELETE FROM " & Tabella
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POLinkRiga IS NULL"
sSQL = sSQL & " AND RV_POTipoRiga IS NULL"

Cn.Execute sSQL

Exit Sub
ERR_ELIMINA_RIGHE_DOCUMENTO_VUOTE:
    MsgBox Err.Description, vbCritical, "ELIMINA_RIGHE_DOCUMENTO_VUOTE"



End Sub
Private Sub GET_IMBALLO_PER_ARTICOLO_CONF(IDArticoloMerce As Long)
On Error GoTo ERR_GET_IMBALLO_PER_ARTICOLO_CONF
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long

sSQL = "SELECT IDArticoloImballo, COUNT(IDRV_PODistintaBaseRigheConf) AS Numero "
sSQL = sSQL & "FROM RV_POIEDistintaBasaConf "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloMerce
sSQL = sSQL & " GROUP BY IDArticoloImballo "

Set rs = Cn.OpenResultset(sSQL)

'If rs.EOF Then
'    NumeroRecord = 0
'Else
'    NumeroRecord = fnNotNullN(rs!Numero)
'End If

NumeroRecord = 0

While Not rs.EOF
    NumeroRecord = NumeroRecord + 1
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If NumeroRecord = 0 Then Exit Sub

If (NumeroRecord = 1) Then
    sSQL = "SELECT IDArticoloImballo  "
    sSQL = sSQL & "FROM RV_POIEDistintaBasaConf "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloMerce
    sSQL = sSQL & " GROUP BY IDArticoloImballo "
    
    Set rs = Cn.OpenResultset(sSQL)

    If rs.EOF Then
        Me.CDImballo.Load 0
    Else
        Me.CDImballo.Load fnNotNullN(rs!IDArticoloImballo)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Else
    frmImballoPerArt.Show vbModal
End If

Exit Sub
ERR_GET_IMBALLO_PER_ARTICOLO_CONF:
    MsgBox Err.Description, vbCritical, "GET_IMBALLO_PER_ARTICOLO_CONF"
End Sub
Private Sub GET_CONFEZIONI_IMBALLO_PER_ARTICOLO(IDArticoloMerce As Long, IDArticoloImballo As Long)
On Error GoTo ERR_GET_CONFEZIONI_IMBALLO_PER_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long

Me.txtNumeroConfImballo.Value = 0
Me.txtTaraConfImballo.Value = 0

sSQL = "SELECT COUNT(IDRV_PODistintaBaseRigheConf) AS Numero "
sSQL = sSQL & "FROM RV_POIEDistintaBasaConf "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloMerce
sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!Numero)
End If

rs.CloseResultset
Set rs = Nothing

If NumeroRecord = 0 Then Exit Sub

If (NumeroRecord = 1) Then
    sSQL = "SELECT *  "
    sSQL = sSQL & "FROM RV_POIEDistintaBasaConf "
    sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloMerce
    sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    
    Set rs = Cn.OpenResultset(sSQL)

    If rs.EOF Then
        Me.txtNumeroConfImballo.Value = 0
        Me.txtTaraConfImballo.Value = 0
        Me.CDImballoPrimario.Load 0
    Else
        Me.txtNumeroConfImballo.Value = fnNotNullN(rs!NumeroConfezioni)
        Me.txtTaraConfImballo.Value = fnNotNullN(rs!Tara)
        Me.CDImballoPrimario.Load fnNotNullN(rs!IDArticoloAssociato)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Else
    frmNumeroConf.Show vbModal
End If

Exit Sub
ERR_GET_CONFEZIONI_IMBALLO_PER_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_CONFEZIONI_IMBALLO_PER_ARTICOLO"
End Sub
Private Sub GET_PROP_CONFEZIONI(IDArticoloMerce As Long, IDArticoloImballo As Long, IDArticoloConfezione As Long)
On Error GoTo ERR_GET_CONFEZIONI_IMBALLO_PER_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NumeroRecord As Long

GET_INFO_ARTICOLO IDArticoloMerce

Me.txtNumeroConfImballo.Value = 0
Me.txtTaraConfImballo.Value = 0

sSQL = "SELECT *  "
sSQL = sSQL & "FROM RV_POIEDistintaBasaConf "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticoloMerce
sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
sSQL = sSQL & " AND IDArticoloAssociato=" & IDArticoloConfezione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtNumeroConfImballo.Value = 0
    Me.txtTaraConfImballo.Value = 0
Else
    Me.txtNumeroConfImballo.Value = fnNotNullN(rs!NumeroConfezioni)
    Me.txtTaraConfImballo.Value = fnNotNullN(rs!Tara)
    
    If (fnNotNullN(rs!QuantitaPerConfezione) * fnNotNullN(rs!NumeroConfezioni)) > 0 Then
        Me.txtQuantitaPerCollo.Value = fnNotNullN(rs!QuantitaPerConfezione) * fnNotNullN(rs!NumeroConfezioni)
    End If
    If fnNotNullN(rs!PesoPerCollo) > 0 Then
        Me.txtPesoPerCollo.Value = fnNotNullN(rs!PesoPerCollo)
    End If
    If fnNotNullN(rs!Moltiplicatore) > 0 Then
        Me.txtMoltiplicatore.Value = fnNotNullN(rs!Moltiplicatore)
    End If
    
'    If fnNotNullN(rs!PesoPerCollo) > 0 Then
'        PESO_LORDO = fnNotNullN(rs!PesoPerCollo)
'    End If
'    If fnNotNullN(rs!QuantitaPerConfezione) > 0 Then
'        QUANTITA_PER_COLLO = fnNotNullN(rs!QuantitaPerConfezione) * fnNotNullN(rs!NumeroConfezioni)
'    End If
End If

rs.CloseResultset
Set rs = Nothing
    
txtColli_LostFocus

Exit Sub
ERR_GET_CONFEZIONI_IMBALLO_PER_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_PROP_CONFEZIONI"
End Sub
Private Sub GET_INFO_ARTICOLO(IDArticolo As Long)
On Error GoTo ERR_GET_INFO_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'PESO_LORDO = 0
'QUANTITA_PER_COLLO = 0

Me.txtPesoPerCollo.Value = 0
Me.txtQuantitaPerCollo.Value = 0
Me.txtMoltiplicatore.Value = 0

sSQL = "SELECT RV_POQuantitaPerCollo, PesoNetto, RV_POMoltiplicatore "
sSQL = sSQL & "FROM Articolo WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtPesoPerCollo.Value = fnNotNullN(rs!PesoNetto)
    Me.txtQuantitaPerCollo.Value = fnNotNullN(rs!RV_POQuantitaPerCollo)
    Me.txtMoltiplicatore.Value = fnNotNullN(rs!RV_POMoltiplicatore)
End If


rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_INFO_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_INFO_ARTICOLO"
End Sub
Private Function GET_QTA_PER_COLLO_CONFEZ(IDArticoloMerce As Long, IDArticoloImballo As Long, IDArticoloConfezione As Long) As Double
On Error GoTo ERR_GET_CONFEZIONI_IMBALLO_PER_ARTICOLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_QTA_PER_COLLO_CONFEZ = 1

sSQL = "SELECT QuantitaPerConfezione  "
sSQL = sSQL & " FROM RV_POIEDistintaBasaConf "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticoloMerce
sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
sSQL = sSQL & " AND IDArticoloAssociato=" & IDArticoloConfezione

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_QTA_PER_COLLO_CONFEZ = fnNotNullN(rs!QuantitaPerConfezione)
End If

rs.CloseResultset
Set rs = Nothing

If GET_QTA_PER_COLLO_CONFEZ = 0 Then GET_QTA_PER_COLLO_CONFEZ = 1

Exit Function
ERR_GET_CONFEZIONI_IMBALLO_PER_ARTICOLO:
    MsgBox Err.Description, vbCritical, "GET_QTA_PER_COLLO_CONFEZ"
End Function

Private Sub GET_CONTRATTO()
On Error GoTo ERR_GET_CONTRATTO

LINK_CONTRATTO = Me.txtIDContratto.Value
LINK_CLIENTE_CONTRATTO = Me.ACSCliente.IDAnagrafica

If LINK_CONTRATTO = 0 Then
    frmContratto.Show vbModal
    
    If LINK_CONTRATTO > 0 Then
        GET_DATI_DA_CONTRATTO LINK_CONTRATTO
    End If
End If
If LINK_CONTRATTO > 0 Then
    SSTab1.Tab = 1
    frmContrattoDettaglio.Show vbModal
    If Not ((rsContrattoDettaglioSel.EOF) And (rsContrattoDettaglioSel.BOF)) Then
        rsContrattoDettaglioSel.MoveFirst
        DATI_DA_CONTRATTO = True
        While Not rsContrattoDettaglioSel.EOF
            Screen.MousePointer = 11
            DoEvents
            cmdNuovo_Click
                GET_RIGA_DA_CONTRATTO rsContrattoDettaglioSel
            cmdSalva_Click
            Screen.MousePointer = 0
            DoEvents
        rsContrattoDettaglioSel.MoveNext
        Wend
        
        rsContrattoDettaglioSel.Close
        Set rsContrattoDettaglioSel = Nothing
        
        DATI_DA_CONTRATTO = False
    End If
End If
Exit Sub
ERR_GET_CONTRATTO:
    Screen.MousePointer = 0
    DATI_DA_CONTRATTO = False
    MsgBox Err.Description, vbCritical, "GET_CONTRATTO"
End Sub

Private Sub GET_DATI_DA_CONTRATTO(IDContratto As Long)
On Error GoTo ERR_DATI_DA_CONTRATTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM ValoriOggettoPerTipo000E "
sSQL = sSQL & "WHERE IDOggetto=" & IDContratto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.cdAnagrafica.Load fnNotNullN(rs!Link_Nom_anagrafica)
    Me.cboAltroSito.WriteOn fnNotNullN(rs!Link_Nom_ult_sito)
    Me.cboLuogoPresaMerce.WriteOn fnNotNullN(rs!RV_POIDLuogoPresaMerce)
    Me.cboVettore.WriteOn fnNotNullN(rs!Link_Vet_vettore)
    Me.txtTargaAutomezzo.Text = fnNotNull(rs!RV_POTargaAutomezzo)
    Me.txtIstruzioniMittente.Text = fnNotNull(rs!RV_POIstruzioniMittente)
    Me.ACSAnaDest.sbLoadCFByIDAnagrafica 0, fnNotNullN(rs!RV_POIDAnagraficaDestinazione)
    
    Me.cboPorto.WriteOn fnNotNullN(rs!Link_Nom_porto)
    Me.cboTrasporto.WriteOn fnNotNullN(rs!Link_Doc_spedizione)
    Me.CDAgenteTesta.Load fnNotNullN(rs!Link_doc_agente)
    Me.cboAspettoEsteriore.WriteOn fnNotNullN(rs!Link_Doc_aspetto_esteriore)
    Me.txtAnnotazioni.Text = fnNotNull(rs!Doc_annotazioni_variazio)
    Me.txtDescrizioneRigaDoc.Text = fnNotNull(rs!RV_PODescrizioneCorpoDocEv)
    Me.txtAnnotazioniInterna.Text = fnNotNull(rs!RV_POAnnotazioniInterna)
    Me.cboVettoreSuccessivo.WriteOn fnNotNullN(rs!RV_POIDTrasportatoreSuccessivo)
    
    If fnNotNullN(rs!Link_doc_pagamento) > 0 Then
        Me.cboPagamento.WriteOn fnNotNullN(rs!Link_doc_pagamento)
    End If
    If fnNotNullN(rs!Link_Nom_contratto_bancario) > 0 Then
        Me.cboBancaCliente.WriteOn fnNotNullN(rs!Link_Nom_contratto_bancario)
    End If
    
    Me.cboTipoOrdine.WriteOn fnNotNullN(rs!RV_POIDTipoOrdine)
    Me.txtCausaleDocumento.Text = fnNotNull(rs!Doc_causale_trasporto)
    
    If fnNotNullN(rs!Link_Doc_contratto_bancario_az) > 0 Then
        Me.cboBancaAzienda.WriteOn fnNotNullN(rs!Link_Doc_contratto_bancario_az)
    End If
    If fnNotNullN(rs!Link_Doc_listino) > 0 Then
        Me.cboListino.WriteOn fnNotNullN(rs!Link_Doc_listino)
    End If
    If fnNotNullN(rs!Link_Doc_listino_base) > 0 Then
        Me.cboListinoAzienda.WriteOn fnNotNullN(rs!Link_Doc_listino_base)
    End If
    If fnNotNullN(rs!Link_Nom_accordi_commerciali) > 0 Then
        Me.cboAccordoCommerciale.WriteOn fnNotNullN(rs!Link_Nom_accordi_commerciali)
    End If
    If fnNotNullN(rs!Link_Nom_raggrup_fatturato) > 0 Then
        Me.cboRaggrFatturato.WriteOn fnNotNullN(rs!Link_Nom_raggrup_fatturato)
    End If
    If (RIP_LET_INT_DA_DOC_COLL = 1) Then
        Me.txtIDLetteraIntento.Value = fnNotNullN(rs!Link_Nom_lettera_intento)
        Me.cboIvaCliente.WriteOn fnNotNullN(rs!Link_Nom_IVA)
    End If
    Me.txtIDContratto.Value = LINK_CONTRATTO
    
End If
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_DATI_DA_CONTRATTO:
    MsgBox Err.Description, vbCritical, "DATI_DA_CONTRATTO"
End Sub
Private Sub GET_RIGA_DA_CONTRATTO(rs As ADODB.Recordset)
On Error GoTo ERR_GET_RIGA_DA_CONTRATTO
Me.CDArticolo.Load fnNotNullN(rs!Link_Art_articolo)
Me.txtDescrizioneArticolo.Text = fnNotNull(rs!Art_descrizione)
Me.cboUnitaDiMisura.WriteOn fnNotNullN(rs!Link_Art_unita_di_misura)
Me.cboCalibro.WriteOn fnNotNullN(rs!RV_POIDCalibro)
Me.cboTipoLavorazione.WriteOn fnNotNullN(rs!RV_POIDTipoLavorazione)
Me.cboTipoCategoria.WriteOn fnNotNullN(rs!RV_POIDTipoCategoria)
Me.txtRaggrRigaOrdine.Text = fnNotNull(rs!RV_PONotaRigaOrdRaggr)

Me.CDTipoPedana.Load fnNotNullN(rs!RV_POIDTipoPedana)
Me.CDPedana.Load fnNotNullN(rs!RV_POIDArticoloPedana)
Me.txtQuantitaPedana.Value = fnNotNullN(rs!RV_POQuantitaPedana)
Me.txtColliSfusi.Value = fnNotNullN(rs!RV_POColliSfusi)

Me.CDImballo.Load fnNotNullN(rs!RV_POIDImballo)
Me.txtTaraUnitaria.Value = fnNotNullN(rs!RV_POTaraUnitariaImballo)
Me.txtQuantitaPerPedana.Value = fnNotNullN(rs!RV_POColliPerPedana)
Me.txtImportoUnitarioImballo.Value = fnNotNullN(rs!RV_POImportoUnitarioImballo)
Me.chkImportoImballoInArticolo.Value = Abs(fnNotNullN(rs!RV_POImportoImballoInArticolo))

Me.CDImballoPrimario.Load fnNotNullN(rs!RV_POIDImballoPrimario)
Me.txtTaraConfImballo.Value = fnNotNullN(rs!RV_POTaraImballoPrimario)
Me.txtNumeroConfImballo.Value = fnNotNullN(rs!RV_PONumeroConfezioniPerImballo)
Me.txtQuantitaPerCollo.Value = fnNotNullN(rs!RV_POQuantitaPerCollo)
Me.txtPesoPerCollo.Value = fnNotNullN(rs!RV_POPesoPerCollo)
Me.txtMoltiplicatore.Value = fnNotNullN(rs!RV_POMoltiplicatorePerCollo)

Me.txtColli.Value = fnNotNullN(rs!Art_numero_colli)
Me.txtPesoLordo.Value = fnNotNullN(rs!Art_peso)
Me.txtTara.Value = fnNotNullN(rs!Art_tara)
Me.txtPesoNetto.Value = fnNotNullN(rs!PesoNetto)
Me.txtPezzi.Value = fnNotNullN(rs!Art_quantita_pezzi)
Me.txtQta_UM.Value = fnNotNullN(rs!Art_quantita_totale)
Me.txtImportoUnitarioArticolo.Value = fnNotNullN(rs!Art_prezzo_unitario_neutro)
Me.txtSconto1.Value = fnNotNullN(rs!Art_sco_in_percentuale_1)
Me.txtSconto2.Value = fnNotNullN(rs!Art_sco_in_percentuale_2)
Me.txtIDRigaContratto.Value = fnNotNullN(rs!IDValoriOggettoDettaglio)

Me.cboReportNew.WriteOn fnNotNullN(rs!RV_POIDConfigurazioneEtichettaLavorazione)
Me.cboReportPedNew.WriteOn fnNotNullN(rs!RV_POIDConfigurazioneEtichettaPedana)

GET_DESCRIZIONI_RIGA_CONTR rs

If (RIP_IVA_DA_DOC_COLL = 1) Then
    GET_IVA_RIGA_ARTICOLO rs
    GET_IVA_RIGA_IMBALLO rs
End If


txtColli_LostFocus
CalcolaImportoScontato
CalcolaTotaleRiga
Exit Sub
ERR_GET_RIGA_DA_CONTRATTO:
    MsgBox Err.Description, vbCritical, "GET_RIGA_DA_CONTRATTO"
End Sub
Private Sub GET_DESCRIZIONI_RIGA_CONTR(rs As ADODB.Recordset)
On Error GoTo ERR_GET_DESCRIZIONI_RIGA_CONTR
Dim sSQL As String
Dim rsDati As DmtOleDbLib.adoResultset

sSQL = "SELECT IDValoriOggettoDettaglio, RV_POAnnotazioniRigaOrdine, RV_POAnnotazioniRigaLavorazione "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0038 "
sSQL = sSQL & " WHERE IDOggetto=" & rs!IDOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & rs!IDValoriOggettoDettaglio

Set rsDati = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    txtAnnotazioniDiRigaLav.Text = fnNotNull(rsDati!RV_POAnnotazioniRigaLavorazione)
    txtAnnotazioniDiRiga.Text = fnNotNull(rsDati!RV_POAnnotazioniRigaOrdine)
End If

rsDati.CloseResultset
Set rsDati = Nothing
Exit Sub
ERR_GET_DESCRIZIONI_RIGA_CONTR:
    MsgBox Err.Description, vbCritical, "GET_DESCRIZIONI_RIGA_CONTR"
End Sub
Private Sub GET_IVA_RIGA_ARTICOLO(rs As ADODB.Recordset)
On Error GoTo ERR_GET_DESCRIZIONI_RIGA_CONTR
Dim sSQL As String
Dim rsDati As DmtOleDbLib.adoResultset

sSQL = "SELECT IDValoriOggettoDettaglio, Art_aliquota_IVA, Link_Art_IVA "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0038 "
sSQL = sSQL & " WHERE IDOggetto=" & rs!IDOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & rs!IDValoriOggettoDettaglio

Set rsDati = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.cboAliquotaArticolo.WriteOn fnNotNullN(rsDati!Link_Art_Iva)
End If

rsDati.CloseResultset
Set rsDati = Nothing
Exit Sub
ERR_GET_DESCRIZIONI_RIGA_CONTR:
    MsgBox Err.Description, vbCritical, "GET_DESCRIZIONI_RIGA_CONTR"
End Sub
Private Sub GET_IVA_RIGA_IMBALLO(rs As ADODB.Recordset)
On Error GoTo ERR_GET_DESCRIZIONI_RIGA_CONTR
Dim sSQL As String
Dim rsDati As DmtOleDbLib.adoResultset
Dim link_riga As Long

link_riga = 0

sSQL = "SELECT IDValoriOggettoDettaglio, RV_POLinkRiga "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0038 "
sSQL = sSQL & " WHERE IDOggetto=" & rs!IDOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & rs!IDValoriOggettoDettaglio

Set rsDati = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    link_riga = fnNotNullN(rsDati!RV_POLinkRiga)
End If

rsDati.CloseResultset
Set rsDati = Nothing

If (link_riga > 0) Then
    sSQL = "SELECT IDValoriOggettoDettaglio, Art_aliquota_IVA, Link_Art_IVA "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0038 "
    sSQL = sSQL & " WHERE IDOggetto=" & rs!IDOggetto
    sSQL = sSQL & " AND RV_POLinkRiga=" & link_riga
    sSQL = sSQL & " AND RV_POTipoRiga=2"
    
    Set rsDati = Cn.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        Me.CboAliquotaImballo.WriteOn fnNotNullN(rsDati!Link_Art_Iva)
    End If
    
    rsDati.CloseResultset
    Set rsDati = Nothing
End If

Exit Sub
ERR_GET_DESCRIZIONI_RIGA_CONTR:
    MsgBox Err.Description, vbCritical, "GET_DESCRIZIONI_RIGA_CONTR"
End Sub
Public Function GET_DESCRIZIONE_LOTTO_PROD_LAV(IDLotto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_DESCRIZIONE_LOTTO_PROD_LAV = ""

sSQL = "SELECT IDRV_PO01_LottoCampagna, CodiceLotto, DescrizioneLotto "
sSQL = sSQL & "FROM RV_PO01_LottoCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLotto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_DESCRIZIONE_LOTTO_PROD_LAV = fnNotNull(rs!CodiceLotto)
    If Len(fnNotNull(rs!DescrizioneLotto)) > 0 Then
        GET_DESCRIZIONE_LOTTO_PROD_LAV = GET_DESCRIZIONE_LOTTO_PROD_LAV & " (" & fnNotNull(rs!DescrizioneLotto) & ")"
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub cmdGestioneQualitaLista()
On Error GoTo ERR_cmdGestioneQualita_Click

    If oDoc.IDOggetto <= 0 Then Exit Sub

    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoApertura", "2"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoQualitaGestione", "10"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDAnagrafica", Me.cdAnagrafica.KeyFieldID
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDRiferimento", oDoc.IDOggetto
    
    Shell MenuOptions.ProgramsPath & "\RV_POQualitaGestioneNew.exe"
    
Exit Sub
ERR_cmdGestioneQualita_Click:
    MsgBox Err.Description, vbCritical, "cmdGestioneQualitaLista_Click"
End Sub
Private Sub cmdGestioneQualita()
On Error GoTo ERR_cmdGestioneQualita_Click

    If oDoc.IDOggetto <= 0 Then Exit Sub

    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoApertura", "1"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoQualitaGestione", "10"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDAnagrafica", Me.cdAnagrafica.KeyFieldID
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDRiferimento", oDoc.IDOggetto

    Shell MenuOptions.ProgramsPath & "\RV_POQualitaGestioneNew.exe"
    
Exit Sub
ERR_cmdGestioneQualita_Click:
    MsgBox Err.Description, vbCritical, "cmdGestioneQualita_Click"
End Sub
Private Sub CREA_RECORDSET_RETTIFICA()
On Error GoTo ERR_CREA_RECORDSET_RETTIFICA
Dim sSQL As String
Dim I As Long
Dim rsImp As ADODB.Recordset

If Not (rsRettifica Is Nothing) Then
    If rsRettifica.State > 0 Then
        rsRettifica.Clone
    End If
    
    Set rsRettifica = Nothing
End If

Set rsRettifica = New ADODB.Recordset
rsRettifica.CursorLocation = adUseClient

sSQL = "SELECT * FROM RV_POCorpoOrdineRettifica"
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=0"

Set rsImp = New ADODB.Recordset

rsImp.Open sSQL, Cn.InternalConnection


''''CREA TABELLA''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For I = 0 To rsImp.Fields.Count - 1
    If (rsImp.Fields(I).Name <> "ID") Then
        Select Case rsImp.Fields(I).Type
            Case adChar, adVarChar, adVarWChar, adWChar, 201
                rsRettifica.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, rsImp.Fields(I).DefinedSize, rsImp.Fields(I).Attributes
            Case adInteger
                rsRettifica.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, , rsImp.Fields(I).Attributes
            Case adDate, adDBDate, adDBTime, adDBTimeStamp
                rsRettifica.Fields.Append rsImp.Fields(I).Name, rsImp.Fields(I).Type, , rsImp.Fields(I).Attributes
            Case adBoolean, adSmallInt, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt
                rsRettifica.Fields.Append rsImp.Fields(I).Name, adBoolean, , rsImp.Fields(I).Attributes
            Case adNumeric, adSingle, adBigInt, adCurrency, adDouble, adDecimal
                rsRettifica.Fields.Append rsImp.Fields(I).Name, adDouble, , rsImp.Fields(I).Attributes
        End Select
    End If
Next

rsImp.Close
Set rsImp = Nothing

rsRettifica.Open , , adOpenKeyset, adLockBatchOptimistic

Exit Sub
ERR_CREA_RECORDSET_RETTIFICA:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_RETTIFICA"
End Sub
Private Sub SALVA_RETTIFICA()
On Error GoTo ERR_SALVA_RETTIFICA
Dim I As Long
Dim rsImp As ADODB.Recordset
Dim sSQL As String

sSQL = "SELECT * FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND RV_POLinkRiga=" & oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio)
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rsImp = New ADODB.Recordset

rsImp.Open sSQL, Cn.InternalConnection

rsRettifica.Filter = "IDOggetto=" & oDoc.IDOggetto
rsRettifica.Filter = rsRettifica.Filter & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
rsRettifica.Filter = rsRettifica.Filter & " AND RV_POLinkRiga=" & oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio)

If rsRettifica.EOF Then
    rsRettifica.AddNew
        For I = 0 To rsRettifica.Fields.Count - 1
            Select Case rsRettifica.Fields(I).Name
                Case "DataModifica"
                    rsRettifica.Fields(I).Value = Now
                Case "UtenteModifica"
                    rsRettifica.Fields(I).Value = TheApp.User
                Case "NumeroRettifica"
                    rsRettifica.Fields(I).Value = GET_NUMERO_RETTIFICHE(oDoc.IDOggetto, oDoc.IDTipoOggetto, oDoc.Field("RV_POLinkRiga", , sTabellaDettaglio)) + 1
                Case Else
                    rsRettifica.Fields(I).Value = rsImp.Fields(rsRettifica.Fields(I).Name).Value
            End Select
        Next
    rsRettifica.Update
End If

rsImp.Close
Set rsImp = Nothing
Exit Sub
ERR_SALVA_RETTIFICA:
    MsgBox Err.Description, vbCritical, "SALVA_RETTIFICA"
End Sub
Private Sub CONSOLIDA_RETTIFICA()
On Error GoTo ERR_CONSOLIDA_RETTIFICA
Dim I As Long
Dim rsImp As ADODB.Recordset
Dim sSQL As String


rsRettifica.Filter = vbNullString

If ((rsRettifica.EOF) And (rsRettifica.BOF)) Then Exit Sub

sSQL = "SELECT * FROM RV_POCorpoOrdineRettifica"
sSQL = sSQL & " WHERE ID=0"

Set rsImp = New ADODB.Recordset

rsImp.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsRettifica.EOF
    rsImp.AddNew
        For I = 0 To rsRettifica.Fields.Count - 1
            rsImp.Fields(rsRettifica.Fields(I).Name).Value = rsRettifica.Fields(I).Value
        Next
    rsImp.Update
rsRettifica.MoveNext
Wend

rsImp.Close
Set rsImp = Nothing

CREA_RECORDSET_RETTIFICA
Exit Sub
ERR_CONSOLIDA_RETTIFICA:
    MsgBox Err.Description, vbCritical, "CONSOLIDA_RETTIFICA"
End Sub
Private Function GET_NUMERO_RETTIFICHE(IDOggetto As Long, IDTipoOggetto As Long, linkRiga As Long) As Long
On Error GoTo ERR_GET_NUMERO_RETTIFICHE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_NUMERO_RETTIFICHE = 0

sSQL = "SELECT COUNT(ID) as Numero "
sSQL = sSQL & "FROM RV_POCorpoOrdineRettifica "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND RV_POLinkRiga=" & linkRiga

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_NUMERO_RETTIFICHE = fnNotNullN(rs!Numero)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_NUMERO_RETTIFICHE:
    GET_NUMERO_RETTIFICHE = -1
End Function
Private Sub sbPopalaListaCommissioni()
    Dim lRow As Long
    Dim oItem As MSComctlLib.ListItem
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset

    'Pilisce la listview
    lvwCommissioni.ListItems.Clear
    
    If (oDoc.IDOggetto = 0) Then Exit Sub
        
    sSQL = "SELECT * FROM RV_POIECommissioniPerDoc "
    sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
    
    Set rs = Cn.OpenResultset(sSQL)

    While Not rs.EOF
        Set oItem = lvwCommissioni.ListItems.Add
        'Popola l'item della listview
        oItem.Text = fnNotNull(rs!TipoCommissione)
        oItem.SubItems(1) = FormatNumber(fnNotNullN(rs!PercentualeDaCommissione), 5)
        oItem.SubItems(2) = FormatNumber(fnNotNullN(rs!Quantita), 5)
        oItem.SubItems(3) = FormatNumber(fnNotNullN(rs!ImportoRiga), 5)
        oItem.SubItems(4) = fnNotNull(rs!CodiceTipoPedana)
        oItem.SubItems(5) = FormatNumber(fnNotNullN(rs!Percentuale), 5)
        oItem.SubItems(6) = GET_TIPO_VALORE_DOC_COMM(fnNotNullN(rs!IDRV_POTipoCommissione))
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
End Sub
Private Function GET_TIPO_VALORE_DOC_COMM(IDCommissione As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POTipoCommissione.IDRV_POTipoCommissione, RV_POTipoValoreDocumento.IDRV_POTipoValoreDocumento, RV_POTipoValoreDocumento.TipoValoreDocumento "
sSQL = sSQL & "FROM RV_POTipoCommissione LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoValoreDocumento ON RV_POTipoCommissione.IDRV_POTipoValoreDocumento = RV_POTipoValoreDocumento.IDRV_POTipoValoreDocumento "
sSQL = sSQL & "WHERE RV_POTipoCommissione.IDRV_POTipoCommissione=" & IDCommissione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_VALORE_DOC_COMM = ""
Else
    GET_TIPO_VALORE_DOC_COMM = fnNotNull(rs!TipoValoreDocumento)
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
sSQL = sSQL & "FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE (IDOggetto = " & oDoc.IDOggetto & ") "
'sSQL = sSQL & " AND (RV_POTipoRiga > 0)"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO = GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO + (fnNotNullN(rs!Art_importo_net_sconto_net_IVA))
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_TIPO_VALORE_COMMISSIONE(IDTipoCommissione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoValoreDocumento FROM RV_POTipoCommissione "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_VALORE_COMMISSIONE = 1
Else
    If fnNotNullN(rs!IDRV_POTipoValoreDocumento) = 0 Then
        GET_TIPO_VALORE_COMMISSIONE = 1
    Else
        GET_TIPO_VALORE_COMMISSIONE = fnNotNullN(rs!IDRV_POTipoValoreDocumento)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TIPO_RICALCOLO_COMMISSIONE(IDTipoCommissione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTipoCommissione "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_RICALCOLO_COMMISSIONE = 2
Else
    If fnNotNullN(rs!IDRV_POTipoRicalcoloComm) = 0 Then
        GET_TIPO_RICALCOLO_COMMISSIONE = 2
    Else
        GET_TIPO_RICALCOLO_COMMISSIONE = fnNotNullN(rs!IDRV_POTipoRicalcoloComm)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub Command2_Click()
On Error GoTo ERR_Command2_Click
Dim Testo As String
    
    If (oDoc.IDOggetto = 0) Then Exit Sub
    If txtNListaPrelievo.Value > 1 Then Exit Sub
    If ATTIVA_COMMISSIONI_DA_ORDINE = 0 Then Exit Sub
    
    If (m_Changed = True) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Per procedere con la gestione delle commissioni è consigliabile salvare il documento, poichè alla chiusura, il documento potrebbe essere salvato automaticamente"
        Testo = Testo & "Vuoi salvare il documento?" & vbCrLf
        
        If (MsgBox(Testo, vbQuestion + vbYesNo, "Gestione commissioni") = vbYes) Then OnSave
        
    End If
    
    TOTALE_MERCE = 0 ' GET_TOTALE_MERCE_DOCUMENTO
    TOTALE_MERCE_LORDA = 0 ' GET_TOTALE_MERCE_DOCUMENTO_INCLUSO_IMBALLO
    TOTALE_DOCUMENTO_NETTO_IVA = 0 ' curTotImponibile.Value
    TOTALE_DOCUMENTO_LORDO_IVA = 0 ' curTotDocumento.Value
    
    frmCommissioni.Show vbModal
    
    If (Changed_Commissioni = True) Then OnSave
    
Exit Sub
ERR_Command2_Click:
    MsgBox Err.Description, vbCritical, "Gestione commissioni"
End Sub
Private Sub IMPOSTA_COMMISSIONI_PER_TIPO_PEDANA(idcliente As Long, IDOggettoDocumento As Long, IDSitoPerAnagrafica As Long)
On Error GoTo ERR_IMPOSTA_COMMISSIONI_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim rsCli As ADODB.Recordset
Dim rsPed As ADODB.Recordset
Dim rsQuantita As ADODB.Recordset

Dim SpesaTrasporto As Double
Dim Totale_merce_lavorato As Double

sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT IDOggetto, RV_POIDTipoPedana, RV_POTipoPedana.IDArticoloImballo,"
sSQL = sSQL & " (SELECT COUNT(*) AS QuantitaTipoPedana "
sSQL = sSQL & " FROM (SELECT RV_POIDPedana "
sSQL = sSQL & " FROM " & sTabellaDettaglio & " AS Tabella INNER JOIN "
sSQL = sSQL & " RV_POTipoPedana ON Tabella.RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
sSQL = sSQL & " Where (IDOggetto = " & sTabellaDettaglio & ".IDOggetto) And (RV_POTipoRiga = 1) And (RV_POIDTipoPedana = " & sTabellaDettaglio & ".RV_POIDTipoPedana) "
sSQL = sSQL & " GROUP BY Tabella.RV_POIDPedana, RV_POTipoPedana.IDArticoloImballo) AS X) AS QuantitaPedana "
sSQL = sSQL & " FROM " & sTabellaDettaglio & " INNER JOIN "
sSQL = sSQL & " RV_POTipoPedana ON " & sTabellaDettaglio & ".RV_POIDTipoPedana = RV_POTipoPedana.IDRV_POTipoPedana "
sSQL = sSQL & " Where IDOggetto = " & oDoc.IDOggetto
sSQL = sSQL & " AND  RV_POTipoRiga = 1"
sSQL = sSQL & " GROUP BY IDOggetto, RV_POIDTipoPedana, IDArticoloImballo"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "SELECT * FROM RV_POIEConfigurazioneClienteTrasporto "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDAnagrafica=" & idcliente
    sSQL = sSQL & " AND IDSitoPerAnagrafica=" & IDSitoPerAnagrafica
    sSQL = sSQL & " AND IDArticolo=" & fnNotNullN(rs!IDArticoloImballo)
    sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
    sSQL = sSQL & " AND CommissionePerPedana=1"
    
    Set rsCli = New ADODB.Recordset
    
    rsCli.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
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
                    rsNew!IDArticoloImballo = rs!IDArticoloImballo
                rsNew.Update
            End If
        rsCli.MoveNext
        Wend
    Else
        sSQL = "SELECT * FROM RV_POTipoPedana "
        sSQL = sSQL & " WHERE IDRV_POTipoPedana=" & fnNotNullN(rs!RV_POIDTipoPedana)
        sSQL = sSQL & " AND IDArticoloImballo=" & fnNotNullN(rs!IDArticoloImballo)
        sSQL = sSQL & " AND IDRV_POTipoCommissione>0"
        
        Set rsPed = New ADODB.Recordset
        rsPed.Open sSQL, Cn.InternalConnection
        
        While Not rsPed.EOF
            If GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED(fnNotNullN(rsPed!IDRV_POTipoCommissione), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNullN(rs!IDArticoloImballo)) = False Then
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
    MsgBox Err.Description, vbCritical, "IMPOSTA_COMMISSIONI_PER_TIPO_PEDANA"

End Sub
Public Function GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED(IDTipoCommissione As Long, IDTipoPedana As Long, IDArticoloImballo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

    sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
    sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione
    sSQL = sSQL & " AND IDRV_POTipoPedana=" & IDTipoPedana
    sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
    sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED = False
    Else
        GET_ESISTENZA_COMMISSIONE_DOC_TIPO_PED = True
    End If
rs.CloseResultset
Set rs = Nothing

End Function
Public Function GET_ESISTENZA_COMMISSIONE_DOC_TIPO_IMB(IDTipoCommissione As Long, IDArticoloImballo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

    sSQL = "SELECT * FROM RV_POCommissioniPerDoc "
    sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione
    sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_ESISTENZA_COMMISSIONE_DOC_TIPO_IMB = False
    Else
        GET_ESISTENZA_COMMISSIONE_DOC_TIPO_IMB = True
    End If
rs.CloseResultset
Set rs = Nothing

End Function
Public Sub RiepilogoTotaliDocumento()
On Error GoTo ERR_RiepilogoTotaliDocumento
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Me.txtTaraTotale.Value = 0
Me.txtPezziTotale.Value = 0
Me.txtPesoNettoTotale.Value = 0
Me.txtNPedaneTotale.Value = 0

sSQL = "SELECT SUM(Art_tara) as TotaleTara, "
sSQL = sSQL & "SUM(Art_quantita_pezzi) as TotalePezzi, "
sSQL = sSQL & "SUM(RV_POQuantitaPedanaEffettiva) as TotalePedana "
sSQL = sSQL & " FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND RV_POTipoRiga=1"
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Me.txtTaraTotale.Value = fnNotNullN(rs!TotaleTara)
    Me.txtPezziTotale.Value = fnNotNullN(rs!TotalePezzi)
    Me.txtNPedaneTotale.Value = fnNotNullN(rs!TotalePedana)
End If

rs.CloseResultset
Set rs = Nothing

Me.txtPesoNettoTotale.Value = Me.txtPesoTotale.Value - Me.txtTaraTotale.Value

Exit Sub
ERR_RiepilogoTotaliDocumento:
    MsgBox Err.Description, vbCritical, "RiepilogoTotaliDocumento"
End Sub
Private Sub SalvaTipoLavorazioniUtilizzare(IDOggetto As Long)
On Error GoTo ERR_SalvaTipoLavorazioniUtilizzare
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsControllo As DmtOleDbLib.adoResultset
Dim SalvaTipoLav As Boolean

sSQL = "DELETE FROM RV_POOrdineTipoLavorazione "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
Cn.Execute sSQL

''''RECUPERO LINEE DALLE RIGHE ORDINE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDValoriOggettoDettaglio, RV_POIDTipoLavorazione "
sSQL = sSQL & " FROM ValoriOggettoDettaglio0010 "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    SalvaTipoLav = False
    sSQL = "SELECT * FROM RV_POOrdineTipoLavorazione "
    sSQL = sSQL & " WHERE IDRV_POTipoLavorazione=" & fnNotNullN(rs!RV_POIDTipoLavorazione)
    sSQL = sSQL & " AND IDOggetto=" & IDOggetto
    
    Set rsControllo = Cn.OpenResultset(sSQL)
    
    If rsControllo.EOF Then
        SalvaTipoLav = True
    End If
    
    rsControllo.CloseResultset
    Set rsControllo = Nothing
    
    If SalvaTipoLav = True Then
        sSQL = "INSERT INTO RV_POOrdineTipoLavorazione "
        sSQL = sSQL & "(IDOggetto, IDRV_POTipoLavorazione) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDOggetto & ", "
        sSQL = sSQL & fnNotNullN(rs!RV_POIDTipoLavorazione) & ")"
        
        Cn.Execute sSQL
    End If
    
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''RECUPERO LINEE DALLE RIGHE RETTIFICATE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDValoriOggettoDettaglio, RV_POIDTipoLavorazione "
sSQL = sSQL & " FROM RV_POCorpoOrdineRettifica "
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    SalvaTipoLav = False
    sSQL = "SELECT * FROM RV_POOrdineTipoLavorazione "
    sSQL = sSQL & " WHERE IDRV_POTipoLavorazione=" & fnNotNullN(rs!RV_POIDTipoLavorazione)
    sSQL = sSQL & " AND IDOggetto=" & IDOggetto
    
    Set rsControllo = Cn.OpenResultset(sSQL)
    
    If rsControllo.EOF Then
        SalvaTipoLav = True
    End If
    
    rsControllo.CloseResultset
    Set rsControllo = Nothing
    
    If SalvaTipoLav = True Then
        sSQL = "INSERT INTO RV_POOrdineTipoLavorazione "
        sSQL = sSQL & "(IDOggetto, IDRV_POTipoLavorazione) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & IDOggetto & ", "
        sSQL = sSQL & fnNotNullN(rs!RV_POIDTipoLavorazione) & ")"
        
        Cn.Execute sSQL
    End If
    
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Exit Sub
ERR_SalvaTipoLavorazioniUtilizzare:
    MsgBox Err.Description, vbCritical, "SalvaTipoLavorazioniUtilizzare"
End Sub

Private Sub GET_PAR_CONN_ALY()
On Error GoTo ERR_GET_PAR_CONN_ALY
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

NOME_SERVER_ALY = ""
NOME_DB_ALY = ""
USER_PROP_SERVER = ""
PWD_USER_PROP = ""
GRUPPO_ALY = ""
USER_ALY = ""
PWD_USER_ALY = ""

sSQL = "SELECT * FROM RV_POAlyanteParametriConnessione "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    NOME_SERVER_ALY = fnNotNull(rs!NomeServer)
    NOME_DB_ALY = fnNotNull(rs!DBName)
    USER_PROP_SERVER = fnNotNull(rs!DBUser)
    PWD_USER_PROP = fnNotNull(rs!Password)
    GRUPPO_ALY = fnNotNull(rs!UserGroup)
    USER_ALY = fnNotNull(rs!ApplicationUser)
    PWD_USER_ALY = fnNotNull(rs!ApplicationPassword)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_GET_PAR_CONN_ALY:
    MsgBox Err.Description, vbCritical, "GET_PAR_CONN_ALY"
End Sub
Private Sub GET_PARAMETRI_FIDO_ALYANTE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    COMPANY_CODE = 0
    ATTIVA_FIDO_ALY = 0
    DISATTIVA_DDT_FIDO_ALY = 0
    DISATTIVA_FD_FIDO_ALY = 0
    DISATTIVA_FA_FIDO_ALY = 0
    DISATTIVA_NC_FIDO_ALY = 0
    DISATTIVA_ND_FIDO_ALY = 0
Else
    COMPANY_CODE = fnNotNullN(rs!Alyante_CompanyCode)
    ATTIVA_FIDO_ALY = fnNotNullN(rs!Alyante_AttivaFido)
    DISATTIVA_DDT_FIDO_ALY = fnNotNullN(rs!Alyante_DisattivaCalcoloDDT)
    DISATTIVA_FD_FIDO_ALY = fnNotNullN(rs!Alyante_DisattivaCalcoloFD)
    DISATTIVA_FA_FIDO_ALY = fnNotNullN(rs!Alyante_DisattivaCalcoloFA)
    DISATTIVA_NC_FIDO_ALY = fnNotNullN(rs!Alyante_DisattivaCalcoloNC)
    DISATTIVA_ND_FIDO_ALY = fnNotNullN(rs!Alyante_DisattivaCalcoloND)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function CONTROLLO_FIDO_ALYANTE(IDAnagraficaCliente As Long) As Boolean
On Error GoTo ERR_CONTROLLO_PLAFOND_ALYANTE
Dim IDPDCCliente As Long
Dim ValoreMappatura As String
Dim Testo As String

CONTROLLO_FIDO_ALYANTE = False

ALY_FIDO_CALCOLATO = 0
ALY_FIDO_CLIENTE = 0
ALY_FIDO_RESIDUO = 0
ALY_FUORI_FIDO = False
ALY_TOTALE_DOC_PREC = 0
ALY_FIDO_CALCOLATO_ALY = 0
ALY_FIDO_RESIDUO_ALY = 0
ALY_FIDO_TOT_DDT = 0
ALY_FIDO_TOT_FA = 0
ALY_FIDO_TOT_FD = 0
ALY_FIDO_TOT_NC = 0
ALY_FIDO_TOT_ND = 0

ALY_TOTALE_DOC = oDoc.Field("Tot_netto_a_pagare_corr", , sTabellaTestata)

If (AttivaMappaturaDaRegFatture = 0) Then
    IDPDCCliente = GET_PDC_CLIENTE(IDAnagraficaCliente)
    If (IDPDCCliente > 0) Then
        ValoreMappatura = GET_MAPPATURA_CLIENTE(IDPDCCliente)
        If (Len(ValoreMappatura) > 0) Then
            Set i_objActiveInterface = GET_CONNESSIONE_ALYANTE(False)
            If Not (i_objActiveInterface Is Nothing) Then
                If (GET_ATTIVAZIONE_FIDO_ALY(ValoreMappatura) = True) Then
                    CalcolaFidoAlyante ValoreMappatura, IDAnagraficaCliente
                    Select Case ALY_TIPO_SEGNALAZIONE_FIDO
                        Case 0
                            CONTROLLO_FIDO_ALYANTE = False
                        Case 1
                            CONTROLLO_FIDO_ALYANTE = ALY_FUORI_FIDO
                        Case 2
                            CONTROLLO_FIDO_ALYANTE = True
                    End Select
                End If
                DeallocaOggettiAlyante
            End If
        End If
    End If
Else
    ValoreMappatura = GET_MAPPATURA_CLIENTE_REG_FATT(IDAnagraficaCliente)
    If (Len(ValoreMappatura) > 0) Then
        Set i_objActiveInterface = GET_CONNESSIONE_ALYANTE(False)
        If Not (i_objActiveInterface Is Nothing) Then
            If (GET_ATTIVAZIONE_FIDO_ALY(ValoreMappatura) = True) Then
                CalcolaFidoAlyante ValoreMappatura, IDAnagraficaCliente
                Select Case ALY_TIPO_SEGNALAZIONE_FIDO
                    Case 0
                        CONTROLLO_FIDO_ALYANTE = False
                    Case 1
                        CONTROLLO_FIDO_ALYANTE = ALY_FUORI_FIDO
                    Case 2
                        CONTROLLO_FIDO_ALYANTE = True
                End Select
            End If
            DeallocaOggettiAlyante
        End If
    End If
End If
Exit Function
ERR_CONTROLLO_PLAFOND_ALYANTE:
CONTROLLO_FIDO_ALYANTE = False
MsgBox Err.Description, vbCritical, "CONTROLLO_PLAFOND_ALYANTE"
End Function
Private Function GET_PDC_CLIENTE(IDAnagraficaCliente As Long) As Long
On Error GoTo ERR_GET_PDC_CLIENTE
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim CodicePDC As Long

GET_PDC_CLIENTE = 0

CodicePDC = 0
sSQL = "SELECT IDPDCContabile "
sSQL = sSQL & "FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagraficaCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_PDC_CLIENTE = fnNotNullN(rs!IDPDCContabile)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_PDC_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_PDC_CLIENTE"
End Function

Private Function GET_MAPPATURA_CLIENTE(IDPDC As Long) As String
On Error GoTo ERR_GET_MAPPATURA_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_MAPPATURA_CLIENTE = ""

sSQL = "SELECT dbo.RV_PO99_Mappatura.IDRV_PO99_Mappatura, dbo.RV_PO99_MappaturaRighe.IDRV_PO99_MappaturaRighe, dbo.RV_PO99_Mappatura.IDAzienda, dbo.RV_PO99_Mappatura.IDFiliale, "
sSQL = sSQL & "dbo.RV_PO99_Mappatura.IDRV_PO99_TipoMappatura, dbo.RV_PO99_Mappatura.IDRV_PO99_TipoFile, dbo.RV_PO99_Mappatura.IDMappatura, dbo.RV_PO99_MappaturaRighe.Valore,"
sSQL = sSQL & "dbo.RV_PO99_MappaturaRighe.IDRV_PO99_ConfigFile , dbo.RV_PO99_ConfigFile.NomeCampo "
sSQL = sSQL & "FROM dbo.RV_PO99_Mappatura INNER JOIN "
sSQL = sSQL & "dbo.RV_PO99_MappaturaRighe ON dbo.RV_PO99_Mappatura.IDRV_PO99_Mappatura = dbo.RV_PO99_MappaturaRighe.IDRV_PO99_Mappatura INNER JOIN "
sSQL = sSQL & "dbo.RV_PO99_ConfigFile ON dbo.RV_PO99_MappaturaRighe.IDRV_PO99_ConfigFile = dbo.RV_PO99_ConfigFile.IDRV_PO99_ConfigFile "
sSQL = sSQL & " WHERE dbo.RV_PO99_Mappatura.IDAzienda = " & TheApp.IDFirm
sSQL = sSQL & " AND dbo.RV_PO99_Mappatura.IDFiliale = " & TheApp.Branch
sSQL = sSQL & " AND dbo.RV_PO99_Mappatura.IDRV_PO99_TipoMappatura = 1"
sSQL = sSQL & " AND dbo.RV_PO99_Mappatura.IDRV_PO99_TipoFile = 7 "
sSQL = sSQL & " AND dbo.RV_PO99_ConfigFile.NomeCampo = 'CodiceAnagrafica' "
sSQL = sSQL & " AND dbo.RV_PO99_Mappatura.IDMappatura = " & IDPDC

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_MAPPATURA_CLIENTE = fnNotNull(rs!Valore)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_MAPPATURA_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_MAPPATURA_CLIENTE"
End Function
Private Function GET_CONNESSIONE_ALYANTE(ConMessaggio As Boolean) As Cinterface
On Error GoTo ERR_cmdTest_Click

Set i_objNucleo = New IEBO_NUCLEO.CLSIE_NUCLEO

i_objNucleo.NomeServer = NOME_SERVER_ALY
i_objNucleo.NomeDB = NOME_DB_ALY
i_objNucleo.UserIdDB = USER_PROP_SERVER
i_objNucleo.PasswordDB = PWD_USER_PROP
i_objNucleo.GruppoUtenti = GRUPPO_ALY
i_objNucleo.UtenteCorrente = USER_ALY
i_objNucleo.PwdUtente = PWD_USER_ALY
i_objNucleo.CodiceDitta = COMPANY_CODE

If Not i_objNucleo.Inizializza() Then
    If ConMessaggio = True Then
        MsgBox "Errore di inizializzazione!", vbCritical, "Verifica connessione"
    End If
    GET_CONNESSIONE_ALYANTE = False
Else
    If ConMessaggio = True Then
        MsgBox "Connessione avvenuta con successo!", vbInformation, "Verifica connessione"
    End If
End If


Set GET_CONNESSIONE_ALYANTE = i_objNucleo.ObjectInterface
Exit Function
ERR_cmdTest_Click:
    GET_CONNESSIONE_ALYANTE = False
    MsgBox Err.Description, vbCritical, "Connessione ad Alyante non avvenuta "
End Function
Private Function GET_ATTIVAZIONE_FIDO_ALY(codicecliente As String) As Boolean
On Error GoTo ERR_GET_ATTIVAZIONE_FIDO_ALY
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim cnAly As ADODB.Connection
Dim IndicaFido As Long

GET_ATTIVAZIONE_FIDO_ALY = False
IndicaFido = 1

Set cnAly = New ADODB.Connection
cnAly.Open i_objNucleo.StringaConnessione

sSQL = "SELECT * "
sSQL = sSQL & " FROM MG19_CLIFORVA"
sSQL = sSQL & " WHERE MG19_DITTA_CG18=" & fnNormNumber(COMPANY_CODE)
sSQL = sSQL & " AND MG19_TIPOCF_CG44 = 0 "
sSQL = sSQL & " AND MG19_CLIFOR_CG44 = " & fnNormNumber(codicecliente)

Set rs = New ADODB.Recordset

rs.Open sSQL, cnAly

If Not rs.EOF Then
    IndicaFido = fnNotNullN(rs!MG19_INDGESFIDO)
    If Len(Trim(fnNotNull(rs!MG19_CODRISCHIO_MG2A))) > 0 Then
        If fnNotNullN(rs!MG19_INDGESFIDO) = 2 Then
            GET_ATTIVAZIONE_FIDO_ALY = True
        End If
    End If
End If

rs.Close
Set rs = Nothing

If (IndicaFido = 1) Then Exit Function


If GET_ATTIVAZIONE_FIDO_ALY = False Then
    sSQL = "SELECT * "
    sSQL = sSQL & " FROM MG3A_DOCPARGEN"
    sSQL = sSQL & " WHERE MG3A_DITTA_CG18=" & fnNormNumber(COMPANY_CODE)
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, cnAly
    
    If Not rs.EOF Then
        If (fnNotNullN(rs!MG3A_INDGESFIDOCL) = 1) Then
            If Len(Trim(fnNotNull(rs!MG3A_CODRISCHIOCLI_MG2A))) > 0 Then
                GET_ATTIVAZIONE_FIDO_ALY = True
            End If
        End If
    End If
        
    rs.Close
    Set rs = Nothing
End If


cnAly.Close
Set cnAly = Nothing
Exit Function
ERR_GET_ATTIVAZIONE_FIDO_ALY:
    MsgBox Err.Description, vbCritical, "GET_ATTIVAZIONE_FIDO_ALY"
End Function
Private Sub CalcolaFidoAlyante(codicecliente As String, IDAnagraficaCliente As Long)
On Error GoTo ERR_CalcolaFidoAlyante
Dim TotaliDDT As Double

Set i_objCalcRischio = New MGBO_CALCRISCHIO.CLSMG_CALCRISCHIO

Set i_objCalcRischio.ActiveInterface = i_objActiveInterface
Set i_objCalcRischio.ActiveConnection = i_objActiveInterface.Connection
   
i_objCalcRischio.CodiceDitta = i_objActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
i_objCalcRischio.GruppoPDC = i_objActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.GruppoPDC ' 80
i_objCalcRischio.MastroClienti = i_objActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroClienti '"1400000000"
i_objCalcRischio.MastroFornitori = i_objActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsPersPdc.MastroFornitori '"4000000000"
i_objCalcRischio.ModalitaSegnalazione = Disattiva
i_objCalcRischio.TipoElaborazioneRischio = ElaborazioneRischioSingola
i_objCalcRischio.DataElaborazione = Date
i_objCalcRischio.ForzaRicalcoloDati = True

i_objCalcRischio.TipoCF = Cliente
i_objCalcRischio.CodiceCF = codicecliente

i_objCalcRischio.ElaborazioneRischio

ALY_FIDO_CLIENTE = i_objCalcRischio.Importi.FidoCliente
ALY_FIDO_CALCOLATO = i_objCalcRischio.Importi.FidoCalcolato
ALY_FIDO_RESIDUO = i_objCalcRischio.Importi.FidoResiduo
ALY_FUORI_FIDO = i_objCalcRischio.IsFuoriFido
ALY_TIPO_SEGNALAZIONE_FIDO = i_objCalcRischio.Importi.ModalitaAggiornamentoBlocco

ALY_FIDO_CALCOLATO_ALY = ALY_FIDO_CALCOLATO
ALY_FIDO_RESIDUO_ALY = ALY_FIDO_RESIDUO
ALY_TIPO_SEGNALAZIONE_FIDO = GET_MOD_ATT_FIDO_SU_DOC_ALY(i_objCalcRischio.CodiceRischio)

Set i_objCalcRischio = Nothing

ALY_FIDO_TOT_DDT = GET_TOTALE_DDT(IDAnagraficaCliente)
ALY_FIDO_TOT_FA = GET_TOTALE_FA_FIDO_ALY(IDAnagraficaCliente)
ALY_FIDO_TOT_FD = GET_TOTALE_FD_FIDO_ALY(IDAnagraficaCliente)
ALY_FIDO_TOT_NC = GET_TOTALE_NC_FIDO_ALY(IDAnagraficaCliente)
ALY_FIDO_TOT_ND = GET_TOTALE_ND_FIDO_ALY(IDAnagraficaCliente)
ALY_TOTALE_DOC_PREC = GET_TOTALE_DOCUMENTO_PREC(oDoc.IDOggetto)


ALY_FIDO_CALCOLATO = ALY_FIDO_CALCOLATO + ALY_FIDO_TOT_DDT + ALY_FIDO_TOT_FA + ALY_FIDO_TOT_FD + ALY_FIDO_TOT_ND - ALY_FIDO_TOT_NC
ALY_FIDO_CALCOLATO = ALY_FIDO_CALCOLATO + ALY_TOTALE_DOC 'oDoc.Field("Tot_netto_a_pagare_corr", , sTabellaTestata)
ALY_FIDO_CALCOLATO = ALY_FIDO_CALCOLATO - ALY_TOTALE_DOC_PREC
'Se il documento è in modifica, ossia non è nuovo, togliere ad ALY_FIDO_CALCOLATO l'importo vecchio
'If oDoc.IDOggetto > 0 Then
'    ALY_FIDO_CALCOLATO = ALY_FIDO_CALCOLATO - GET_TOTALE_DOCUMENTO_PREC(oDoc.IDOggetto)
'End If

ALY_FIDO_RESIDUO = ALY_FIDO_CLIENTE - ALY_FIDO_CALCOLATO
ALY_FUORI_FIDO = ALY_FIDO_RESIDUO < 0
Exit Sub
ERR_CalcolaFidoAlyante:
    MsgBox Err.Description, vbCritical, "CalcolaFidoAlyante"
End Sub

Private Sub DeallocaOggettiAlyante()
On Error GoTo ERR_DeallocaOggettiAlyante
    Set i_objActiveInterface = Nothing
    If Not i_objNucleo Is Nothing Then
       i_objNucleo.Terminate
       Set i_objNucleo.adoConnection = Nothing
    End If
    Set i_objNucleo = Nothing
Exit Sub
ERR_DeallocaOggettiAlyante:
    MsgBox Err.Description, vbCritical, "DeallocaOggettiAlyante"
End Sub
Private Function GET_TOTALE_DOCUMENTO_PREC(IDOggetto As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TOTALE_DOCUMENTO_PREC = 0

sSQL = "SELECT Tot_netto_a_pagare_corr "
sSQL = sSQL & " FROM " & sTabellaTestata
sSQL = sSQL & " WHERE IDOggetto=" & IDOggetto

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_TOTALE_DOCUMENTO_PREC = fnNotNullN(rs!Tot_netto_a_pagare_corr)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_TOTALE_DDT(IDAnagraficaCliente As Long) As Double
On Error GoTo ERR_GET_TOTALE_DDT
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
GET_TOTALE_DDT = 0

If DISATTIVA_DDT_FIDO_ALY = 1 Then Exit Function

sSQL = "SELECT SUM(Tot_netto_a_pagare_corr) as Totale "
sSQL = sSQL & "FROM RV_POIEAlyanteDDTNoFD "
sSQL = sSQL & "WHERE IDFlussoFunzione = 1"
sSQL = sSQL & " AND FlussoFunzioneCollegato = 0"
sSQL = sSQL & " AND Link_Nom_anagrafica=" & IDAnagraficaCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_TOTALE_DDT = fnNotNullN(rs!Totale)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_TOTALE_DDT:
    MsgBox Err.Description, vbCritical, "GET_TOTALE_DDT"
End Function
Private Function GET_TOTALE_FD_FIDO_ALY(IDAnagraficaCliente As Long) As Double
On Error GoTo ERR_GET_TOTALE_FD_FIDO_ALY
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TOTALE_FD_FIDO_ALY = 0

If DISATTIVA_FD_FIDO_ALY = 1 Then Exit Function

sSQL = "SELECT ValoriOggettoPerTipo0004.IDOggetto, ValoriOggettoPerTipo0004.IDTipoOggetto, "
sSQL = sSQL & " ValoriOggettoPerTipo0004.Tot_netto_a_pagare_corr, ValoriOggettoPerTipo0004.Link_Nom_anagrafica, "
sSQL = sSQL & " RV_PO99_DettaglioStoricoExp.IDOggetto AS IDOggettoEsportato "
sSQL = sSQL & " FROM RV_PO99_DettaglioStoricoExp RIGHT OUTER JOIN "
sSQL = sSQL & " ValoriOggettoPerTipo0004 ON RV_PO99_DettaglioStoricoExp.IDOggetto = ValoriOggettoPerTipo0004.IDOggetto "
sSQL = sSQL & " WHERE ValoriOggettoPerTipo0004.Link_Nom_anagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND RV_PO99_DettaglioStoricoExp.IDOggetto IS NULL "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_TOTALE_FD_FIDO_ALY = GET_TOTALE_FD_FIDO_ALY + fnNotNullN(rs!Tot_netto_a_pagare_corr)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_TOTALE_FD_FIDO_ALY:
    MsgBox Err.Description, vbCritical, "GET_TOTALE_FD_FIDO_ALY"
End Function
Private Function GET_TOTALE_FA_FIDO_ALY(IDAnagraficaCliente As Long) As Double
On Error GoTo ERR_GET_TOTALE_FA_FIDO_ALY
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TOTALE_FA_FIDO_ALY = 0
If DISATTIVA_FA_FIDO_ALY = 1 Then Exit Function
sSQL = "SELECT ValoriOggettoPerTipo0072.IDOggetto,"
sSQL = sSQL & " ValoriOggettoPerTipo0072.Tot_netto_a_pagare_corr, ValoriOggettoPerTipo0072.Link_Nom_anagrafica, "
sSQL = sSQL & " RV_PO99_DettaglioStoricoExp.IDOggetto AS IDOggettoEsportato "
sSQL = sSQL & " FROM RV_PO99_DettaglioStoricoExp RIGHT OUTER JOIN "
sSQL = sSQL & " ValoriOggettoPerTipo0072 ON RV_PO99_DettaglioStoricoExp.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto "
sSQL = sSQL & " WHERE ValoriOggettoPerTipo0072.Link_Nom_anagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND RV_PO99_DettaglioStoricoExp.IDOggetto IS NULL "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_TOTALE_FA_FIDO_ALY = GET_TOTALE_FA_FIDO_ALY + fnNotNullN(rs!Tot_netto_a_pagare_corr)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_TOTALE_FA_FIDO_ALY:
    MsgBox Err.Description, vbCritical, "GET_TOTALE_FA_FIDO_ALY"
End Function
Private Function GET_MOD_ATT_FIDO_SU_DOC_ALY(CodiceRischio As String) As Long
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim cnAly As ADODB.Connection

GET_MOD_ATT_FIDO_SU_DOC_ALY = 0

Set cnAly = New ADODB.Connection
cnAly.Open i_objNucleo.StringaConnessione

sSQL = "SELECT * "
sSQL = sSQL & " FROM MG2A_RISCHIO"
sSQL = sSQL & " WHERE MG2A_DITTA_CG18=" & fnNormNumber(COMPANY_CODE)
sSQL = sSQL & " AND MG2A_CODICE = " & fnNormString(CodiceRischio)

Set rs = New ADODB.Recordset

rs.Open sSQL, cnAly

If Not rs.EOF Then
    GET_MOD_ATT_FIDO_SU_DOC_ALY = fnNotNullN(rs!MG2A_INDATTIVADOC)
End If

rs.Close
Set rs = Nothing

cnAly.Close
Set cnAly = Nothing
End Function
Private Function GET_TOTALE_NC_FIDO_ALY(IDAnagraficaCliente As Long) As Double
On Error GoTo ERR_GET_TOTALE_FD_FIDO_ALY
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TOTALE_NC_FIDO_ALY = 0
If DISATTIVA_NC_FIDO_ALY = 1 Then Exit Function
sSQL = "SELECT ValoriOggettoPerTipo000B.IDOggetto, ValoriOggettoPerTipo000B.IDTipoOggetto, "
sSQL = sSQL & " ValoriOggettoPerTipo000B.Tot_netto_a_pagare_corr, ValoriOggettoPerTipo000B.Link_Nom_anagrafica, "
sSQL = sSQL & " RV_PO99_DettaglioStoricoExp.IDOggetto AS IDOggettoEsportato "
sSQL = sSQL & " FROM RV_PO99_DettaglioStoricoExp RIGHT OUTER JOIN "
sSQL = sSQL & " ValoriOggettoPerTipo000B ON RV_PO99_DettaglioStoricoExp.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto "
sSQL = sSQL & " WHERE ValoriOggettoPerTipo000B.Link_Nom_anagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND RV_PO99_DettaglioStoricoExp.IDOggetto IS NULL "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_TOTALE_NC_FIDO_ALY = GET_TOTALE_NC_FIDO_ALY + fnNotNullN(rs!Tot_netto_a_pagare_corr)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_TOTALE_FD_FIDO_ALY:
    MsgBox Err.Description, vbCritical, "GET_TOTALE_NC_FIDO_ALY"
End Function
Private Function GET_TOTALE_ND_FIDO_ALY(IDAnagraficaCliente As Long) As Double
On Error GoTo ERR_GET_TOTALE_FD_FIDO_ALY
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_TOTALE_ND_FIDO_ALY = 0
If DISATTIVA_ND_FIDO_ALY = 1 Then Exit Function
sSQL = "SELECT ValoriOggettoPerTipo006B.IDOggetto, ValoriOggettoPerTipo006B.IDTipoOggetto, "
sSQL = sSQL & " ValoriOggettoPerTipo006B.Tot_netto_a_pagare_corr, ValoriOggettoPerTipo006B.Link_Nom_anagrafica, "
sSQL = sSQL & " RV_PO99_DettaglioStoricoExp.IDOggetto AS IDOggettoEsportato "
sSQL = sSQL & " FROM RV_PO99_DettaglioStoricoExp RIGHT OUTER JOIN "
sSQL = sSQL & " ValoriOggettoPerTipo006B ON RV_PO99_DettaglioStoricoExp.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto "
sSQL = sSQL & " WHERE ValoriOggettoPerTipo006B.Link_Nom_anagrafica=" & IDAnagraficaCliente
sSQL = sSQL & " AND RV_PO99_DettaglioStoricoExp.IDOggetto IS NULL "

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    GET_TOTALE_ND_FIDO_ALY = GET_TOTALE_ND_FIDO_ALY + fnNotNullN(rs!Tot_netto_a_pagare_corr)
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_TOTALE_FD_FIDO_ALY:
    MsgBox Err.Description, vbCritical, "GET_TOTALE_ND_FIDO_ALY"
End Function
Private Function GET_MAPPATURA_CLIENTE_REG_FATT(ID As Long) As String
On Error GoTo ERR_GET_MAPPATURA_CLIENTE
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsMappatura As ADODB.Recordset
Dim cnRegFatt As ADODB.Connection
Dim StringaDiConnessione As String

StringaDiConnessione = "Provider=SQLOLEDB.1;SERVER=" & NOME_SERVER_ALY & ";Initial Catalog=" & DBNameRegFatture & ";User ID=" & USER_PROP_SERVER & ";Password=" & PWD_USER_PROP & ";Persist Security Info=True"
Set cnRegFatt = New ADODB.Connection
cnRegFatt.ConnectionString = StringaDiConnessione
cnRegFatt.Open

GET_MAPPATURA_CLIENTE_REG_FATT = ""

sSQL = "SELECT GuidMappatura "
sSQL = sSQL & "FROM PO_TabellaRaccordoDatiGestionale "
sSQL = sSQL & " WHERE IDPOS_TipoTabellaRaccordo=1"
sSQL = sSQL & " AND IDGestionale=1"
sSQL = sSQL & " AND Chiave1=" & fnNormString(ID)
sSQL = sSQL & " AND RIF_IDAzienda=" & fnNormString(TheApp.IDFirm)

Set rs = New ADODB.Recordset
rs.Open sSQL, cnRegFatt

If (Not rs.EOF) Then
    sSQL = "SELECT Chiave2 "
    sSQL = sSQL & "FROM PO_TabellaRaccordoDatiGestionale "
    sSQL = sSQL & " WHERE IDPOS_TipoTabellaRaccordo=1"
    sSQL = sSQL & " AND IDGestionale=2"
    sSQL = sSQL & " AND GuidMappatura=" & fnNormString(rs!GuidMappatura)
    
    Set rsMappatura = New ADODB.Recordset
    rsMappatura.Open sSQL, cnRegFatt
    If (Not rsMappatura.EOF) Then
        GET_MAPPATURA_CLIENTE_REG_FATT = rsMappatura!Chiave2
    End If
    
    rsMappatura.Close
    Set rsMappatura = Nothing
End If

rs.Close
Set rs = Nothing

cnRegFatt.Close
Set cnRegFatt = Nothing

Exit Function
ERR_GET_MAPPATURA_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_MAPPATURA_CLIENTE"
    
    If Not (rs Is Nothing) Then
        Set rs = Nothing
    End If
    
    If Not (cnRegFatt Is Nothing) Then
        Set cnRegFatt = Nothing
    End If
End Function
Private Sub AltriParametriCooperativa()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDClassificazioneLottoProdPerFuoriQuota, MsgInDocSeRigaMerceSenzaImballo, "
sSQL = sSQL & "DBNameRegFatture, AttivaMappaturaDaRegFatture "
sSQL = sSQL & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    IDClassLottoProdPerFuoriQuota = fnNotNullN(rs!IDClassificazioneLottoProdPerFuoriQuota)
    MsgInDocSeRigaMerceSenzaImballo = fnNotNullN(rs!MsgInDocSeRigaMerceSenzaImballo)
    AttivaMappaturaDaRegFatture = fnNotNullN(rs!AttivaMappaturaDaRegFatture)
    DBNameRegFatture = fnNotNull(rs!DBNameRegFatture)
Else
    IDClassLottoProdPerFuoriQuota = 0
    MsgInDocSeRigaMerceSenzaImballo = 0
    AttivaMappaturaDaRegFatture = 0
    DBNameRegFatture = ""
End If

rs.CloseResultset
Set rs = Nothing
End Sub

