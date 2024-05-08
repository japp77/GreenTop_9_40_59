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
Object = "{F95AA20B-3F80-11D3-A741-00105A2E9BAF}#2.1#0"; "DmtSearchAccount2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   11145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17880
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
   ScaleHeight     =   11145
   ScaleWidth      =   17880
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   10800
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Width           =   17880
      _LayoutVersion  =   2
      _ExtentX        =   31538
      _ExtentY        =   19050
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
      Begin DMTPrinterDialog.DMTDialog DmtPrnDlg 
         Left            =   120
         Top             =   6240
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   1320
         TabIndex        =   79
         Top             =   600
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
      End
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H8000000A&
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
         Height          =   4935
         Left            =   840
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   157
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.PictureBox PicForm 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   10215
         Left            =   0
         ScaleHeight     =   10185
         ScaleWidth      =   17745
         TabIndex        =   80
         Top             =   0
         Width           =   17775
         Begin VB.PictureBox PicForm2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   9975
            Left            =   120
            ScaleHeight     =   9945
            ScaleWidth      =   17505
            TabIndex        =   81
            Top             =   120
            Width           =   17535
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
               Left            =   120
               TabIndex        =   220
               Top             =   8880
               Width           =   15855
               Begin DMTEDITNUMLib.dmtCurrency curTotImponibile 
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   221
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   503
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
                  CurrencySymbol  =   "€"
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curTotImposta 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   222
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   503
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
                  CurrencySymbol  =   "€"
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curTotDocumento 
                  Height          =   285
                  Left            =   5160
                  TabIndex        =   223
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   503
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
                  CurrencySymbol  =   "€"
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curTotArrotondamenti 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   224
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   503
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
                  CurrencySymbol  =   "€"
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curNettoAPagare 
                  Height          =   285
                  Left            =   6840
                  TabIndex        =   225
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   503
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
                  CurrencySymbol  =   "€"
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency curNettoAPagare_naz 
                  Height          =   285
                  Left            =   8520
                  TabIndex        =   226
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1605
                  _Version        =   65536
                  _ExtentX        =   2831
                  _ExtentY        =   503
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
                  CurrencySymbol  =   "€"
                  DecFinalZeros   =   -1  'True
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale imponibile"
                  Height          =   195
                  Index           =   19
                  Left            =   3480
                  TabIndex        =   232
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale imposta"
                  Height          =   195
                  Index           =   20
                  Left            =   1800
                  TabIndex        =   231
                  Top             =   240
                  Width           =   1050
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale documento"
                  Height          =   195
                  Index           =   21
                  Left            =   5160
                  TabIndex        =   230
                  Top             =   240
                  Width           =   1530
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Totale arrotondamenti"
                  Height          =   195
                  Index           =   22
                  Left            =   120
                  TabIndex        =   229
                  Top             =   240
                  Width           =   1605
               End
               Begin VB.Label lblDocument 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Netto a pagare corr."
                  Height          =   195
                  Index           =   23
                  Left            =   6840
                  TabIndex        =   228
                  Top             =   240
                  Width           =   1485
               End
               Begin VB.Label Label4 
                  Caption         =   "Netto a pagare naz."
                  Height          =   255
                  Index           =   4
                  Left            =   8520
                  TabIndex        =   227
                  Top             =   240
                  Width           =   1575
               End
            End
            Begin TabDlg.SSTab SSTab1 
               Height          =   8895
               Left            =   120
               TabIndex        =   82
               Top             =   0
               Width           =   15855
               _ExtentX        =   27966
               _ExtentY        =   15690
               _Version        =   393216
               TabHeight       =   520
               TabCaption(0)   =   "Documento (F8)"
               TabPicture(0)   =   "frmMain.frx":479EA
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "FraTab(0)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "FraTab(1)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "FraTab(2)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "FraTab(4)"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "FraTab(5)"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "FraTab(7)"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "Frame3"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).ControlCount=   7
               TabCaption(1)   =   "Corpo (F9)"
               TabPicture(1)   =   "frmMain.frx":47A06
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "lblInfoTesta"
               Tab(1).Control(1)=   "FraTab(3)"
               Tab(1).Control(2)=   "Frame4"
               Tab(1).Control(3)=   "Frame5"
               Tab(1).ControlCount=   4
               TabCaption(2)   =   "Commissioni"
               TabPicture(2)   =   "frmMain.frx":47A22
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Frame7"
               Tab(2).ControlCount=   1
               Begin VB.Frame Frame3 
                  Caption         =   "Intrastat"
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
                  Height          =   975
                  Left            =   120
                  TabIndex        =   83
                  Top             =   6840
                  Width           =   15645
                  Begin VB.ComboBox cboIntraSezDaComp 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   36
                     Top             =   360
                     Width           =   1575
                  End
                  Begin VB.ComboBox cboIntraTrimestre 
                     Height          =   315
                     Left            =   5400
                     TabIndex        =   38
                     Top             =   360
                     Width           =   735
                  End
                  Begin DMTDataCmb.DMTCombo cboIntraProvincia 
                     Height          =   315
                     Left            =   8760
                     TabIndex        =   41
                     Top             =   360
                     Width           =   1335
                     _ExtentX        =   2355
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
                  Begin DMTDataCmb.DMTCombo cboIntraTrasporto 
                     Height          =   315
                     Left            =   6840
                     TabIndex        =   40
                     Top             =   360
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
                  Begin DMTEDITNUMLib.dmtNumber txtIntraAnno 
                     Height          =   315
                     Left            =   6120
                     TabIndex        =   39
                     Top             =   360
                     Width           =   615
                     _Version        =   65536
                     _ExtentX        =   1085
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboIntraMese 
                     Height          =   315
                     Left            =   4320
                     TabIndex        =   37
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
                  Begin DMTDataCmb.DMTCombo cboIntraDatiRichiesti 
                     Height          =   315
                     Left            =   1080
                     TabIndex        =   35
                     Top             =   360
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
                  Begin VB.CheckBox chkCessione 
                     Caption         =   "Cessione"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   34
                     Top             =   360
                     Width           =   975
                  End
                  Begin DMTDataCmb.DMTCombo cboIntraNazione 
                     Height          =   315
                     Left            =   10200
                     TabIndex        =   165
                     TabStop         =   0   'False
                     Top             =   360
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
                  Begin VB.Label Label8 
                     Caption         =   "Nazione di pagamento"
                     Height          =   255
                     Index           =   7
                     Left            =   10200
                     TabIndex        =   166
                     Top             =   160
                     Width           =   2415
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Dati richiesti"
                     Height          =   255
                     Index           =   0
                     Left            =   1080
                     TabIndex        =   90
                     Top             =   160
                     Width           =   1455
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Sezione da compilare"
                     Height          =   255
                     Index           =   1
                     Left            =   2640
                     TabIndex        =   89
                     Top             =   160
                     Width           =   1575
                  End
                  Begin VB.Label Label8 
                     Alignment       =   2  'Center
                     Caption         =   "Mese o trimestre"
                     Height          =   255
                     Index           =   2
                     Left            =   4320
                     TabIndex        =   88
                     Top             =   160
                     Width           =   1575
                  End
                  Begin VB.Label Label8 
                     Height          =   255
                     Index           =   3
                     Left            =   8160
                     TabIndex        =   87
                     Top             =   120
                     Width           =   255
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Anno"
                     Height          =   255
                     Index           =   4
                     Left            =   6240
                     TabIndex        =   86
                     Top             =   160
                     Width           =   495
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Trasporto"
                     Height          =   255
                     Index           =   5
                     Left            =   6840
                     TabIndex        =   85
                     Top             =   160
                     Width           =   1815
                  End
                  Begin VB.Label Label8 
                     Caption         =   "Provincia"
                     Height          =   255
                     Index           =   6
                     Left            =   8760
                     TabIndex        =   84
                     Top             =   160
                     Width           =   1335
                  End
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Intrastat imballo"
                  ForeColor       =   &H00800000&
                  Height          =   855
                  Left            =   -74880
                  TabIndex        =   118
                  Top             =   7920
                  Visible         =   0   'False
                  Width           =   15615
                  Begin VB.CheckBox chkRiportoIntra_Imb 
                     Caption         =   "Non riporto"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   66
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   1215
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIntra_Imb_MassaNetta 
                     Height          =   285
                     Left            =   5640
                     TabIndex        =   68
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDIntra_Imb_Nom_Comb 
                     Height          =   855
                     Left            =   1440
                     TabIndex        =   67
                     TabStop         =   0   'False
                     Top             =   120
                     Width           =   4215
                     _ExtentX        =   7435
                     _ExtentY        =   1508
                     PropCodice      =   $"frmMain.frx":47A3E
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":47A99
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":47AFC
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
                  Begin DmtCodDescCtl.DmtCodDesc CDIntra_Imb_Nat_Trans 
                     Height          =   855
                     Left            =   6720
                     TabIndex        =   69
                     TabStop         =   0   'False
                     Top             =   120
                     Width           =   4215
                     _ExtentX        =   7435
                     _ExtentY        =   1508
                     PropCodice      =   $"frmMain.frx":47B56
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":47BA5
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":47C03
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
                  Begin VB.Label Label10 
                     Caption         =   "Massa netta"
                     Height          =   255
                     Left            =   5640
                     TabIndex        =   119
                     Top             =   120
                     Width           =   975
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "Instrastat articolo"
                  ForeColor       =   &H00800000&
                  Height          =   855
                  Left            =   -74880
                  TabIndex        =   120
                  Top             =   7200
                  Width           =   15615
                  Begin VB.CheckBox chkRiportoIntra_Art 
                     Caption         =   "Non riporto"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   62
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   1215
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIntra_Art_MassaNetta 
                     Height          =   285
                     Left            =   5640
                     TabIndex        =   64
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDIntra_art_Nom_Comb 
                     Height          =   855
                     Left            =   1440
                     TabIndex        =   63
                     TabStop         =   0   'False
                     Top             =   120
                     Width           =   4215
                     _ExtentX        =   7435
                     _ExtentY        =   1508
                     PropCodice      =   $"frmMain.frx":47C5D
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":47CB8
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":47D1B
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
                  Begin DmtCodDescCtl.DmtCodDesc CDIntra_art_Nat_Trans 
                     Height          =   855
                     Left            =   6720
                     TabIndex        =   65
                     TabStop         =   0   'False
                     Top             =   120
                     Width           =   4215
                     _ExtentX        =   7435
                     _ExtentY        =   1508
                     PropCodice      =   $"frmMain.frx":47D75
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":47DC4
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":47E22
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
                  Begin VB.Label Label9 
                     Caption         =   "Massa netta"
                     Height          =   255
                     Left            =   5640
                     TabIndex        =   121
                     Top             =   120
                     Width           =   975
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
                  Height          =   1500
                  Index           =   7
                  Left            =   120
                  TabIndex        =   91
                  Top             =   5520
                  Width           =   15645
                  Begin VB.TextBox txtAnnotazioni 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   33
                     Top             =   1080
                     Width           =   6375
                  End
                  Begin DMTDATETIMELib.dmtTime txtOraTrasporto 
                     Height          =   315
                     Left            =   11400
                     TabIndex        =   32
                     Top             =   480
                     Width           =   855
                     _Version        =   65536
                     _ExtentX        =   1508
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataTrasporto 
                     Height          =   315
                     Left            =   10200
                     TabIndex        =   31
                     Top             =   480
                     Width           =   1095
                     _Version        =   65536
                     _ExtentX        =   1931
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboAspettoEsteriore 
                     Height          =   315
                     Left            =   8040
                     TabIndex        =   30
                     Top             =   480
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
                  Begin DMTEDITNUMLib.dmtNumber txtPesoTotale 
                     Height          =   315
                     Left            =   7200
                     TabIndex        =   29
                     Top             =   480
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtColliTotali 
                     Height          =   315
                     Left            =   6360
                     TabIndex        =   28
                     Top             =   480
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboVettore 
                     Height          =   315
                     Left            =   4200
                     TabIndex        =   27
                     Top             =   480
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
                  Begin DMTDataCmb.DMTCombo cboTrasporto 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   26
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
                     TabIndex        =   25
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
                  Begin DmtCodDescCtl.DmtCodDesc CDAgenteTesta 
                     Height          =   585
                     Left            =   6600
                     TabIndex        =   167
                     Top             =   860
                     Width           =   5745
                     _ExtentX        =   10134
                     _ExtentY        =   1032
                     PropCodice      =   $"frmMain.frx":47E7C
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":47ECD
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":47F1E
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
                  Begin VB.Label Label1 
                     Caption         =   "Porto"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   92
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Annotazioni"
                     Height          =   255
                     Index           =   8
                     Left            =   120
                     TabIndex        =   100
                     Top             =   840
                     Width           =   7815
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Ora trasp."
                     Height          =   255
                     Index           =   7
                     Left            =   11400
                     TabIndex        =   99
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Data trasp."
                     Height          =   255
                     Index           =   6
                     Left            =   10200
                     TabIndex        =   98
                     Top             =   240
                     Width           =   975
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Aspetto esteriore"
                     Height          =   255
                     Index           =   5
                     Left            =   8040
                     TabIndex        =   97
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Peso"
                     Height          =   255
                     Index           =   4
                     Left            =   7200
                     TabIndex        =   96
                     Top             =   240
                     Width           =   735
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Colli"
                     Height          =   255
                     Index           =   3
                     Left            =   6360
                     TabIndex        =   95
                     Top             =   240
                     Width           =   735
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Vettore"
                     Height          =   255
                     Index           =   2
                     Left            =   4200
                     TabIndex        =   94
                     Top             =   240
                     Width           =   2055
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Trasporto"
                     Height          =   255
                     Index           =   1
                     Left            =   2160
                     TabIndex        =   93
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
                  TabIndex        =   105
                  Top             =   4200
                  Width           =   9045
                  Begin DMTEDITNUMLib.dmtCurrency curSpeseIncasso 
                     Height          =   285
                     Left            =   120
                     TabIndex        =   21
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
                     Left            =   120
                     TabIndex        =   23
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
                     Left            =   1200
                     TabIndex        =   22
                     Top             =   360
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
                     Left            =   1200
                     TabIndex        =   24
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
                  Begin VB.Line Line3 
                     X1              =   2160
                     X2              =   2160
                     Y1              =   120
                     Y2              =   1320
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sc.Imp"
                     Height          =   195
                     Index           =   18
                     Left            =   1200
                     TabIndex        =   109
                     Top             =   640
                     Width           =   495
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sc.%"
                     Height          =   195
                     Index           =   17
                     Left            =   1200
                     TabIndex        =   108
                     Top             =   180
                     Width           =   390
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Trasporto"
                     Height          =   195
                     Index           =   16
                     Left            =   120
                     TabIndex        =   107
                     Top             =   640
                     Width           =   705
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Incasso"
                     Height          =   195
                     Index           =   15
                     Left            =   120
                     TabIndex        =   106
                     Top             =   180
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
                  TabIndex        =   103
                  Top             =   4200
                  Width           =   2940
                  Begin MSComctlLib.ListView lvwScadenze 
                     Height          =   1050
                     Left            =   120
                     TabIndex        =   104
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
                  TabIndex        =   101
                  Top             =   4200
                  Width           =   3810
                  Begin MSComctlLib.ListView lvwIVA 
                     Height          =   1050
                     Left            =   120
                     TabIndex        =   102
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
                  Height          =   6525
                  Index           =   3
                  Left            =   -74880
                  TabIndex        =   139
                  Top             =   720
                  Width           =   15615
                  Begin VB.Frame FraValoriOriginali 
                     Caption         =   "Valori originali"
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
                     Height          =   2055
                     Left            =   10440
                     TabIndex        =   243
                     Top             =   720
                     Width           =   5055
                     Begin DMTDATETIMELib.dmtDate txtDataLavorazione 
                        Height          =   315
                        Left            =   3360
                        TabIndex        =   244
                        Top             =   960
                        Width           =   1575
                        _Version        =   65536
                        _ExtentX        =   2778
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
                     Begin DMTEDITNUMLib.dmtNumber txtQuantitaOriginale 
                        Height          =   315
                        Left            =   1680
                        TabIndex        =   245
                        Top             =   960
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
                        DecimalPlaces   =   3
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPrezzoOriginale 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   246
                        Top             =   1515
                        Width           =   1455
                        _Version        =   65536
                        _ExtentX        =   2566
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
                        DecimalPlaces   =   5
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtIDTipoVariazione 
                        Height          =   375
                        Left            =   360
                        TabIndex        =   247
                        Top             =   2880
                        Visible         =   0   'False
                        Width           =   1335
                        _Version        =   65536
                        _ExtentX        =   2355
                        _ExtentY        =   661
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
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtColliOriginali 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   248
                        Top             =   435
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
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPesoLordoOriginale 
                        Height          =   315
                        Left            =   1200
                        TabIndex        =   249
                        Top             =   435
                        Width           =   1215
                        _Version        =   65536
                        _ExtentX        =   2143
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
                        DecimalPlaces   =   3
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPesoNettoOriginale 
                        Height          =   315
                        Left            =   3720
                        TabIndex        =   250
                        Top             =   435
                        Width           =   1215
                        _Version        =   65536
                        _ExtentX        =   2143
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
                        DecimalPlaces   =   3
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPezziOriginali 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   251
                        Top             =   960
                        Width           =   1455
                        _Version        =   65536
                        _ExtentX        =   2566
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
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtTaraOriginale 
                        Height          =   315
                        Left            =   2520
                        TabIndex        =   252
                        Top             =   435
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
                        DecimalPlaces   =   3
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtSconto2Ori 
                        Height          =   315
                        Left            =   2520
                        TabIndex        =   253
                        Top             =   1500
                        Width           =   735
                        _Version        =   65536
                        _ExtentX        =   1296
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
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtSconto1Ori 
                        Height          =   315
                        Left            =   1680
                        TabIndex        =   254
                        Top             =   1500
                        Width           =   735
                        _Version        =   65536
                        _ExtentX        =   1296
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
                        DecFinalZeros   =   -1  'True
                        AllowEmpty      =   0   'False
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Importo"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Index           =   1
                        Left            =   120
                        TabIndex        =   264
                        Top             =   1320
                        Width           =   975
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Q.tà Mov."
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Index           =   2
                        Left            =   1680
                        TabIndex        =   263
                        Top             =   760
                        Width           =   975
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Peso netto"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Index           =   3
                        Left            =   3720
                        TabIndex        =   262
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Tara"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Index           =   4
                        Left            =   2520
                        TabIndex        =   261
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Peso lordo"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Index           =   5
                        Left            =   1200
                        TabIndex        =   260
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Pezzi"
                        ForeColor       =   &H00FF0000&
                        Height          =   210
                        Index           =   6
                        Left            =   120
                        TabIndex        =   259
                        Top             =   760
                        Width           =   975
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Colli"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Index           =   7
                        Left            =   120
                        TabIndex        =   258
                        Top             =   240
                        Width           =   975
                     End
                     Begin VB.Label Label13 
                        Caption         =   "Data lavorazione"
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Index           =   9
                        Left            =   3360
                        TabIndex        =   257
                        Top             =   765
                        Width           =   1575
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "% Sc.1"
                        ForeColor       =   &H00FF0000&
                        Height          =   195
                        Index           =   9
                        Left            =   1680
                        TabIndex        =   256
                        Top             =   1320
                        Width           =   525
                     End
                     Begin VB.Label lblDocument 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "% Sc. 2"
                        ForeColor       =   &H00FF0000&
                        Height          =   195
                        Index           =   11
                        Left            =   2520
                        TabIndex        =   255
                        Top             =   1320
                        Width           =   570
                     End
                  End
                  Begin VB.CommandButton cmdElencoProdottiMix 
                     Height          =   285
                     Left            =   9960
                     Picture         =   "frmMain.frx":47F78
                     Style           =   1  'Graphical
                     TabIndex        =   241
                     ToolTipText     =   "Composizione articolo Mix"
                     Top             =   825
                     Width           =   375
                  End
                  Begin VB.CheckBox chkRiscontroPeso 
                     Caption         =   "Riscontro peso"
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Left            =   12000
                     TabIndex        =   240
                     Top             =   2880
                     Width           =   2175
                  End
                  Begin VB.CommandButton cmdPianoDeiContiRiga 
                     Height          =   315
                     Left            =   14880
                     Picture         =   "frmMain.frx":48502
                     Style           =   1  'Graphical
                     TabIndex        =   211
                     ToolTipText     =   "Piano dei conti della merce e degli imballi"
                     Top             =   3480
                     Width           =   375
                  End
                  Begin MSComctlLib.ListView lvwArticoli 
                     Height          =   3015
                     Left            =   120
                     TabIndex        =   140
                     Top             =   3360
                     Width           =   14205
                     _ExtentX        =   25056
                     _ExtentY        =   5318
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
                  Begin VB.TextBox txtNomeSocioFatt 
                     BackColor       =   &H8000000F&
                     Height          =   340
                     Left            =   12840
                     Locked          =   -1  'True
                     TabIndex        =   200
                     TabStop         =   0   'False
                     Top             =   340
                     Width           =   1335
                  End
                  Begin VB.CommandButton cmdAgenteRiga 
                     Height          =   315
                     Left            =   14400
                     Picture         =   "frmMain.frx":48A8C
                     Style           =   1  'Graphical
                     TabIndex        =   176
                     ToolTipText     =   "Dati agente"
                     Top             =   3480
                     Width           =   375
                  End
                  Begin VB.Frame FraAgenteRiga 
                     Height          =   735
                     Left            =   120
                     TabIndex        =   168
                     Top             =   3360
                     Visible         =   0   'False
                     Width           =   12855
                     Begin DMTEDITNUMLib.dmtNumber txtImportoProvv 
                        Height          =   315
                        Left            =   11760
                        TabIndex        =   169
                        Top             =   315
                        Width           =   975
                        _Version        =   65536
                        _ExtentX        =   1720
                        _ExtentY        =   556
                        _StockProps     =   253
                        Text            =   "0"
                        BackColor       =   16777215
                        Appearance      =   1
                        AllowEmpty      =   0   'False
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtPercProvv 
                        Height          =   315
                        Left            =   10800
                        TabIndex        =   170
                        Top             =   315
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
                     Begin DMTDataCmb.DMTCombo cboRegolaProvvigione 
                        Height          =   315
                        Left            =   5880
                        TabIndex        =   171
                        Top             =   315
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
                     Begin DmtCodDescCtl.DmtCodDesc CDAgenteRiga 
                        Height          =   585
                        Left            =   120
                        TabIndex        =   172
                        Top             =   120
                        Width           =   5745
                        _ExtentX        =   10134
                        _ExtentY        =   1032
                        PropCodice      =   $"frmMain.frx":49016
                        BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        PropDescrizione =   $"frmMain.frx":49067
                        BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   -1  'True
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        MenuFunctions   =   $"frmMain.frx":490B8
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
                     Begin DMTDataCmb.DMTCombo cboTipoOrdine 
                        Height          =   315
                        Left            =   8760
                        TabIndex        =   218
                        Top             =   315
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
                     Begin VB.Label Label12 
                        Caption         =   "Tipo ordine"
                        Height          =   255
                        Index           =   2
                        Left            =   8760
                        TabIndex        =   219
                        Top             =   120
                        Width           =   1815
                     End
                     Begin VB.Label Label17 
                        Caption         =   "Imp. Provv."
                        Height          =   255
                        Left            =   11760
                        TabIndex        =   175
                        Top             =   120
                        Width           =   855
                     End
                     Begin VB.Label Label16 
                        Caption         =   "% Provv."
                        Height          =   255
                        Left            =   10800
                        TabIndex        =   174
                        Top             =   120
                        Width           =   855
                     End
                     Begin VB.Label Label12 
                        Caption         =   "Regola di provvigione"
                        Height          =   255
                        Index           =   1
                        Left            =   5880
                        TabIndex        =   173
                        Top             =   120
                        Width           =   3015
                     End
                  End
                  Begin VB.CheckBox chkPrezzoMedioInLiq 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Prezzo medio in liquid."
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   8040
                     TabIndex        =   164
                     Top             =   1920
                     Width           =   2295
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQuantitaConferita 
                     Height          =   255
                     Left            =   11400
                     TabIndex        =   163
                     Top             =   2520
                     Visible         =   0   'False
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin VB.TextBox txtDescrizioneArticolo 
                     Height          =   315
                     Left            =   1560
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   47
                     Top             =   1340
                     Width           =   3015
                  End
                  Begin VB.CommandButton cmdElimina 
                     Caption         =   "Elimina"
                     Height          =   285
                     Left            =   14400
                     TabIndex        =   61
                     TabStop         =   0   'False
                     Top             =   6000
                     Width           =   1065
                  End
                  Begin VB.CommandButton cmdSalva 
                     Caption         =   "&Salva"
                     Height          =   285
                     Left            =   14400
                     MaskColor       =   &H00FF0000&
                     TabIndex        =   59
                     Top             =   5280
                     Width           =   1065
                  End
                  Begin VB.CommandButton cmdNuovo 
                     Caption         =   "&Nuovo"
                     Height          =   285
                     Left            =   14400
                     TabIndex        =   60
                     Top             =   4560
                     Width           =   1065
                  End
                  Begin VB.TextBox txtCodiceLottoVendita 
                     Height          =   315
                     Left            =   4680
                     TabIndex        =   48
                     TabStop         =   0   'False
                     Top             =   1340
                     Width           =   3855
                  End
                  Begin VB.TextBox txtSocio 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   4200
                     TabIndex        =   44
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   3135
                  End
                  Begin VB.TextBox txtCodiceSocio 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   3600
                     TabIndex        =   43
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   615
                  End
                  Begin VB.TextBox txtConferimentoRighe 
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   45
                     TabStop         =   0   'False
                     Top             =   825
                     Width           =   9855
                  End
                  Begin VB.TextBox txtDocumentoRiferimento 
                     Appearance      =   0  'Flat
                     Height          =   315
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   58
                     TabStop         =   0   'False
                     Top             =   2955
                     Width           =   6045
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtTotaleRiga 
                     Height          =   315
                     Left            =   4680
                     TabIndex        =   54
                     Top             =   2445
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   " 0"
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
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtImponibileUnitario 
                     Height          =   315
                     Left            =   4680
                     TabIndex        =   53
                     TabStop         =   0   'False
                     Top             =   2445
                     Visible         =   0   'False
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
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
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioArticolo 
                     Height          =   315
                     Left            =   1800
                     TabIndex        =   52
                     Top             =   2445
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecimalPlaces   =   4
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtAliquotaArticolo 
                     Height          =   315
                     Left            =   1080
                     TabIndex        =   51
                     TabStop         =   0   'False
                     Top             =   2445
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
                     Left            =   120
                     TabIndex        =   50
                     TabStop         =   0   'False
                     Top             =   2445
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
                     Left            =   6720
                     TabIndex        =   49
                     Top             =   1845
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
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
                  Begin DMTDataCmb.DMTCombo cboUnitaDiMisura 
                     Height          =   315
                     Left            =   8640
                     TabIndex        =   141
                     TabStop         =   0   'False
                     Top             =   1320
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
                  Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   46
                     Top             =   1080
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":49112
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4916A
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":491C1
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
                  Begin DMTDATETIMELib.dmtDate txtDataConferimento 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   42
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   1095
                     _Version        =   65536
                     _ExtentX        =   1931
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                  End
                  Begin DMTDataCmb.DMTCombo cboTipoVariazione 
                     Height          =   315
                     Left            =   6240
                     TabIndex        =   55
                     TabStop         =   0   'False
                     Top             =   2445
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
                  Begin DMTEDITNUMLib.dmtNumber txtPezzi 
                     Height          =   315
                     Left            =   5400
                     TabIndex        =   190
                     Top             =   1845
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
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
                     Left            =   2760
                     TabIndex        =   191
                     Top             =   1845
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
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
                  Begin DMTEDITNUMLib.dmtNumber txtPesoNetto 
                     Height          =   315
                     Left            =   4080
                     TabIndex        =   192
                     Top             =   1845
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
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
                  Begin DMTEDITNUMLib.dmtNumber txtPesoLordo 
                     Height          =   315
                     Left            =   1440
                     TabIndex        =   193
                     Top             =   1845
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
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
                  Begin DMTEDITNUMLib.dmtNumber txtColli 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   194
                     Top             =   1845
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDSocioFatt 
                     Height          =   615
                     Left            =   8880
                     TabIndex        =   201
                     Top             =   105
                     Width           =   3975
                     _ExtentX        =   7011
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4921B
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":49269
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":492CE
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
                  Begin DMTDataCmb.DMTCombo cboTipoImportoLiq 
                     Height          =   315
                     Left            =   6240
                     TabIndex        =   56
                     TabStop         =   0   'False
                     Top             =   2955
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
                  Begin DMTDataCmb.DMTCombo cboTipoDocumentoCoop 
                     Height          =   315
                     Left            =   1320
                     TabIndex        =   204
                     TabStop         =   0   'False
                     Top             =   360
                     Width           =   2175
                     _ExtentX        =   3836
                     _ExtentY        =   556
                     Enabled         =   0   'False
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
                  Begin DMTEDITNUMLib.dmtCurrency txtImportoLiqVarMan 
                     Height          =   315
                     Left            =   10440
                     TabIndex        =   57
                     TabStop         =   0   'False
                     Top             =   2955
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
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
                  Begin VB.Frame FraPDCRiga 
                     Height          =   735
                     Left            =   120
                     TabIndex        =   206
                     Top             =   3360
                     Visible         =   0   'False
                     Width           =   12855
                     Begin VB.TextBox txtCodiceConto 
                        Height          =   315
                        Left            =   960
                        TabIndex        =   208
                        Top             =   360
                        Width           =   1815
                     End
                     Begin VB.TextBox txtDescrizioneConto 
                        Height          =   315
                        Left            =   2760
                        TabIndex        =   207
                        Top             =   360
                        Width           =   3975
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtIDPianoDeiConti 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   209
                        TabStop         =   0   'False
                        Top             =   360
                        Width           =   855
                        _Version        =   65536
                        _ExtentX        =   1508
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
                        AllowEmpty      =   0   'False
                     End
                     Begin VB.Label lblPianodeiDeiConti 
                        Caption         =   "Piano dei conti articolo"
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
                        Height          =   255
                        Left            =   120
                        TabIndex        =   210
                        Top             =   120
                        Width           =   6615
                     End
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtSconto2 
                     Height          =   315
                     Left            =   3960
                     TabIndex        =   265
                     Top             =   2445
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
                     Left            =   3240
                     TabIndex        =   266
                     Top             =   2445
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
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "% Sc. 1"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   8
                     Left            =   3240
                     TabIndex        =   268
                     Top             =   2250
                     Width           =   570
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "% Sc. 2"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   32
                     Left            =   3960
                     TabIndex        =   267
                     Top             =   2250
                     Width           =   570
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Var. Liq. Manuale"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   45
                     Left            =   10440
                     TabIndex        =   212
                     ToolTipText     =   "Variazione importo liquidazione manuale"
                     Top             =   2760
                     Width           =   1245
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tipo docucumento"
                     Height          =   255
                     Index           =   10
                     Left            =   1320
                     TabIndex        =   205
                     Top             =   165
                     Width           =   1935
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Forza tipo importo in liquidazione"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   43
                     Left            =   6240
                     TabIndex        =   203
                     Top             =   2760
                     Width           =   4020
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Riferimento del documento"
                     Height          =   255
                     Index           =   9
                     Left            =   120
                     TabIndex        =   202
                     Top             =   2760
                     Width           =   3015
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Colli"
                     Height          =   255
                     Index           =   6
                     Left            =   120
                     TabIndex        =   199
                     Top             =   1650
                     Width           =   735
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Peso lordo"
                     Height          =   255
                     Index           =   1
                     Left            =   1440
                     TabIndex        =   198
                     Top             =   1650
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Pezzi"
                     Height          =   255
                     Index           =   2
                     Left            =   5400
                     TabIndex        =   197
                     Top             =   1650
                     Width           =   615
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara "
                     Height          =   255
                     Index           =   3
                     Left            =   2760
                     TabIndex        =   196
                     Top             =   1650
                     Width           =   615
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Peso netto"
                     Height          =   255
                     Index           =   4
                     Left            =   4080
                     TabIndex        =   195
                     Top             =   1650
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tipo variazione"
                     Height          =   255
                     Index           =   0
                     Left            =   6240
                     TabIndex        =   177
                     Top             =   2250
                     Width           =   3015
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Q.tà mov."
                     Height          =   195
                     Index           =   5
                     Left            =   6720
                     TabIndex        =   154
                     Top             =   1650
                     Width           =   855
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Imp. Unit. Art."
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   31
                     Left            =   1800
                     TabIndex        =   153
                     Top             =   2250
                     Width           =   1125
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "U. M."
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   28
                     Left            =   8640
                     TabIndex        =   152
                     Top             =   1110
                     Width           =   975
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "% Aliq."
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   27
                     Left            =   1080
                     TabIndex        =   151
                     Top             =   2250
                     Width           =   525
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Cod. IVA"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   26
                     Left            =   120
                     TabIndex        =   150
                     Top             =   2250
                     Width           =   1005
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
                     MouseIcon       =   "frmMain.frx":49328
                     MousePointer    =   99  'Custom
                     TabIndex        =   149
                     Top             =   1110
                     Width           =   2700
                  End
                  Begin VB.Label Label13 
                     Caption         =   "Codice lotto"
                     Height          =   255
                     Index           =   0
                     Left            =   4680
                     TabIndex        =   148
                     Top             =   1110
                     Width           =   1095
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Imponibile riga"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Index           =   33
                     Left            =   4680
                     TabIndex        =   147
                     Top             =   2250
                     Visible         =   0   'False
                     Width           =   1035
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Totale riga"
                     Height          =   255
                     Index           =   8
                     Left            =   4680
                     TabIndex        =   146
                     Top             =   2250
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Data conf."
                     Height          =   255
                     Index           =   7
                     Left            =   120
                     TabIndex        =   145
                     Top             =   160
                     Width           =   1095
                  End
                  Begin VB.Label lblCodiceSocio 
                     Caption         =   "Codice"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   3600
                     MouseIcon       =   "frmMain.frx":49632
                     MousePointer    =   99  'Custom
                     TabIndex        =   144
                     Top             =   165
                     Width           =   495
                  End
                  Begin VB.Label lblNomeSocio 
                     BorderStyle     =   1  'Fixed Single
                     Height          =   315
                     Left            =   7320
                     TabIndex        =   143
                     Top             =   360
                     Width           =   1455
                  End
                  Begin VB.Label lblSocio 
                     Caption         =   "Socio/Fornitore"
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   4200
                     MouseIcon       =   "frmMain.frx":4993C
                     MousePointer    =   99  'Custom
                     TabIndex        =   142
                     Top             =   165
                     Width           =   3135
                  End
                  Begin VB.Label lblCollegamentoConferimento 
                     Caption         =   "Collegamento al conferimento "
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   120
                     MouseIcon       =   "frmMain.frx":49C46
                     MousePointer    =   99  'Custom
                     TabIndex        =   155
                     Top             =   640
                     Width           =   7215
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
                  Height          =   3855
                  Index           =   1
                  Left            =   4560
                  TabIndex        =   130
                  Top             =   540
                  Width           =   11205
                  Begin VB.Frame fraAltriDati 
                     Caption         =   "Riferimento Ordine Cliente"
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
                     Height          =   975
                     Left            =   5880
                     TabIndex        =   269
                     Top             =   2400
                     Width           =   5205
                     Begin VB.TextBox txtNumeroOrdineCliente 
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
                        Height          =   315
                        Left            =   1680
                        TabIndex        =   10
                        Top             =   360
                        Width           =   3465
                     End
                     Begin DMTDATETIMELib.dmtDate txtDataOrdineCliente 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   9
                        Top             =   360
                        Width           =   1455
                        _Version        =   65536
                        _ExtentX        =   2566
                        _ExtentY        =   556
                        _StockProps     =   253
                        ForeColor       =   16711680
                        BackColor       =   -2147483643
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Tahoma"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Appearance      =   1
                     End
                     Begin VB.Label Label4 
                        Caption         =   "Data "
                        Height          =   255
                        Index           =   5
                        Left            =   120
                        TabIndex        =   271
                        Top             =   170
                        Width           =   1095
                     End
                     Begin VB.Label Label5 
                        Caption         =   "Numero"
                        Height          =   255
                        Index           =   3
                        Left            =   1680
                        TabIndex        =   270
                        Top             =   165
                        Width           =   2655
                     End
                  End
                  Begin VB.TextBox txtNLetteraIntento 
                     Height          =   315
                     Left            =   840
                     Locked          =   -1  'True
                     TabIndex        =   242
                     Top             =   2040
                     Width           =   495
                  End
                  Begin DmtCodDescCtl.DmtCodDesc cdAnagrafica 
                     Height          =   585
                     Left            =   120
                     TabIndex        =   4
                     Top             =   720
                     Visible         =   0   'False
                     Width           =   5625
                     _ExtentX        =   9922
                     _ExtentY        =   1032
                     PropCodice      =   $"frmMain.frx":49F50
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":49FB4
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4A005
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
                     Picture         =   "frmMain.frx":4A05F
                     Style           =   1  'Graphical
                     TabIndex        =   216
                     ToolTipText     =   "Lettere di intento del cliente"
                     Top             =   2040
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdEliminaRifLetInt 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":4A5E9
                     Style           =   1  'Graphical
                     TabIndex        =   213
                     ToolTipText     =   "Elimina riferimento lettera intento"
                     Top             =   2040
                     Width           =   375
                  End
                  Begin VB.CommandButton cmdCambioValuta 
                     Height          =   280
                     Left            =   10680
                     Picture         =   "frmMain.frx":4AB73
                     Style           =   1  'Graphical
                     TabIndex        =   189
                     ToolTipText     =   "Trova il cambio valuta"
                     Top             =   1920
                     Width           =   320
                  End
                  Begin VB.TextBox txtProvinciaAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   9960
                     Locked          =   -1  'True
                     TabIndex        =   182
                     TabStop         =   0   'False
                     Top             =   1380
                     Width           =   615
                  End
                  Begin VB.TextBox txtComuneAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   6840
                     Locked          =   -1  'True
                     TabIndex        =   181
                     TabStop         =   0   'False
                     Top             =   1380
                     Width           =   3135
                  End
                  Begin VB.TextBox txtCapAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   6000
                     Locked          =   -1  'True
                     TabIndex        =   180
                     TabStop         =   0   'False
                     Top             =   1380
                     Width           =   855
                  End
                  Begin VB.TextBox txtIndirizzoAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   6000
                     Locked          =   -1  'True
                     TabIndex        =   179
                     TabStop         =   0   'False
                     Top             =   1110
                     Width           =   4575
                  End
                  Begin VB.TextBox txtReferenteAltroSito 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   6000
                     Locked          =   -1  'True
                     TabIndex        =   178
                     TabStop         =   0   'False
                     Top             =   840
                     Width           =   4575
                  End
                  Begin VB.TextBox txtProvincia 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   3870
                     Locked          =   -1  'True
                     TabIndex        =   135
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   1855
                  End
                  Begin VB.TextBox txtComune 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1125
                     Locked          =   -1  'True
                     TabIndex        =   134
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   2730
                  End
                  Begin VB.TextBox txtCAP 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   133
                     TabStop         =   0   'False
                     Top             =   1080
                     Width           =   1005
                  End
                  Begin VB.TextBox txtIndirizzo 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   1560
                     Locked          =   -1  'True
                     TabIndex        =   132
                     TabStop         =   0   'False
                     Top             =   800
                     Width           =   4170
                  End
                  Begin VB.TextBox txtPartitaIva 
                     BackColor       =   &H8000000F&
                     Height          =   285
                     Left            =   120
                     Locked          =   -1  'True
                     TabIndex        =   131
                     TabStop         =   0   'False
                     Top             =   800
                     Width           =   1425
                  End
                  Begin DMTDataCmb.DMTCombo cboIvaCliente 
                     Height          =   315
                     Left            =   2520
                     TabIndex        =   12
                     TabStop         =   0   'False
                     Top             =   2040
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
                  Begin DMTDataCmb.DMTCombo cboBancaCliente 
                     Height          =   315
                     Left            =   3120
                     TabIndex        =   11
                     TabStop         =   0   'False
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
                     TabIndex        =   5
                     TabStop         =   0   'False
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
                     Left            =   4440
                     TabIndex        =   159
                     Top             =   2040
                     Width           =   1335
                     _ExtentX        =   2355
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
                  Begin DMTDataCmb.DMTCombo cboAltroSito 
                     Height          =   315
                     Left            =   6000
                     TabIndex        =   8
                     Top             =   480
                     Width           =   4575
                     _ExtentX        =   8070
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
                  Begin DMTDATETIMELib.dmtDate txtDataCambio 
                     Height          =   285
                     Left            =   7080
                     TabIndex        =   184
                     Top             =   1920
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   503
                     _StockProps     =   253
                     BackColor       =   8454143
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
                  Begin DMTEDITNUMLib.dmtNumber txtValoreCambioValuta 
                     Height          =   285
                     Left            =   9480
                     TabIndex        =   185
                     Top             =   1920
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   503
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   8454143
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
                     DecimalPlaces   =   5
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboCambioValuta 
                     Height          =   315
                     Left            =   8400
                     TabIndex        =   186
                     Top             =   2280
                     Visible         =   0   'False
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
                  Begin DMTEDITNUMLib.dmtNumber txtIDLetteraIntento 
                     Height          =   255
                     Left            =   120
                     TabIndex        =   214
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
                     TabIndex        =   215
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
                  Begin DmtSearchAccount2.DmtSearchACS2 ACSCliente 
                     Height          =   600
                     Left            =   120
                     TabIndex        =   3
                     Top             =   160
                     Width           =   5565
                     _ExtentX        =   9816
                     _ExtentY        =   1058
                     WidthCode       =   700
                     WidthDescription=   3450
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
                     CaptionDescription=   "Cognome o ragione sociale"
                     CaptionCode     =   "Codice"
                     OnlyAccounts    =   -1  'True
                  End
                  Begin DmtSearchAccount.DmtSearchACS ACSAnaDest 
                     Height          =   585
                     Left            =   120
                     TabIndex        =   7
                     Top             =   3000
                     Width           =   5670
                     _ExtentX        =   10001
                     _ExtentY        =   1032
                     WidthCode       =   700
                     WidthDescription=   3700
                     WidthSecondDescription=   1150
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
                  Begin DmtSearchAccount.DmtSearchACS ACSSocioTesta 
                     Height          =   585
                     Left            =   120
                     TabIndex        =   6
                     Top             =   2400
                     Width           =   5625
                     _ExtentX        =   9922
                     _ExtentY        =   1032
                     WidthCode       =   700
                     WidthDescription=   3700
                     WidthSecondDescription=   1100
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
                     CaptionDescription=   "Cedente"
                     CaptionCode     =   "Codice"
                     OnlyAccounts    =   -1  'True
                  End
                  Begin VB.Label lblLetteraIntento 
                     Caption         =   "Lettera d'intento"
                     Height          =   255
                     Left            =   840
                     TabIndex        =   217
                     Top             =   1860
                     Width           =   1575
                  End
                  Begin VB.Label Label18 
                     Caption         =   "Valore cambio"
                     Height          =   255
                     Left            =   8400
                     TabIndex        =   188
                     Top             =   1920
                     Width           =   1215
                  End
                  Begin VB.Label Label19 
                     Caption         =   "Data cambio"
                     Height          =   255
                     Left            =   6000
                     TabIndex        =   187
                     Top             =   1920
                     Width           =   1215
                  End
                  Begin VB.Line Line2 
                     X1              =   6000
                     X2              =   11160
                     Y1              =   1800
                     Y2              =   1800
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Destinazione diversa"
                     Height          =   195
                     Index           =   4
                     Left            =   6000
                     TabIndex        =   183
                     Top             =   240
                     Width           =   3525
                  End
                  Begin VB.Line Line1 
                     X1              =   5880
                     X2              =   5880
                     Y1              =   240
                     Y2              =   3240
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Valuta"
                     Height          =   255
                     Index           =   3
                     Left            =   4440
                     TabIndex        =   160
                     Top             =   1860
                     Width           =   1335
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Pagamento"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   138
                     Top             =   1350
                     Width           =   2775
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Banca"
                     Height          =   255
                     Index           =   1
                     Left            =   3120
                     TabIndex        =   137
                     Top             =   1350
                     Width           =   2655
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Esenzione I.V.A."
                     Height          =   255
                     Index           =   2
                     Left            =   2520
                     TabIndex        =   136
                     Top             =   1860
                     Width           =   1815
                  End
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
                  Height          =   3735
                  Index           =   0
                  Left            =   120
                  TabIndex        =   122
                  Top             =   540
                  Width           =   4455
                  Begin VB.TextBox txtCausaleDocumentoEF 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   238
                     TabStop         =   0   'False
                     Top             =   2340
                     Width           =   2175
                  End
                  Begin VB.CheckBox chkAddebitaBollo 
                     Caption         =   "Addebita bollo fatture esenti I.V.A."
                     Height          =   195
                     Left            =   120
                     TabIndex        =   237
                     TabStop         =   0   'False
                     Top             =   3000
                     Width           =   3735
                  End
                  Begin VB.CheckBox chkNomIva 
                     Caption         =   "Permetti indicazione cod. IVA differente ..."
                     Height          =   195
                     Left            =   120
                     TabIndex        =   236
                     TabStop         =   0   'False
                     ToolTipText     =   "Permetti indicazione cod. IVA differente per alcune vodi del doc."
                     Top             =   2760
                     Width           =   4215
                  End
                  Begin VB.CheckBox chkLordoIVA 
                     Caption         =   "Prezzi lordo IVA"
                     Height          =   240
                     Left            =   120
                     TabIndex        =   17
                     TabStop         =   0   'False
                     Top             =   1845
                     Width           =   1905
                  End
                  Begin VB.CheckBox chkRaggruppBolle 
                     Caption         =   "Raggruppa bolle"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   19
                     TabStop         =   0   'False
                     Top             =   2280
                     Width           =   2055
                  End
                  Begin VB.CheckBox chkRaggruppaScadenze 
                     Caption         =   "Raggruppa scadenze"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   20
                     TabStop         =   0   'False
                     Top             =   2520
                     Width           =   1935
                  End
                  Begin VB.TextBox txtCausaleDocumento 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   15
                     TabStop         =   0   'False
                     Top             =   1380
                     Width           =   1935
                  End
                  Begin VB.CheckBox chkSpospensioneIva 
                     Caption         =   "Sospensione IVA"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   18
                     TabStop         =   0   'False
                     Top             =   2070
                     Width           =   1935
                  End
                  Begin DMTDataCmb.DMTCombo cboBancaAzienda 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   16
                     TabStop         =   0   'False
                     Top             =   1380
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
                  Begin DMTDataCmb.DMTCombo cboSezionale 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   0
                     Top             =   400
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
                  Begin DMTDataCmb.DMTCombo cboMagazzino 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   13
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
                     Left            =   2160
                     TabIndex        =   14
                     TabStop         =   0   'False
                     Top             =   900
                     Width           =   2160
                     _ExtentX        =   3810
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
                     Left            =   2160
                     TabIndex        =   1
                     Top             =   405
                     Width           =   1185
                     _Version        =   65536
                     _ExtentX        =   2090
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTEDITNUMLib.dmtNumber lngNumero 
                     Height          =   315
                     Left            =   3480
                     TabIndex        =   2
                     Top             =   405
                     Width           =   810
                     _Version        =   65536
                     _ExtentX        =   1429
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin DMTDATETIMELib.dmtDate txtDataPlafond 
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   233
                     Top             =   1875
                     Width           =   2145
                     _Version        =   65536
                     _ExtentX        =   3784
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Appearance      =   1
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Causale documento"
                     Height          =   255
                     Index           =   4
                     Left            =   2160
                     TabIndex        =   239
                     Top             =   2160
                     Width           =   1935
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Competenza plafond"
                     Height          =   195
                     Index           =   10
                     Left            =   2160
                     TabIndex        =   234
                     Top             =   1680
                     Width           =   1485
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Listino"
                     Height          =   195
                     Index           =   1
                     Left            =   2160
                     TabIndex        =   129
                     Top             =   720
                     Width           =   2130
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Numero"
                     Height          =   195
                     Index           =   2
                     Left            =   3480
                     TabIndex        =   128
                     Top             =   200
                     Width           =   555
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Data Doc."
                     Height          =   195
                     Index           =   0
                     Left            =   2160
                     TabIndex        =   127
                     Top             =   200
                     Width           =   720
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Magazzino"
                     Height          =   195
                     Index           =   3
                     Left            =   120
                     TabIndex        =   126
                     Top             =   720
                     Width           =   1935
                  End
                  Begin VB.Label lblDocument 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Sezionale"
                     Height          =   195
                     Index           =   5
                     Left            =   120
                     TabIndex        =   125
                     Top             =   200
                     Width           =   1875
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Banca azienda"
                     Height          =   255
                     Index           =   0
                     Left            =   2160
                     TabIndex        =   124
                     Top             =   1200
                     Width           =   1935
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Causale trasporto"
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   123
                     Top             =   1200
                     Width           =   1935
                  End
               End
               Begin VB.Frame Frame7 
                  Height          =   9255
                  Left            =   -74880
                  TabIndex        =   110
                  Top             =   600
                  Width           =   12855
                  Begin VB.CommandButton cmdNuovaCommissione 
                     Caption         =   "Nuovo"
                     Height          =   375
                     Left            =   8400
                     TabIndex        =   75
                     Top             =   2520
                     Width           =   1335
                  End
                  Begin VB.CommandButton cmdSalvaCommissione 
                     Caption         =   "Salva"
                     Height          =   375
                     Left            =   8400
                     TabIndex        =   74
                     Top             =   3240
                     Width           =   1335
                  End
                  Begin VB.CommandButton cmdEliminaCommissione 
                     Caption         =   "Elimina"
                     Height          =   375
                     Left            =   8400
                     TabIndex        =   76
                     Top             =   3960
                     Width           =   1335
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtImportoRigaComm 
                     Height          =   315
                     Left            =   8400
                     TabIndex        =   73
                     Top             =   1320
                     Visible         =   0   'False
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "€ 0"
                     ForeColor       =   65535
                     BackColor       =   65535
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
                  Begin DmtGridCtl.DmtGrid GrigliaCommissioni 
                     Height          =   3255
                     Left            =   120
                     TabIndex        =   112
                     Top             =   1680
                     Width           =   8175
                     _ExtentX        =   14420
                     _ExtentY        =   5741
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
                     EnableMove      =   0   'False
                     ColumnsHeaderHeight=   20
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtImportoCommissioni 
                     Height          =   315
                     Left            =   6240
                     TabIndex        =   72
                     Top             =   1320
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
                  Begin DMTEDITNUMLib.dmtNumber txtPercCommissioni 
                     Height          =   315
                     Left            =   5280
                     TabIndex        =   71
                     Top             =   1320
                     Width           =   855
                     _Version        =   65536
                     _ExtentX        =   1508
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboCommissioni 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   70
                     Top             =   1320
                     Width           =   5055
                     _ExtentX        =   8916
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
                  Begin DMTEDITNUMLib.dmtNumber txtImportoTotaleCommissioni 
                     Height          =   375
                     Left            =   5520
                     TabIndex        =   111
                     Top             =   4560
                     Visible         =   0   'False
                     Width           =   2775
                     _Version        =   65536
                     _ExtentX        =   4895
                     _ExtentY        =   661
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   65535
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Tahoma"
                        Size            =   14.25
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
                  Begin VB.Label lblTotaleMercePerComm 
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Left            =   120
                     TabIndex        =   162
                     Top             =   600
                     Width           =   9615
                  End
                  Begin VB.Label Label15 
                     Caption         =   "TERMINARE E SALVARE IL DOCUMENTO PRIMA DI INSERIRE LE COMMISSIONI"
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   375
                     Left            =   120
                     TabIndex        =   161
                     Top             =   240
                     Width           =   10335
                  End
                  Begin VB.Label Label11 
                     Caption         =   "Tipo di commissioni"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   117
                     Top             =   1080
                     Width           =   5055
                  End
                  Begin VB.Label Label11 
                     Caption         =   "Perc."
                     Height          =   255
                     Index           =   1
                     Left            =   5280
                     TabIndex        =   116
                     Top             =   1080
                     Width           =   855
                  End
                  Begin VB.Label Label11 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Importo"
                     Height          =   255
                     Index           =   2
                     Left            =   6240
                     TabIndex        =   115
                     Top             =   1080
                     Width           =   1335
                  End
                  Begin VB.Label Label11 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Importo riga"
                     Height          =   255
                     Index           =   3
                     Left            =   8400
                     TabIndex        =   114
                     Top             =   1080
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label Label12 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Importo totale commissioni"
                     Height          =   255
                     Index           =   0
                     Left            =   5520
                     TabIndex        =   113
                     Top             =   4320
                     Visible         =   0   'False
                     Width           =   2775
                  End
               End
               Begin VB.Label lblInfoTesta 
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   156
                  Top             =   480
                  Width           =   12615
               End
            End
         End
         Begin DmtGridCtl.DmtGrid BrwMain 
            Height          =   2415
            Left            =   0
            TabIndex        =   235
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
         Left            =   120
         TabIndex        =   158
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
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
         Left            =   0
         Top             =   1440
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   0
         MousePointer    =   9  'Size W E
         Top             =   240
         Width           =   60
      End
   End
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   77
      Top             =   10800
      Width           =   17880
      _ExtentX        =   31538
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
Private oAgManager As DmtAgentsLib.AgentManager

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
Private bLoading As Boolean

Private A_Riga(1) As Long
'Variabili recordset per le visualizzazioni delle griglie
Private rsGriglia As DmtOleDbLib.adoResultset

Private NuovoRecordComm As Integer


Private NumeroRigaSelezionata As Integer
Private NumeroRecordPerModifica As Integer
Private NumeroRecordLista As Integer

Private Mov As DmtMovim.cMovimentazione

Public LOADING_NEW_DOC As Boolean

Private Const IDDocumento As Long = 15
Private Sel_IDArt_dettaglio As Long
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
    Dim Col As dmtgridctl.dgColumnHeader
    Dim bValue As Boolean

    'Legge le impostazioni dal registry
    bValue = IIf(IsMissing(IDVisible), AppOptions.IDFieldsVisibility, IDVisible)

    For Each Col In BrwMain.ColumnsHeader
        If Left(Col.FieldName, 2) = "ID" Then
            Col.Visible = bValue
        End If
    Next Col
    
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
    PermissionToSave = True
    
    If dtData.Value = 0 Then
        sbMsgInfo "Impossibile salvare il documento corrente senza aver specificato la data documento", m_App.FunctionName
        'dtData.SetFocus
        PermissionToSave = False
        Exit Function
    End If
    
    If cboPagamento.CurrentID = 0 Then
        sbMsgInfo "Impossibile salvare il documento corrente senza aver specificato la modalità di pagamento", m_App.FunctionName
        'cboPagamento.SetFocus
        PermissionToSave = False
        Exit Function
    End If
    If cdAnagrafica.KeyFieldID = 0 Then
        sbMsgInfo "Impossibile salvare il documento corrente senza aver specificato il cliente", m_App.FunctionName
        'cdAnagrafica.SetFocus
        PermissionToSave = False
        Exit Function
    End If
    If LINK_BLOCCO_CLIENTE = 1 Then
        sbMsgInfo "Impossibile salvare il documento corrente poichè il cliente risulta bloccato", m_App.FunctionName
        'cdAnagrafica.SetFocus
        PermissionToSave = False
        Exit Function
    End If



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
On Error Resume Next
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
On Error Resume Next
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
    LOADING_NEW_DOC = True
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
    Me.txtDataPlafond.Value = Date

    DATA_COMPETENZA_LIQ = Me.dtData.Text
    ANNOTAZIONE_01 = GET_NOTA_DOCUMENTO(oDoc.IDTipoOggetto, 1)
    ANNOTAZIONE_02 = GET_NOTA_DOCUMENTO(oDoc.IDTipoOggetto, 2)
    ANNOTAZIONE_03 = GET_NOTA_DOCUMENTO(oDoc.IDTipoOggetto, 3)


    fnSetSezionale
    fnSetCausaleDocumento
    
    Me.cboMagazzino.WriteOn fnGetParametriMagazzino("IDMagazzino_Vendita")
    Me.cboValuta.WriteOn oDoc.DBDefaults.Link_Val_valuta_nazionale
    fnSetProtocolloICE
    Me.cboPagamento.WriteOn oDoc.DBDefaults.IDPagamentoDocDefault
    oDoc.Field "Link_Doc_contratto_bancario_az", oDoc.DBDefaults.Link_Doc_contratto_bancario_az, sTabellaTestata
    'Imposta la data con la data di sistema
    'oDoc.Field "Doc_data", Date, sTabellaTestata
    
    NumeroRiga = 0

    Me.chkCessione.Value = 0
    Me.cboIntraSezDaComp.ListIndex = 0
    Me.cboIntraTrimestre.ListIndex = 0
    
    Me.cboIntraDatiRichiesti.WriteOn 0
    Me.cboIntraMese.WriteOn 0
    Me.cboIntraProvincia.WriteOn 0
    Me.cboIntraTrasporto.WriteOn 0
    Me.txtIntraAnno.Value = 0
    Me.cboIntraNazione.WriteOn 0
    
    chkCessione_Click

    'Inizializza il documento
    oDoc.InitializeDocument
    
    LOADING_NEW_DOC = False
    'Imposta la riga attiva per la tabella di testata (sempre la prima ed unica riga)
    'oDoc.Tables(sTabellaTestata).SetActiveRetail 1
    
    'Imposta il magazzino di default
    oDoc.ReadDataFromStore oDoc.IDMagazzinoDefault, MainStore
    
    'Imposta la valuta con la valuta nazionale

    oDoc.Field "Link_Doc_contratto_bancario_az", oDoc.DBDefaults.Link_Doc_contratto_bancario_az, sTabellaTestata

    oDoc.Field "Link_Val_valuta", oDoc.DBDefaults.Link_Val_valuta_nazionale, sTabellaTestata
    oDoc.Field "Link_Val_cambio", Null, sTabellaTestata
    oDoc.Field "Doc_data_inizio_trasporto", Date, sTabellaTestata
    oDoc.Field "Doc_ora_inizio_trasporto", time, sTabellaTestata
    oDoc.Field "RV_POIndicaProtICE", fnNormBoolean(1), sTabellaTestata
    
    IDListinoDefault = GET_LISTINO_DEFAULT(Me.cdAnagrafica.KeyFieldID)
    If IDListinoDefault = 0 Then
        IDListinoDefault = GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA
    End If
    
    Me.cboListino.WriteOn IDListinoDefault
    oDoc.Field "Link_Doc_Listino", IDListinoDefault, sTabellaTestata
    
    
    Rif_UltimoIDOggetto = 0
    Rif_UltimoTipoOggetto = 0
    Rif_UltimoNumeroDoc = ""
    Rif_UltimoDataDoc = ""
    Rif_UltimoPrefissoDoc = ""
    Rif_LetteraIntento = 0
    
    NuovoDocumento = 0
    'Aggiorna il contenuto delle listview
    fnEliminaDatiTemporanei
    
    sbPopalaListaArticoli True
    sbPopalaListaIva
    sbPopalaListaScadenze
    
    'Si predispone per l'inserimento di un nuovo articolo
    'cmdNuovo_Click
    
    CONTROLLA_BLOCCHI_INSERIMENTI
    SSTab1.TabEnabled(2) = False
    SSTab1.Tab = 0
    If Me.dtData.Enabled = True Then
        dtData.SetFocus
    End If

    'Refresh delle variabili di stato
    m_Search = False
    m_Changed = False
    m_Saved = False
    
    'Refresh della toolbar in modalità inserimento
    SetStatus4Modality 0 'Insert
    
    'Ripristina la vista del Form
    BrwMain.Visible = False
    
    'Il primo campo del Form riceve l'input focus
    
    
    
    
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
    
    If sType = "fpDateTime" Or sType = "TextBox" Or sType = "fpText" Or sType = "fpLongInteger" Or sType = "fpCurrency" Or sType = "fpDoubleSingle" Or sType = "dmtDate" Then
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
        ctrControl.Text = ""
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDAnagrafica = 0
        ctrControl.Code = ""
        ctrControl.Description = ""
        ctrControl.SecondDescription = ""
    ElseIf sType = "DmtFirmGerarchy" Then
        ctrControl.LoadActivity 0
    ElseIf sType = "DMTProgControl" Then
        'Queste istruzioni forzano il refresh
        'e il reset del componente
        ctrControl.IDArticolo = 0
        ctrControl.Show
    ElseIf sType = "DmtCodDesc" Then
        ctrControl.Load 0
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
    ClearControl cboAltroSito
    ClearControl ACSSocioTesta
    ClearControl txtDataPlafond
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
        Case "DatiEFatturaDocumento"
            Command3_Click
        Case "DatiEFatturaRigaDocumento"
            Command4_Click
'        Case "RiscontroPeso"
'            cmdRiscontroPeso_Click
'        Case "PrezzaturaVeloce"
'            Command1_Click
        Case "DatiContropartite"
            cmdPianoDeiContiRiga_Click
        Case "DatiAgenteRiga"
            cmdAgenteRiga_Click
'        Case "DatiTour"
'            cmdTour_Click
        Case "AvviaAltriDati"
            Gestione_Altri_dati
        Case "RigeneraDescrizioniCausaliXML"
            RICALCOLA_DESCRIZIONI_CAUSALI
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
        bLoading = True
        oDoc.ClearValues
        'Leggiamo il documento
        oDoc.ReadWithTO m_Document("IDOggetto").Value, m_Document("IDTipoOggetto").Value
        
        'Aggiorniamo il contenuto delle listview
        'fnEliminaDatiTemporanei
        'Viene impostata la variabile per indicare la lettura del documento è terminata
        bLoading = False
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
Dim cl As dmtgridctl.dgColumnHeader
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
        oDoc.SetTipoOggetto 107 'Fattura immediata
        'Imposta la funzione di DMT che gestisce il tipo di documento interessato (Ved. tabella Funzione)
        oDoc.IDFunzione = 109 'Fattura immediata
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
        oDoc.Descrizione = GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto)
        
        'Imposta l'identificativo dell'utente corrente
        oDoc.IDUtente = TheApp.IDUser
        
        oDoc.UpdateOnlyModified = False
    End If






    'Inizializza la combo dei listini
    With cboListino
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDListino"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Listino"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDListino, Listino FROM Listino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino = 0"
        .SQL = .SQL & " ORDER BY Listino"
    End With
    With Me.cboBancaAzienda
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDBancaPerAnagrafica"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "BancaPerAnagrafica"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT BancaPerAnagrafica.IDBancaPerAnagrafica, BancaPerAnagrafica.BancaPerAnagrafica "
        .SQL = .SQL & "FROM Anagrafica INNER JOIN "
        .SQL = .SQL & "Azienda ON Anagrafica.IDAnagrafica = Azienda.IDAnagrafica INNER JOIN "
        .SQL = .SQL & "BancaPerAnagrafica ON Azienda.IDAnagrafica = BancaPerAnagrafica.IDAnagrafica "
        .SQL = .SQL & " WHERE ((BancaPerAnagrafica.IDAzienda = " & TheApp.IDFirm & "))"
        .SQL = .SQL & " ORDER BY BancaPerAnagrafica.BancaPerAnagrafica"
    End With

    With Me.cboMagazzino
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDMagazzino"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Magazzino"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDMagazzino, Magazzino FROM Magazzino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " ORDER BY Magazzino"
    End With
    With Me.cboSezionale
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDSezionale"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Sezionale"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT  Sezionale.IDSezionale, Sezionale.Sezionale, RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "FROM Sezionale INNER JOIN "
        .SQL = .SQL & "RegistroIvaPerTipoOggetto ON Sezionale.IDRegistroIva = RegistroIvaPerTipoOggetto.IDRegistroIva AND "
        .SQL = .SQL & "Sezionale.IDFiliale = RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & oDoc.IDTipoOggetto
        .SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & oDoc.IDFiliale
        .SQL = .SQL & " ORDER BY Sezionale.Sezionale"
        '.SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDTipoModulo=1 "
        '.SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDRegistroIVA=1 "    End With
    End With
    'Inizializza la combo dei pagamenti
    With cboPagamento
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDPagamento"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Pagamento"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDPagamento, Pagamento FROM Pagamento"
        .SQL = .SQL & " ORDER BY Pagamento"
    End With
    With Me.cboPorto
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDPorto"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Porto"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDPorto, Porto FROM Porto"
        .SQL = .SQL & " ORDER BY Porto"
    End With
    With Me.cboTrasporto
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDTipoSpedizione"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "TipoSpedizione"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDTipoSpedizione, TipoSpedizione FROM TipoSpedizione"
        .SQL = .SQL & " ORDER BY TipoSpedizione"
    End With
    With Me.cboVettore
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDVettore"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Vettore"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDVettore, Vettore FROM Vettore"
        .SQL = .SQL & " ORDER BY Vettore"
    End With
    With Me.cboAspettoEsteriore
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDAspettoEsterioreArticolo"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "AspettoEsterioreArticolo"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDAspettoEsterioreArticolo, AspettoEsterioreArticolo FROM AspettoEsterioreArticolo"
        .SQL = .SQL & " ORDER BY AspettoEsterioreArticolo"
    End With

    With Me.cboAliquotaArticolo
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDIva"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Codice"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDIva, Codice FROM Iva"
        .SQL = .SQL & " ORDER BY Codice"
    End With

    With Me.cboIvaCliente
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDIva"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Iva"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDIva, Iva FROM Iva"
        .SQL = .SQL & " ORDER BY AliquotaIva"
    End With

    With Me.cboIntraDatiRichiesti
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDTipoDatoRichiesto"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "TipoDatoRichiesto"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM TipoDatoRichiesto"
        
    End With

    With Me.cboIntraTrasporto
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDModoDiTrasporto"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "ModoDiTrasporto"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM ModoDiTrasporto"
        
    End With

    With Me.cboIntraMese
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDMese"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Mese"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM Mese"
        
    End With
    With Me.cboIntraProvincia
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDProvincia"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "NomeProvincia"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM Provincia"
        
    End With

    With Me.cboIntraNazione
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtOledbLib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDNazione"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Nazione"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM Nazione ORDER BY Nazione"
    End With
    
    With Me.cboIntraSezDaComp
        .Clear
        .AddItem ""
        .ItemData(0) = 0
        .AddItem "Cessione"
        .ItemData(1) = 1
        .AddItem "Rettifica"
        .ItemData(2) = 2
    End With

    With Me.cboIntraTrimestre
        .Clear
        .AddItem ""
        .ItemData(0) = 0
        .AddItem "01"
        .ItemData(1) = 1
        .AddItem "02"
        .ItemData(2) = 2
        .AddItem "03"
        .ItemData(3) = 3
        .AddItem "04"
        .ItemData(4) = 4

    End With


    With Me.cboUnitaDiMisura
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDUnitaDiMisura"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "UnitaDiMisura"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT IDUnitaDiMisura, UnitaDiMisura FROM UnitaDiMisura"
        .SQL = .SQL & " ORDER BY UnitaDiMisura"
    End With

    'Inizializza il controllo Codice-Descrizione per la ricerca dei clienti
    With cdAnagrafica
        'Imposta l'oggetto Application del modello ad oggetti
        Set .Application = TheApp
        'Imposta la connessione corrente al database DMT
        Set .Database = TheApp.Database  '<-- Notare la connessione DMTDBLib.Database
        'Imposta l'Handle del form che contiene il controllo
        .HwndContainer = Me.hwnd
        'Imposta la descrizione dell'intestazione di colonna per il campo codice
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .CodeCaption4Find = "Cognome / Ragione sociale"
        'Imposta il campo da utilizzare come campo codice
        .CodeField = "Anagrafica"
        'indica se il codice è un campo numerico
        .CodeIsNumeric = False
        'Imposta la descrizione dell'intestazione di colonna per il campo descrizione
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .DescriptionCaption4Find = "Nome"
        'Imposta il campo da utilizzare come campo descrizione
        .DescriptionField = "Nome"
        'Indica il campo chiave univoco di accesso al record
        .KeyField = "IDAnagrafica"
        'Indica il nome della tabella o View da utilizzare per il reperimento dei dati
        .TableName = "IERepCliente"
        'Indica eventuali filtri fissi da utilizzare per l'estrazione dei record
        .Filter = "IDAzienda = " & TheApp.IDFirm
        'Abilita la voce del menu popup per l'esegui gestione
        .MenuFunctions("EseguiGestione").Enabled = True
        'Indica la funzione DMT da eseguire quando viene lanciata l'esegui gestione
        .IDExecuteFunction = 29 'Anagrafica
    End With
   
    
   
    'Articolo
    With Me.CDArticolo
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm
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
   
   
    
    'Nomenclatura combinata
    With Me.CDIntra_art_Nom_Comb
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceDoganale"
        .DescriptionField = "NomenclaturaCombinata"
        .KeyField = "IDNomenclaturaCombinata"
        .TableName = "NomenclaturaCombinata"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice doganale"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nomenclatura"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice doganale"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nomenclatura"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
  
    With Me.CDIntra_Imb_Nom_Comb
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceDoganale"
        .DescriptionField = "NomenclaturaCombinata"
        .KeyField = "IDNomenclaturaCombinata"
        .TableName = "NomenclaturaCombinata"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice doganale"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Nomenclatura"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice doganale"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Nomenclatura"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Natura transazione articolo
    With Me.CDIntra_art_Nat_Trans
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "NaturaTransazione"
        .KeyField = "IDNaturaTransazione"
        .TableName = "NaturaTransazione"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Transazione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Transazione"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Natura transazione imballo
    With Me.CDIntra_Imb_Nat_Trans
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "NaturaTransazione"
        .KeyField = "IDNaturaTransazione"
        .TableName = "NaturaTransazione"
        '.Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Transazione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Transazione"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        '.IDExecuteFunction = 6 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With


    With Me.cboCommissioni
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDRV_POTipoCommissione"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "TipoCommissione"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM RV_POTipoCommissione"
        .SQL = .SQL & " ORDER BY TipoCommissione"
    End With

 
    'Inizializza la ListView contenente l'elenco degli articoli presenti in un documento
    With lvwArticoli
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .HideSelection = False

        .ColumnHeaders.Add , , "NR", 300
        .ColumnHeaders.Add , , "Tipo", 300
        .ColumnHeaders.Add , , "ID Art.", 100
        .ColumnHeaders.Add , , "Cod. Articolo", 1000
        .ColumnHeaders.Add , , "Articolo", 2000
        .ColumnHeaders.Add , , "Q.tà", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Imp. Unit.", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Sc. 1", 700, lvwColumnRight
        .ColumnHeaders.Add , , "Sc. 2", 700, lvwColumnRight
        .ColumnHeaders.Add , , "Imp.Un.Net.", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "ID Lotto", 100
        .ColumnHeaders.Add , , "Cod. Lotto", 1000
        .ColumnHeaders.Add , , "Socio", 2000
        .ColumnHeaders.Add , , "Colli", 800, lvwColumnRight
        .ColumnHeaders.Add , , "Pezzi", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Peso lordo", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Tara", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Vol. Imballo", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "% IVA", 500, lvwColumnRight
        .ColumnHeaders.Add , , "Netto riga", 1300, lvwColumnRight
        .ColumnHeaders.Add , , "Lordo riga", 1300, lvwColumnRight

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
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtOledbLib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDValuta"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "Valuta"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM Valuta"
        .SQL = .SQL & " ORDER BY Valuta"
    End With
    
    With Me.cboCambioValuta
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtOledbLib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDCambio"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "DataCambio"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM Cambio"
    End With

    'Inizializza il controllo Codice-Descrizione per la ricerca dei clienti
    With Me.CDAgenteTesta
        'Imposta l'oggetto Application del modello ad oggetti
        Set .Application = TheApp
        'Imposta la connessione corrente al database DMT
        Set .Database = TheApp.Database  '<-- Notare la connessione DMTDBLib.Database
        'Imposta l'Handle del form che contiene il controllo
        .HwndContainer = Me.hwnd
        'Imposta la descrizione dell'intestazione di colonna per il campo codice
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .CodeCaption4Find = "Cognome / Ragione sociale"
        'Imposta il campo da utilizzare come campo codice
        .CodeField = "Anagrafica"
        'indica se il codice è un campo numerico
        .CodeIsNumeric = False
        'Imposta la descrizione dell'intestazione di colonna per il campo descrizione
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .DescriptionCaption4Find = "Nome"
        'Imposta il campo da utilizzare come campo descrizione
        .DescriptionField = "Nome"
        'Indica il campo chiave univoco di accesso al record
        .KeyField = "IDAnagrafica"
        'Indica il nome della tabella o View da utilizzare per il reperimento dei dati
        .TableName = "IERepAgente"
        'Indica eventuali filtri fissi da utilizzare per l'estrazione dei record
        .Filter = "IDAzienda = " & TheApp.IDFirm
        'Abilita la voce del menu popup per l'esegui gestione
        .MenuFunctions("EseguiGestione").Enabled = True
        'Indica la funzione DMT da eseguire quando viene lanciata l'esegui gestione
        .IDExecuteFunction = 29 'Anagrafica
    End With


    'Inizializza il controllo Codice-Descrizione per la ricerca dei clienti
    With Me.CDAgenteRiga
        'Imposta l'oggetto Application del modello ad oggetti
        Set .Application = TheApp
        'Imposta la connessione corrente al database DMT
        Set .Database = TheApp.Database  '<-- Notare la connessione DMTDBLib.Database
        'Imposta l'Handle del form che contiene il controllo
        .HwndContainer = Me.hwnd
        'Imposta la descrizione dell'intestazione di colonna per il campo codice
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .CodeCaption4Find = "Cognome / Ragione sociale"
        'Imposta il campo da utilizzare come campo codice
        .CodeField = "Anagrafica"
        'indica se il codice è un campo numerico
        .CodeIsNumeric = False
        'Imposta la descrizione dell'intestazione di colonna per il campo descrizione
        'da utilizzare quando si mostra la griglia di ricerca delle informazioni
        .DescriptionCaption4Find = "Nome"
        'Imposta il campo da utilizzare come campo descrizione
        .DescriptionField = "Nome"
        'Indica il campo chiave univoco di accesso al record
        .KeyField = "IDAnagrafica"
        'Indica il nome della tabella o View da utilizzare per il reperimento dei dati
        .TableName = "IERepAgente"
        'Indica eventuali filtri fissi da utilizzare per l'estrazione dei record
        .Filter = "IDAzienda = " & TheApp.IDFirm
        'Abilita la voce del menu popup per l'esegui gestione
        .MenuFunctions("EseguiGestione").Enabled = True
        'Indica la funzione DMT da eseguire quando viene lanciata l'esegui gestione
        .IDExecuteFunction = 29 'Anagrafica
    End With

    With Me.cboRegolaProvvigione
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtOledbLib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDRegolaProvv"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "RegolaProvv"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM RegolaProvv WHERE IDAzienda=" & TheApp.IDFirm
    End With


    With Me.cboTipoVariazione
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtOledbLib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDRV_POTipoVariazione"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "TipoVariazione"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM RV_POTipoVariazione"
    End With

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
        .SQL = "SELECT * FROM RV_POTipoImportoVenditaLiq "
        .Fill
    End With

    With Me.cboTipoDocumentoCoop
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoDocumentoCoop"
        .DisplayField = "TipoDocumentoCoop"
        .SQL = "SELECT * FROM RV_POTipoDocumentoCoop"
        .SQL = .SQL & " ORDER BY TipoDocumentoCoop"
    End With

    Set Me.ACSSocioTesta.Connection = TheApp.Database.Connection
    ACSSocioTesta.ApplicationName = App.Title
    ACSSocioTesta.Client = App.EXEName
    ACSSocioTesta.IDFirm = TheApp.IDFirm
    ACSSocioTesta.IDUser = TheApp.IDUser
    ACSSocioTesta.UserName = TheApp.User
    ACSSocioTesta.SearchType = DmtSearchSuppliers
    ACSSocioTesta.HwndContainer = Me.hwnd


    Set Me.ACSCliente.Connection = TheApp.Database.Connection
    ACSCliente.ApplicationName = App.Title
    ACSCliente.Client = App.EXEName
    ACSCliente.IDFirm = TheApp.IDFirm
    ACSCliente.IDUser = TheApp.IDUser
    ACSCliente.UserName = TheApp.User
    ACSCliente.SearchType = DmtSearchCustomers
    ACSCliente.HwndContainer = Me.hwnd

    With Me.cboTipoOrdine
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDTipoOrdine"
        .DisplayField = "TipoOrdine"
        .SQL = "SELECT * FROM TipoOrdine WHERE IDAzienda=" & m_App.IDFirm
        .SQL = .SQL & " ORDER BY TipoOrdine"
        .Fill
    End With

    Set Me.ACSAnaDest.Connection = TheApp.Database.Connection
    ACSAnaDest.ApplicationName = App.Title
    ACSAnaDest.Client = App.EXEName
    ACSAnaDest.IDFirm = TheApp.IDFirm
    ACSAnaDest.IDUser = TheApp.IDUser
    ACSAnaDest.UserName = TheApp.User
    ACSAnaDest.SearchType = DmtSearchCustomers
    ACSAnaDest.HwndContainer = Me.hwnd
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
        oActivity.Load m_DocType.ID, TheApp.Branch
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
    
    ParametroImballo
    ParametroLavorato
    ParametroGrezzo
    ParametroSocio
    ParametroObbligatorio
    ParametroTipoCaloPeso
    ParametroTipoAumentoPeso
    ParametroTipoScarto
    ParametroNumeroDecimali
    ParametroNuovoCalcolo
    'OnStart
    GET_MODULO_ATTIVATO MODULO_CODICE, 80
    COLLEGAMENTO_NOTA_PER_LOTTO = REC_DATI_LIQ_FILIALE_LONG("CollegamentoNotaPerCodiceLotto")
    NON_CALC_COMM = REC_DATI_LIQ_FILIALE_LONG("NonCalcolareCommissioniInNDeNC")
    NON_CALC_INCIDENZA_IMB = REC_DATI_LIQ_FILIALE_LONG("NonCalcolareIncidenzaImballoInNDeNC")
    RECUPERA_CONFIG_CAUS_XML
    OnBeforeOpenDoc
    
    
    If Len(m_App.Caller) > 0 And m_App.CallerFieldValue > 0 Then
        '-------------------------------------------------
        '     Il programma è stato chiamato da un link.
        '-------------------------------------------------
        
        'In tal caso occorre mostrare in modalità variazione il record richiesto dal programma client.
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next

        'Imposta una condizione di ricerca basata sull'ID del record richiesto dal programma client.
        m_DocType.Fields("IDOggetto").Value = m_App.CallerFieldValue
        
        
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"

        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")

        'Inidica, nel caso di esegui gestione, se riportare il valore corrente al chiamante
        bNotReturnValue = CBool(Val(GetSetting(REGISTRY_KEY, App.EXEName, "NoReturnValue", "0")))

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
        m_Document.Dataset.Recordset.Sort = "Doc_data Desc, Doc_numero Desc"
        Set BrwMain.Recordset = m_Document.Dataset.Recordset
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
   ' With DmtPrnDlg
   '     Set .Application = m_App
   '     Set .DocType = m_DocType
   ' End With
    
    
    
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
Private Function ConditionType(ByVal DBType As Integer) As dmtgridctl.ConditionTypeConstants
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
    Dim Cond As dmtgridctl.dgCondition
    
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
    BrwMain.Conditions.WidthConditions = 350
    BrwMain.Conditions.WidthFields = 250
    BrwMain.Conditions.WidthIntervals = 100
    
    BrwMain.Title.BackColor = vb3DFace
    BrwMain.Title.ForeColor = vbBlack
    BrwMain.Title.Font.Bold = True

    'Inserisce il filtro per data documento
    Set Cond = BrwMain.Conditions.Add("Link_doc_sezionale", "Sezionale", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.Indentation = 20
        Cond.RecordSource = "SELECT * FROM Sezionale WHERE IDFiliale=" & TheApp.Branch & "  ORDER BY Sezionale"
        Cond.DisplayField = "Sezionale"
        Cond.KeyField = "IDSezionale"
    Set Cond = BrwMain.Conditions.Add("Doc_data", "Data doc.", m_DocType.TableName, , True, False, dgCondTypeDate)
        Cond.RangeChecked = True
        Cond.FromValue = Date - 90
        Cond.ToValue = Date
    'Inserisce il filtro per numero documento
    Set Cond = BrwMain.Conditions.Add("Doc_numero", "Numero doc.", m_DocType.TableName, , True, False, dgCondTypeNumber)
        Cond.Indentation = 20
    'Inserisce il filtro per la ragione sociale
    Set Cond = BrwMain.Conditions.Add("Nom_codice", "Codice cliente", m_DocType.TableName, False, False, , dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("Nom_ragione_sociale_o_cognome", "Anagrafica", m_DocType.TableName, True, False, , dgCondTypeText)
        Cond.Indentation = 20
    Set Cond = BrwMain.Conditions.Add("SitoPerAnagrafica", "Destinazione diversa", m_DocType.TableName, False, False, , dgCondTypeText)
       Cond.Indentation = 20

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
    m_Report.PrinterName = ""
    
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
        
        'Se non verrà correttamente selezionato un elemento sarà restituito il valore -1 all'applicazione client.
        lIDField = -1
        
        'Se il documento è vuoto non si deve far nulla.
        'Se la browse è in modalità Filter Definition non formula la domanda di riporto dei dati nel programma chiamante.
        If (Not (m_Document.EOF And m_Document.BOF)) And (BrwMain.GuiMode <> dgFilterDefinition) Then
            If bNotReturnValue > 0 Then
                'ATTENZIONE: La stringa sMessage1 deve essere personalizzata a seconda dei casi!!!
                sMessage1 = " il " & m_DocType.Name
                sMessage = sMessage1 & " """ & Caption2Display(False) & """"
                
                gResource.CustomStrings.Clear
                gResource.CustomStrings.Add sMessage, 1
                   
                'Viene chiesto se si intende riportare il record corrente al programma chiamante.
                If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYPASTE), m_App.FunctionName) = vbYes Then
                    'Legge l'ID del record corrente affinchè venga riportato all'applicazione chiamante.
                    lIDField = m_Document.Fields("IDOggetto").Value
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
            If Me.ACSCliente.Enabled = True Then
                Me.ACSCliente.SetFocus
            End If
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
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "DatiEFatturaDocumento"
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaDocumento").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaDocumento").SetPicture 0, gResource.GetBitmap(IDB_ESTRATTOANALITICO_16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaDocumento").ToolTipText = "Dati e-fattura documento" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaDocumento").Description = "Dati e-fattura documento"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaDocumento").Caption = "Dati e-fattura documento"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep11"
    BarMenu.Bands("StandardPO").Tools("Sep11").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "DatiEFatturaRigaDocumento"
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaRigaDocumento").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaRigaDocumento").SetPicture 0, gResource.GetBitmap(IDB_ESTRATTOANALITICO_16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaRigaDocumento").ToolTipText = "Dati e-fattura riga documento" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaRigaDocumento").Description = "Dati e-fattura riga documento"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiEFatturaRigaDocumento").Caption = "Dati e-fattura riga documento"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep12"
    BarMenu.Bands("StandardPO").Tools("Sep12").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "DatiContropartite"
    BarMenu.Bands("StandardPO").Tools("DatiContropartite").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("DatiContropartite").SetPicture 0, gResource.GetBitmap(IDB_CONF_FOR16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("DatiContropartite").ToolTipText = "Visualizza contropartite riga documento" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiContropartite").Description = "Visualizza contropartite riga documento"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiContropartite").Caption = "Vis. contropartite riga doc."  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep15"
    BarMenu.Bands("StandardPO").Tools("Sep15").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "DatiAgenteRiga"
    BarMenu.Bands("StandardPO").Tools("DatiAgenteRiga").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("DatiAgenteRiga").SetPicture 0, gResource.GetBitmap(IDB_CONF_FOR16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("DatiAgenteRiga").ToolTipText = "Visualizza configurazione agente riga documento" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiAgenteRiga").Description = "Visualizza configurazione agente riga documento"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("DatiAgenteRiga").Caption = "Vis. config. agente riga doc."  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep16"
    BarMenu.Bands("StandardPO").Tools("Sep16").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep19"
    BarMenu.Bands("Standard").Tools("Sep19").ControlType = ddTTSeparator
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "AvviaAltriDati"
    BarMenu.Bands("Standard").Tools("AvviaAltriDati").Style = ddSIconText
    BarMenu.Bands("Standard").Tools("AvviaAltriDati").SetPicture 0, gResource.GetBitmap(IDB_ANNOTAZ16), &HC0C0C0
    BarMenu.Bands("Standard").Tools("AvviaAltriDati").ToolTipText = "Gestione altri dati" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("AvviaAltriDati").Description = "Gestione altri dati"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("AvviaAltriDati").Caption = "Gestione altri dati"  'GetDescription4StatusBar("Mnu_FormView")
    
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "Sep20"
    BarMenu.Bands("Standard").Tools("Sep20").ControlType = ddTTSeparator
    ToolID = ToolID + 1
    BarMenu.Bands("Standard").Tools.Add ToolID, "RigeneraDescrizioniCausaliXML"
    BarMenu.Bands("Standard").Tools("RigeneraDescrizioniCausaliXML").Style = ddSIconText
    BarMenu.Bands("Standard").Tools("RigeneraDescrizioniCausaliXML").SetPicture 0, gResource.GetBitmap(IDB_AGG_PROGRESSIVI16), &HC0C0C0
    BarMenu.Bands("Standard").Tools("RigeneraDescrizioniCausaliXML").ToolTipText = "Ricalcola descrizioni causali XML" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("RigeneraDescrizioniCausaliXML").Description = "Ricalcola descrizioni causali XML"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("Standard").Tools("RigeneraDescrizioniCausaliXML").Caption = "Ricalcola descrizioni causali XML"  'GetDescription4StatusBar("Mnu_FormView")

    BarMenu.RecalcLayout
    
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
    Dim Cond As dmtgridctl.dgCondition
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
    Dim Cond As dmtgridctl.dgCondition
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
    
    m_DocType.Fields("IDAzienda").Value = TheApp.IDFirm
    
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
Private Sub OnSave()
On Error GoTo ERR_OnSave
    Dim Testo As String
    Dim Field As DmtDocManLib.Field
    Dim DocLink As DmtDocManLib.DocumentsLink
    Dim sSQL As String
    Dim OLD_Cursor As Long
    Dim NuovoDocumento As Boolean
    
    
    If MODULO_ATTIVATO = 0 Then
        If Len(MODULO_DESCRIZIONE) > 0 Then
            MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
        Else
            MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
        End If
    Exit Sub
    End If
    
    'Controlli preliminari sulla validità e consistenza dei dati da salvare
    If Not PermissionToSave Then
        Exit Sub
    End If
    
    'Se la proprietò IDOggetto dell'oggetto cDocument è uguale a 0 (zero)
    'vuol dire che si tratta di un nuovo documento e quindi procediamo con il suo
    'inserimento altrimenti effettuiamo un aggiornamento del documento esistente
    
    
    oDoc.AllowCreateMovements = False
    
    OLD_Cursor = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    If oDoc.IsLocked = True Then
        oDoc.UpdateRates = False
    Else
        If GET_CONTROLLO_GESTIONE_INCASSI(oDoc.IDOggetto, oDoc.IDTipoOggetto) = True Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "Le scadenze relative al documento corrente sono state passate in contabilità." & vbCrLf
            Testo = Testo & "Si desidera lasciare inalterate le scadenze? Scegliere:" & vbCrLf
            Testo = Testo & "Sì - Per non ricrearle" & vbCrLf
            Testo = Testo & "No - Per riscriverle (I dati precedenti verranno persi)" & vbCrLf
            Testo = Testo & "Annulla - Per interrompere il salvataggio dell'intero documento."
            
            
            Select Case MsgBox(Testo, vbQuestion + vbYesNoCancel, TheApp.FunctionName)
                Case vbYes
                    oDoc.UpdateRates = False
                Case vbNo
                    oDoc.UpdateRates = True
                Case vbCancel
                    Exit Sub
                Case Else
                    Exit Sub
            End Select
        Else
            oDoc.UpdateRates = True
        End If
    End If
    
    ConsolidaDettaglioFatturaElettronica
    oDoc.PerformDocument Nothing, False
    
    If Me.cboCambioValuta.CurrentID = 0 Then
        oDoc.Field "Link_Val_cambio", Null, sTabellaTestata
        oDoc.Field "Val_valore_cambio", Null, sTabellaTestata
        oDoc.Field "Val_data_cambio", Null, sTabellaTestata
    Else
        oDoc.Field "Link_Val_cambio", Me.cboCambioValuta.CurrentID, sTabellaTestata
        oDoc.Field "Val_valore_cambio", Me.txtValoreCambioValuta.Value, sTabellaTestata
        oDoc.Field "Val_data_cambio", Me.txtDataCambio.Text, sTabellaTestata
    End If
    
    If oDoc.IDOggetto = 0 Then
        NuovoDocumento = True
        Me.Caption = "SALVATAGGIO IN CORSO....................."
        'Crea il nuovo documento
        oDoc.Insert
    
    Else
        NuovoDocumento = False
        'Aggiorna il documento esistente
        sSQL = "DELETE FROM Movimento "
        sSQL = sSQL & "WHERE IDOggetto=" & oDoc.IDOggetto
        sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
        
        Cn.Execute sSQL

        Me.Caption = "SALVATAGGIO IN CORSO....................."
        If oDoc.Update = False Then
            MsgBox oDoc.LastErrorNumber
            Me.Caption = Caption2Display(False)
        End If
        
    End If
    
    'Se la proprietà oDoc.LastErrorNumber = 0 vuol dire che il savataggio è andato a buon fine
    If oDoc.LastErrorNumber = 0 Then
        
        'Dopo aver salvato il documento con la DmtDocs dobbiamo rinfrescare
        'il contenuto del modello ad oggetti in modo da mostrare correttamente i nuovi dati
        Link_IDOggetto_OLD = oDoc.IDOggetto
        
        AGGIORNA_RIGHE_DOCUMENTO sTabellaDettaglio
        
        AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI 0, 0
        
        SCRIVI_RIFERIMENTI sTabellaDettaglio
        
        
        'CREA_PROVV_AGENTI NuovoDocumento
        
        Cn.CursorLocation = OLD_Cursor
        
        If NuovoDocumento = True Then
            SCRIVI_CAUSALI_DOC oDoc.IDOggetto
        End If
        
        If NuovoDocumento = True Then
            RIPRISTINA_GRIGLIA_CONDIZIONI
        End If
        
        m_Document.OpenDoc
        'Dopo aver rinfrescato il contenuto del modello ad oggetti ci posizioniamo sul record corretto
        m_Document.FindLocalData "IDOggetto = " & Link_IDOggetto_OLD, sdSearchForward
        
        m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        'm_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("IDOggetto").Value, SemAllActions
        m_Semaphore.SetObjectAction m_DocType.ID, oDoc.IDOggetto, SemAllActions
        'Refresh delle variabili di stato
        m_Changed = False
        m_Search = False
        m_Saved = True
        Me.Caption = Caption2Display(False)
        
        'Refresh dello stato della ToolBar standard in modalità variazione
        SetStatus4Modality Modify
        Me.cboSezionale.Enabled = False
    Else
        sbMsgError "Si è verificato un errore durante il salvataggio del documento!" & vbCrLf & "L'errore è:" & oDoc.LastErrorNumber, TheApp.FunctionName
        AGGIORNA_RIGHE_DOCUMENTO sTabellaDettaglio
        AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI 0, 0
        'CREA_PROVV_AGENTI NuovoDocumento
        Cn.CursorLocation = OLD_Cursor
    End If
    
Exit Sub
ERR_OnSave:
    MsgBox Err.Description, vbCritical, "Salvataggio"
    Me.Caption = Caption2Display(False)
End Sub

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
    
    
    'Se si è in modalità tabellare potrebbe essere necessario sincronizzare
    'il documento con il record evidenziato nella browse
    If BrwMain.Visible = True Then
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    
    
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, oDoc.IDOggetto, SemDeleteAction) Then
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
        
        'Cancelliamo il docomento con la DmtDocs
        If oDoc.DeleteWithTO(oDoc.IDOggetto, oDoc.IDTipoOggetto) <> 0 Then
            'In ogni caso, se la cancellazione fa a buon fine, eliminiamo anche il record
            'con il modello ad oggetti in modo da tenere sincronizzata il
            If Not ((m_Document.EOF) And (m_Document.BOF)) Then
                m_Document.Delete
            End If
        
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
    MsgBox Err.Description, vbCritical, "Eliminazione documento"
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
Private Sub OnPrint(ByVal ToolName As String)
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
        oReport.Copies = m_Report.Copies
        oReport.Orientation = m_OrientamentoDefault
            
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



Private Sub ACSCliente_ChangedElement()
    If Me.ACSCliente.IDAnagrafica > 0 Then
        If Me.ACSCliente.IDAnagrafica <> Me.cdAnagrafica.KeyFieldID Then
            Me.cdAnagrafica.Load Me.ACSCliente.IDAnagrafica
        End If
    End If
End Sub

Private Sub ACSSocioTesta_LostFocus()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Sezionale_Socio As Long

''''''''''''''''''''''''''''SEZIONALE PARAMETRI CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT * FROM RV_PO01_ConfigurazioneSocioSez "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDTipoOggetto=" & 114
sSQL = sSQL & " AND IDAnagrafica=" & Me.ACSSocioTesta.IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Link_Sezionale_Socio = 0
Else
    Link_Sezionale_Socio = fnNotNullN(rs!IDSezionale)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If oDoc.IDOggetto = 0 Then
    If Link_Sezionale_Socio > 0 Then
        Me.cboSezionale.WriteOn Link_Sezionale_Socio
    End If
    sbImpostaDatiDocumento
Else
    If Me.ACSSocioTesta.IDAnagrafica <> oDoc.Field("RV_POIDAnagraficaSocio", , sTabellaTestata) Then
        sbImpostaDatiDocumento
    End If
End If
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
Private Sub cboAltroSito_Click()
'On Error Resume Next

   
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

Private Sub cboCommissioni_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POTipoCommissione "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & Me.cboCommissioni.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtPercCommissioni.Value = 0
Else
    Me.txtPercCommissioni.Value = fnNotNullN(rs!Percentuale)
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub cboIntraDatiRichiesti_Click()
sbImpostaDatiDocumento
End Sub

Private Sub cboIntraMese_Click()
    If Me.cboIntraMese.CurrentID > 0 Then
        Me.cboIntraTrimestre.ListIndex = 0
    End If
    sbImpostaDatiDocumento
End Sub



Private Sub cboIntraProvincia_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub cboIntraSezDaComp_Click()
    If Me.cboIntraSezDaComp.ListIndex = 2 Then
        Me.cboIntraMese.Enabled = True
        Me.cboIntraTrimestre.Enabled = True
        Me.txtIntraAnno.Enabled = True
    Else
        Me.cboIntraMese.Enabled = False
        Me.cboIntraTrimestre.Enabled = False
        Me.txtIntraAnno.Enabled = False
    End If
    sbImpostaDatiDocumento
End Sub



Private Sub cboIntraTrasporto_Click()
sbImpostaDatiDocumento
End Sub

Private Sub cboIntraTrimestre_Click()
    If Me.cboIntraTrimestre.ListIndex > 0 Then
        Me.cboIntraMese.WriteOn 0
    End If
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
Private Sub cboMagazzino_Click()
    If oDoc.IDOggetto = 0 Then
        Me.cboMagazzino.WriteOn fnGetParametriMagazzino("IDMagazzino_Vendita")
    End If
    sbImpostaDatiDocumento
End Sub

Private Sub cboPagamento_Click()
On Error GoTo ERR_cboPagamento_Click
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    oDoc.ReadDataFromPayment Me.cboPagamento.CurrentID, sTabellaTestata
    
    oDoc.Field "Doc_data_inizio_scadenza", oDoc.DataEmissione, sTabellaTestata
    
    oDoc.Field "Link_Val_valuta", Me.cboValuta.CurrentID, sTabellaTestata
    
    oDoc.Scadenze.ParametersChanged = True
    
    sbImpostaDatiDocumento
    
    oDoc.Payments.Delete
    
    sbCalcolaDocumento
    
    'Change

Exit Sub
ERR_cboPagamento_Click:
    MsgBox Err.Description, vbCritical, "cboPagamento_Click"
End Sub

Private Sub cboPorto_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub cboSezionale_Click()
    
    oDoc.IDSezionale = Me.cboSezionale.CurrentID
    
    If oDoc.IDOggetto = 0 Then
        oDoc.Field "Doc_numero", fnDocumentNumber(Me.dtData.Text), sTabellaTestata
    Else
        If oDoc.Field("Link_Doc_sezionale", , sTabellaTestata) <> Me.cboSezionale.CurrentID Then
            oDoc.Field "Doc_numero", fnDocumentNumber(Me.dtData.Text), sTabellaTestata
        End If
    End If
    
    oDoc.Field "Link_Doc_sezionale", oDoc.IDSezionale, sTabellaTestata
    oDoc.Field "Doc_prefisso", GET_PREFISSO_SEZ(oDoc.IDSezionale), sTabellaTestata
End Sub

Private Sub cboTrasporto_Click()
    If Me.cboTrasporto.CurrentID = 3 Then
        Me.cboVettore.Enabled = True
    Else
        Me.cboVettore.WriteOn 0
        Me.cboVettore.Enabled = False
    End If

    sbImpostaDatiDocumento
End Sub
Private Sub cboUnitaDiMisura_Click()
    Link_UMCoop = fnGetUMCoop(Me.cboUnitaDiMisura.CurrentID)
End Sub

Private Sub cboValuta_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If (cboValuta.CurrentID = oDoc.Field("Link_Val_valuta", , sTabellaTestata)) Then Exit Sub

If Me.cboValuta.CurrentID = oDoc.DBDefaults.Link_Val_valuta_nazionale Then
    Me.txtValoreCambioValuta.Enabled = False
    Me.txtDataCambio.Enabled = False
    Me.cboCambioValuta.WriteOn 0
Else
    If oDoc.IsLocked Then Exit Sub

    Me.cboCambioValuta.Refresh
    
    If Me.cboValuta.CurrentID <> oDoc.Field("Link_Val_valuta", , sTabellaTestata) Then
    
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
    End If
End If

sbImpostaDatiDocumento
End Sub
Private Sub cboVettore_Click()
    oDoc.ReadDataFromCarrier Me.cboVettore.CurrentID, MainCarrier, sTabellaTestata
    sbImpostaDatiDocumento
End Sub

Private Sub CDAgenteRiga_ChangeElement()
    If bVariazioneDettaglio = False Then
        Me.cboRegolaProvvigione.WriteOn GET_LINK_REGOLA_PROVV_AGE(Me.CDAgenteRiga.KeyFieldID, TheApp.IDFirm)
    End If
End Sub

Private Sub CDAgenteTesta_ChangeElement()
    oDoc.ReadDataFromAgent Me.CDAgenteTesta.KeyFieldID
    sbImpostaDatiDocumento
End Sub

Private Sub cdAnagrafica_ChangeElement()
On Error Resume Next
Dim IDListinoDefault As Long
Dim IDPagamento As Long
Dim IDAnagraficaDest  As Long
Dim IDSezionaleCliente As Long

    AggiornaAltreDestinazioni
    AggiornaContrattiBancariCliente

    Me.ACSCliente.IDAnagrafica = 0
    Me.ACSCliente.Description = ""
    Me.ACSCliente.Code = ""
    Me.ACSCliente.SecondDescription = ""
    
    Me.ACSCliente.sbLoadCFByIDAnagrafica 0, Me.cdAnagrafica.KeyFieldID


    If oDoc.IDOggetto = 0 Then
    'Legge tutti i dati relativi al cliente selezionato
        Me.cboPagamento.WriteOn 0
        'sbImpostaDatiDocumento
        
        oDoc.ReadDataFromCliFo cdAnagrafica.KeyFieldID, sTabellaTestata
        
        IDSezionaleCliente = GET_SEZ_PER_CLIENTE(oDoc.IDTipoOggetto, Me.cdAnagrafica.KeyFieldID)
        
        If IDSezionaleCliente > 0 Then
            Me.cboSezionale.WriteOn IDSezionaleCliente
        End If
        
        oDoc.ReadDataFromAgent GET_LINK_AGENTE_CLIENTE(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm)

        'If oDoc.IDOggetto = 0 Then
        '    Me.chkCessione.Value = GET_CESSIONE_INTRA_CLIENTE(Me.cdAnagrafica.KeyFieldID)
        'End If
        
        LINK_CLIENTE_IVA = fnNotNullN(oDoc.Field("Link_Nom_IVA", , sTabellaTestata))
        
        'If GET_CONTROLLO_NUMERO_LETTERE_INTENTO(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm, Year(Me.dtData.Text)) = 1 Then
        '    Me.txtIDLetteraIntento.Value = GET_LINK_LETTERA_INTENTO(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm, Year(Me.dtData.Text))
        
        Me.txtIDLetteraIntento.Value = fnNotNullN(oDoc.Field("Link_Nom_lettera_intento", , sTabellaTestata))
        LINK_CLIENTE_IVA = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, LINK_CLIENTE_IVA)
        'End If
        
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
            'sbImpostaDatiDocumento
            
            oDoc.ReadDataFromCliFo cdAnagrafica.KeyFieldID
            LINK_CLIENTE_IVA = fnNotNullN(oDoc.Field("Link_Nom_IVA", , sTabellaTestata))
            Me.cboIvaCliente.WriteOn LINK_CLIENTE_IVA
            'If TIPO_SCONTO_CLIENTE = 1 Then
            '    Me.lngScontoDocPer.Value = fnNotNullN(oDoc.Field("Nom_sconto", , sTabellaTestata))
            'End If
            If Me.dtData.Value > 0 Then
                IDPagamento = GET_MODALITA_PAGAMENTO(Me.dtData.Text, Me.cdAnagrafica.KeyFieldID)
                If IDPagamento > 0 Then
                    Me.cboPagamento.WriteOn IDPagamento
                End If
            End If
            
             Me.chkCessione.Value = GET_CESSIONE_INTRA_CLIENTE(Me.cdAnagrafica.KeyFieldID)
             oDoc.ReadDataFromAgent GET_LINK_AGENTE_CLIENTE(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm)

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
    
    'Me.cboPorto.SetFocus
End Sub


Private Sub CDArticolo_ChangeElement()
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim Link_Lotto As Long

If bVariazioneDettaglio = False Then
    If Me.CDArticolo.KeyFieldID > 0 Then
        sSQL = "SELECT IDIvaVendita, AliquotaIva, Articolo, IDUnitaDiMisuraVendita, "
        sSQL = sSQL & "NonRiportoIntrastat, MassaNettaInKg, IDNomenclaturaCombinata, "
        sSQL = sSQL & "RV_POIDNaturaTransazione, RV_POIDCalibro, RV_POIDTipoCategoria, RV_POIDTipoLavorazione, "
        sSQL = sSQL & "RV_POIDImballoVendita, IDTipoProdotto, RV_POMoltiplicatore "
        sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
        sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
        sSQL = sSQL & "WHERE IDArticolo = " & Me.CDArticolo.KeyFieldID
        Set rs = Cn.OpenResultset(sSQL)
        
        If rs.EOF = False Then
            If LINK_CLIENTE_IVA = 0 Then
                Me.cboAliquotaArticolo.WriteOn fnNotNullN(rs!IDIvaVendita)
            Else
                Me.cboAliquotaArticolo.WriteOn LINK_CLIENTE_IVA
            End If
            Me.txtDescrizioneArticolo.Text = fnNotNull(rs!Articolo)
            If fnNotNullN(rs!RV_POMoltiplicatore) = 0 Then
                Moltiplicatore = 1
            Else
                Moltiplicatore = fnNotNullN(rs!RV_POMoltiplicatore)
            End If
            Me.cboUnitaDiMisura.WriteOn fnNotNullN(rs!IDUnitaDiMisuraVendita)
            
            If Me.chkCessione.Value = Checked Then
                MassaNetta_Art = fnNotNullN(rs!MassaNettaInKg)
                Me.CDIntra_art_Nat_Trans.Load 2 'GET_LINK_NATURA_TRANSAZIONE(Me.CDArticolo.KeyFieldID, TheApp.IDFirm, TheApp.Branch)
                Link_Nat_Trans_Art = Me.CDIntra_art_Nat_Trans.KeyFieldID

                If fnNormBoolean(rs!NonRiportoIntrastat) = True Then
                    Me.chkRiportoIntra_Art = Checked
                Else
                    Me.chkRiportoIntra_Art = Unchecked
                    chkRiportoIntra_Art_Click
                End If
            End If
            
        End If
        
                    
        CalcolaTotaleRiga

    End If
    
End If



End Sub





Private Sub CDSocioFatt_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Nome FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & Me.CDSocioFatt.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtNomeSocioFatt.Text = ""
Else
    Me.txtNomeSocioFatt.Text = fnNotNull(rs!Nome)
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub chkCessione_Click()



    If Me.chkCessione.Value = Unchecked Then
        Me.cboIntraDatiRichiesti.Enabled = False
        Me.cboIntraSezDaComp.Enabled = False
        Me.cboIntraMese.Enabled = False
        Me.cboIntraTrasporto.Enabled = False
        Me.cboIntraTrimestre.Enabled = False
        Me.cboIntraProvincia.Enabled = False
        Me.txtIntraAnno.Enabled = False
        Me.chkRiportoIntra_Art.Enabled = False
        Me.chkRiportoIntra_Imb.Enabled = False
        Me.cboIntraNazione.Enabled = False

        Me.cboIntraDatiRichiesti.WriteOn 0
        Me.cboIntraSezDaComp.ListIndex = 0
        Me.cboIntraMese.WriteOn 0
        Me.cboIntraTrasporto.WriteOn 0
        Me.cboIntraTrimestre.ListIndex = 0
        Me.cboIntraProvincia.WriteOn 0
        Me.txtIntraAnno.Value = 0
        Me.cboIntraNazione.WriteOn 0
    Else
        Me.cboIntraDatiRichiesti.Enabled = True
        Me.cboIntraSezDaComp.Enabled = True
        If (Me.chkCessione.Value = 1) And (Me.cboIntraSezDaComp.ListIndex = 1) Then
            Me.cboIntraMese.Enabled = False
            Me.cboIntraTrimestre.Enabled = False
            Me.txtIntraAnno.Enabled = False
        Else
            Me.cboIntraMese.Enabled = True
            Me.cboIntraTrimestre.Enabled = True
            Me.txtIntraAnno.Enabled = True
        End If
        Me.cboIntraTrasporto.Enabled = True
        Me.cboIntraProvincia.Enabled = True
        
        Me.chkRiportoIntra_Art.Enabled = True
        Me.chkRiportoIntra_Imb.Enabled = True
        Me.cboIntraNazione.Enabled = True
    End If
    
    If (oDoc.IDOggetto = 0) And (Me.chkCessione.Value = 1) Then
        ''''DEFAULT INTRA CLIENTE DI TESTA
        Me.cboIntraDatiRichiesti.WriteOn 1
        Me.cboIntraSezDaComp.ListIndex = 1
        Me.cboIntraTrasporto.WriteOn GET_LINK_MODO_DI_TRASPORTO(Me.cdAnagrafica.KeyFieldID, TheApp.IDFirm, TheApp.Branch)
        GET_CAMPI_INSTRASTAT_AZIENDA TheApp.IDFirm
        
    End If
    
    
    sbImpostaDatiDocumento
End Sub

Private Sub chkLordoIVA_Click()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    If (chkLordoIVA.Value = oDoc.Field("Doc_prezzi_lordo_IVA", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub


Private Sub chkRaggruppaScadenze_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub chkRaggruppBolle_Click()
    sbImpostaDatiDocumento
End Sub

Private Sub chkRiportoIntra_Art_Click()
    If Me.chkRiportoIntra_Art.Value = Checked Then
        Me.CDIntra_art_Nat_Trans.Enabled = False
        Me.CDIntra_art_Nom_Comb.Enabled = False
        Me.txtIntra_Art_MassaNetta.Enabled = False
        
        Me.CDIntra_art_Nat_Trans.Load 0
        Me.CDIntra_art_Nom_Comb.Load 0
        Me.txtIntra_Art_MassaNetta.Value = 0
        
    Else
        Me.CDIntra_art_Nat_Trans.Enabled = True
        Me.CDIntra_art_Nom_Comb.Enabled = True
        Me.txtIntra_Art_MassaNetta.Enabled = True
        
        Me.CDIntra_art_Nat_Trans.Load 2 'GET_LINK_NATURA_TRANSAZIONE(Me.CDArticolo.KeyFieldID, TheApp.IDFirm, TheApp.Branch)
        Me.CDIntra_art_Nom_Comb.Load GET_LINK_NOMENCLATURA_COMBINATA(Me.CDArticolo.KeyFieldID, TheApp.IDFirm)
        Me.txtIntra_Art_MassaNetta.Value = GET_QUANTITA_MASSA_NETTA_INSTRASTAT(Me.CDArticolo.KeyFieldID, TheApp.IDFirm, Me.txtQta_UM.Value)
    
    End If
End Sub

Private Sub chkSpospensioneIva_Click()
    If (chkSpospensioneIva.Value = oDoc.Field("Nom_IVA_in_spospensione", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
End Sub

Private Sub cmdAgenteRiga_Click()
Me.FraPDCRiga.Visible = False

If Me.FraAgenteRiga.Visible = False Then
    Me.FraAgenteRiga.Visible = True
    Me.lvwArticoli.Top = 4320
    Me.lvwArticoli.Height = 2055
Else
    Me.FraAgenteRiga.Visible = False
    Me.lvwArticoli.Top = 3360
    Me.lvwArticoli.Height = 3015
End If
End Sub

Private Sub cmdCambioValuta_Click()
    If Me.cboValuta.CurrentID = oDoc.DBDefaults.Link_Val_valuta_nazionale Then Exit Sub
    
    frmCambioValuta.Show vbModal
End Sub

Private Sub cmdElencoProdottiMix_Click()
    frmComponentiMix.Show vbModal
End Sub

Private Sub cmdElimina_Click()
    'Se è stato selezionato una riga nella listview degli articoli
    
    
    If Not lvwArticoli.SelectedItem Is Nothing Then
        'Rimuoviamo il dettaglio selezionato dall'oggetto cDocument
        oDoc.Tables(sTabellaDettaglio).RemoveRetail lvwArticoli.SelectedItem.Index
        EliminaDettaglioFatturaElettronica ID_ART_PROG_MERCE
        'Aggiorna il contenuto della listview articoli
        sbPopalaListaArticoli False
        'Ricalcola il totale del documento
        sbCalcolaDocumento
        'Si predispone per l'inserimento di un nuovo dettaglio
        'cmdNuovo_Click
    End If
    
    
End Sub

Private Sub cmdEliminaCommissione_Click()
Dim sSQL As String
If NuovoRecordComm = 0 Then
    sSQL = "DELETE FROM RV_POCommissioniPerDoc "
    sSQL = sSQL & "WHERE IDRV_POCommissioniPerDoc=" & fnNotNullN(Me.GrigliaCommissioni("IDRV_POCommissioniPerDoc").Value)
    Cn.Execute sSQL
    
    fnGrigliaCommissioni
    
End If


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

sbImpostaDatiDocumento
Exit Sub
ERR_cmdEliminaRifLetInt_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRifLetInt_Click"
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

Private Sub cmdNuovaCommissione_Click()
    NuovoRecordComm = 1
    
    Me.cboCommissioni.WriteOn 0
    Me.txtPercCommissioni.Value = 0
    Me.txtImportoCommissioni.Value = 0
    Me.txtImportoRigaComm.Value = 0
    
    'Me.cboCommissioni.SetFocus
End Sub

Private Sub cmdNuovo_Click()
On Error GoTo ERR_cmdNuovo_Click
    'Imposta la variabile per indicare che stiamo inserendo un nuovo dettaglio
    If oDoc.IsLocked = True Then Exit Sub
    
    bVariazioneDettaglio = False

    
    A_Riga(0) = 0
    A_Riga(1) = 0
    
    'Aggiungiamo una riga alla tabella di dettaglio del documento
    'oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
    
    'Azzeriamo i campi del dettaglio
    
    Me.CDArticolo.Load 0
    Me.txtDescrizioneArticolo.Text = ""
    Me.cboAliquotaArticolo.WriteOn 0
    Me.txtAliquotaArticolo.Value = 0
    Me.cboUnitaDiMisura.WriteOn 0
    Me.txtImportoUnitarioArticolo.Value = 0
    
    Me.txtQta_UM.Value = 0
    Me.txtColli.Value = 0
    Me.txtPesoLordo.Value = 0
    Me.txtTara.Value = 0
    Me.txtPesoNetto.Value = 0
    Me.txtPezzi.Value = 0
    
    Me.txtCodiceLottoVendita.Text = ""
    
    
    Me.txtImponibileUnitario.Value = 0
    CalcolaTotaleRiga
    

    Me.chkRiportoIntra_Art.Value = vbChecked
    Me.CDIntra_art_Nat_Trans.Load 0
    Me.CDIntra_art_Nom_Comb.Load 0
    Me.txtIntra_Art_MassaNetta.Value = 0
    
    Me.chkRiportoIntra_Imb.Value = 0
    Me.CDIntra_Imb_Nat_Trans.Load 0
    Me.CDIntra_Imb_Nom_Comb.Load 0
    Me.txtIntra_Imb_MassaNetta.Value = 0
    
    Me.txtDataConferimento.Text = ""
    Link_Socio = 0
    Me.txtCodiceSocio.Text = ""
    Me.txtSocio.Text = ""
    Me.lblNomeSocio.Caption = ""
    Link_RigaConferimento = 0
    Me.txtConferimentoRighe.Text = ""
    Me.CDSocioFatt.Load 0
    Me.txtDocumentoRiferimento.Text = ""
    Me.cboTipoDocumentoCoop.WriteOn 0
    
    Me.txtQuantitaOriginale.Value = 0
    Me.txtPrezzoOriginale.Value = 0
    Me.txtIDTipoVariazione.Value = 0
    Me.chkPrezzoMedioInLiq.Value = GET_PREZZO_MEDIO_CLIENTE(Me.cdAnagrafica.KeyFieldID)
    Me.cboTipoImportoLiq.WriteOn GET_FORZATURA_PREZZO_LIQ_CLIENTE(Me.cdAnagrafica.KeyFieldID)
    Me.txtImportoLiqVarMan.Value = 0
    Me.txtIDPianoDeiConti.Value = 0
    Me.txtDescrizioneConto.Text = ""
    Me.txtCodiceConto.Text = ""
    Me.cboTipoVariazione.WriteOn 0

    Me.txtPesoLordoOriginale.Value = 0
    Me.txtPesoNettoOriginale.Value = 0
    Me.txtColliOriginali.Value = 0
    Me.txtPezziOriginali.Value = 0
    Me.txtDataLavorazione.Value = 0
    Me.txtPrezzoOriginale.Value = 0
    Me.txtTaraOriginale.Value = 0
    Me.txtQuantitaOriginale.Value = 0
    Me.chkRiscontroPeso.Value = vbUnchecked
    
    Rif_PA_Riga_Doc_Merce = ""
    Rif_PA_Riga_Doc_Imballo = ""
    ID_ART_PROG_IMBALLO = 0
    ID_ART_PROG_MERCE = 0
    Link_TipoUtilizzoProcesso = 0
    
    'Da il fuoco al campo per la selezione dell'articolo

    If Me.CDAgenteTesta.KeyFieldID > 0 Then
        Me.CDAgenteRiga.Load Me.CDAgenteTesta.KeyFieldID
        Me.cboRegolaProvvigione.WriteOn GET_LINK_REGOLA_PROVV_AGE(Me.CDAgenteRiga.KeyFieldID, TheApp.IDFirm)
    End If
    
    If Me.SSTab1.Tab = 1 Then
        If Me.cdAnagrafica.KeyFieldID > 0 Then
            frmSelezionaRigaFattura.Show vbModal
        
            If Elaborazione_da_wizard = True Then
                If RiportaTuttoDocumento = False Then
                    sbElaborazioneRighe
                Else
                    sbElaborazioneRigheAll
                End If
            End If
        Else
            MsgBox "Inserire il cliente", vbInformation, "Elaborazione"
            Me.SSTab1.Tab = 0
            
            Me.ACSCliente.SetFocus
        End If
    End If
    '
Exit Sub
ERR_cmdNuovo_Click:
    MsgBox Err.Description, vbCritical, "cmdNuovo_Click"
End Sub

Private Sub cmdPianoDeiContiRiga_Click()
Me.FraAgenteRiga.Visible = False

If Me.FraPDCRiga.Visible = False Then
    Me.FraPDCRiga.Visible = True
    Me.lvwArticoli.Top = 4320
    Me.lvwArticoli.Height = 2055
Else
    Me.FraPDCRiga.Visible = False
    Me.lvwArticoli.Top = 3360
    Me.lvwArticoli.Height = 3015
End If
End Sub

Private Sub cmdSalva_Click()
On Error GoTo ERR_cmdSalva_Click

AggiornaRiga = 1

If PermessoSalvataggio = True Then
    If bVariazioneDettaglio = False Then
        NumeroRiga = NumeroRiga + 1
    End If
    
        
        SalvataggioRiga
                    
        oDoc.PerformTable sTabellaDettaglio, True
        
        'Aggiorna il contenuto della listview degli articoli
        sbPopalaListaArticoli False
        'Ricalcola il documento
        sbCalcolaDocumento
        AggiornaRiga = 0
        'Se eravamo in presenza di un nuovo dettaglio
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
        
        Me.txtDocumentoRiferimento.Text = GET_DOCUMENTO_DI_RIFERIMENTO
End If

Exit Sub

ERR_cmdSalva_Click:
    MsgBox Err.Description, vbCritical, "cmdSalva_Click"
End Sub
Private Function PermessoSalvataggio() As Boolean

PermessoSalvataggio = True


If Link_RigaConferimento > 0 Then
    If Me.txtQta_UM.Value > Me.txtQuantitaOriginale.Value Then
        MsgBox "La quantità non può essere maggiore della quantità della riga del documento collegato", vbCritical, "Impossibile salvare"
        PermessoSalvataggio = False
        Me.txtQta_UM.SetFocus
        Exit Function
    End If

'    If Me.txtImportoUnitarioArticolo.Value > Me.txtPrezzoOriginale.Value Then
'        MsgBox "Il prezzo unitario non può essere maggiore del prezzo unitario della riga del documento collegato", vbCritical, "Impossibile salvare"
'        PermessoSalvataggio = False
'        Me.txtImportoUnitarioArticolo.SetFocus
'        Exit Function
'    End If

End If

    


If Me.CDArticolo.KeyFieldID = 0 And Len(Trim(Me.txtDescrizioneArticolo.Text)) = 0 Then
    MsgBox "Inserire un articolo", vbCritical, "Impossibile salvare"
    PermessoSalvataggio = False
    Me.CDArticolo.SetFocus
    Exit Function
End If
 
 
If Me.CDArticolo.KeyFieldID = 0 Then
    If Me.cboTipoVariazione.CurrentID = 0 Then
        If Link_RigaConferimento > 0 Then
            MsgBox "Il tipo di variazione deve essere valorizzato", vbCritical, "Impossibile salvare"
            PermessoSalvataggio = False
            Me.cboTipoVariazione.SetFocus
            Exit Function
        End If
    End If
End If
 
If Me.CDArticolo.KeyFieldID > 0 Then
    If Me.cboUnitaDiMisura.CurrentID = 0 Then
        MsgBox "Deve essere impostata l'unità di misura dell'articolo", vbCritical, "Impossibile salvare"
        PermessoSalvataggio = False
        Me.cboUnitaDiMisura.SetFocus
        Exit Function
    End If
End If

If Me.CDArticolo.KeyFieldID > 0 Then
    If Me.txtQta_UM.Value < 0 Then
        MsgBox "ATTENZIONE!!!" & vbCrLf & "La quantita da movimentare deve essere maggiore o uguale a zero"
        PermessoSalvataggio = False
        Exit Function
    End If
End If
'If Link_Articolo > 0 Then
'    If Link_Socio = 0 Then
'        MsgBox "ATTENZIONE!!!" & vbCrLf & "Inserire il socio", vbInformation, "Salvataggio dati"
'        PermessoSalvataggio = False
'        Exit Function
'    End If
'End If
'If Link_Articolo > 0 Then
'    If Len(Me.txtDataConferimento.Text) = 0 Then
'        MsgBox "ATTENZIONE!!!" & vbCrLf & "Inserire la data di conferimento", vbInformation, "Salvataggio dati"
'        PermessoSalvataggio = False
'        Exit Function
'    End If
'End If
'If Link_Articolo > 0 Then
'    If Link_RigaConferimento = 0 Then
'        If Par_OBBLIGATORIO = 1 Then
'            MsgBox "ATTENZIONE!!!" & vbCrLf & "Inserire la riga di conferimento", vbInformation, "Salvataggio dati"
'            PermessoSalvataggio = False
'            Exit Function
'        End If
'    End If
'End If

End Function
Private Function ControllaParametroIvaBloccata() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IvaBloccata FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDAzienda=" & m_App.IDFirm & " AND "
sSQL = sSQL & "IDFIliale=" & m_App.Branch & " AND "
sSQL = sSQL & "IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    ControllaParametroIvaBloccata = False
Else
    If IsNull(rs!IvaBloccata) Then
        ControllaParametroIvaBloccata = False
    Else
        ControllaParametroIvaBloccata = rs!IvaBloccata
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Sub cmdSalvaCommissione_Click()
Dim sSQL As String
If Me.cboCommissioni.CurrentID = 0 Then
    MsgBox "Inserire il tipo di commissione", vbInformation, "Salvataggio dati"
    Exit Sub
End If

If NuovoRecordComm = 1 Then
    sSQL = "INSERT INTO RV_POCommissioniPerDoc ("
    sSQL = sSQL & "IDRV_POCommissioniPerDoc, IDOggetto, IDRV_POTipoCommissione, Percentuale, Importo, ImportoRiga) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & fnGetNewKey("RV_POCommissioniPerDoc", "IDRV_POCommissioniPerDoc") & ", "
    sSQL = sSQL & oDoc.IDOggetto & ", "
    sSQL = sSQL & Me.cboCommissioni.CurrentID & ", "
    sSQL = sSQL & fnNormNumber(Me.txtPercCommissioni.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtImportoCommissioni.Value) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtImportoRigaComm.Value) & ")"
Else
    sSQL = "UPDATE RV_POCommissioniPerDoc SET "
    sSQL = sSQL & "IDRV_POTipoCommissione=" & Me.cboCommissioni.CurrentID & ", "
    sSQL = sSQL & "Percentuale=" & fnNormNumber(Me.txtPercCommissioni.Value) & ", "
    sSQL = sSQL & "Importo=" & fnNormNumber(Me.txtImportoCommissioni.Value) & ", "
    sSQL = sSQL & "ImportoRiga=" & fnNormNumber(Me.txtImportoRigaComm.Value) & " "
    sSQL = sSQL & "WHERE IDRV_POCommissioniPerDoc=" & fnNotNullN(Me.GrigliaCommissioni("IDRV_POCommissioniPerDoc").Value)
End If

Cn.Execute sSQL

fnGrigliaCommissioni

End Sub

Private Sub curScontoDocImp_LostFocus()
    If (curScontoDocImp.Value = oDoc.Field("Sco_ad_importo_fine_documento", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub

Private Sub curSpeseIncasso_LostFocus()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    If (curSpeseIncasso.Value = oDoc.Field("Spe_incasso_neutro", , sTabellaTestata)) Then Exit Sub
    
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub

Private Sub curSpeseTrasporto_LostFocus()
    'Aggiorna l'oggetto cDocument della DmtDocs con i dati del form
    If (curSpeseTrasporto.Value = oDoc.Field("Spe_trasporto_neutro", , sTabellaTestata)) Then Exit Sub
    
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub

Private Sub dmtNumber3_Change()

End Sub

Private Sub dtData_LostFocus()
Dim LinkPagamento_Prec As Long
Dim IDPagamento As Long
Dim Testo As String

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
        
        Me.txtDataPlafond.Value = Me.dtData.Value
        
        If Me.dtData.Value > 0 Then
            IDPagamento = GET_MODALITA_PAGAMENTO(Me.dtData.Text, Me.cdAnagrafica.KeyFieldID)
            If IDPagamento > 0 Then
                Me.cboPagamento.WriteOn IDPagamento
            End If
        End If
        
        If oDoc.IDOggetto > 0 Then
            If Me.dtData.Text <> DATA_COMPETENZA_LIQ Then
                Testo = "La data del documento è cambiata." & vbCrLf
                Testo = Testo & "Vuoi cambiare la data di competenza di liquidazione?"
                If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then Exit Sub
                DATA_COMPETENZA_LIQ = Me.dtData.Text
                oDoc.UpdateOnlyModified = False
                sbImpostaDatiDocumento
            End If
        Else
            DATA_COMPETENZA_LIQ = Me.dtData.Text
            oDoc.UpdateOnlyModified = False
            sbImpostaDatiDocumento
        End If
        
        
    End If
    
'    If oDoc.IDOggetto > 0 Then
'        If Me.dtData.Text <> oDoc.Field("Doc_data_inizio_trasporto", , sTabellaTestata) Then
'            Testo = "La data del documento è cambiata." & vbCrLf
'            Testo = Testo & "Vuoi cambiare la data di inizio trasporto?"
'            If MsgBox(Testo, vbQuestion + vbYesNo, TheApp.FunctionName) = vbNo Then Exit Sub
'            Me.txtDataTrasporto.Value = Me.dtData.Value
'            oDoc.UpdateOnlyModified = False
'            sbImpostaDatiDocumento
'        End If
'    Else
'        Me.txtDataTrasporto.Value = Me.dtData.Value
'        oDoc.UpdateOnlyModified = False
'        sbImpostaDatiDocumento
'    End If
    
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

        m_bOnFirstTime = False

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
    If KeyCode = vbKeyF12 Then
        If fnGestioneArticoli(10) = True Then
            If frmMain.CDArticolo.KeyFieldID = 0 Then
                Link_ArticoloPadre = GET_ARTICOLO_CONFERITO
                If Link_ArticoloPadre > 0 Then
                    frmArticoliDerivati.Show vbModal
                End If
            End If
        End If
    End If
    If KeyCode = vbKeyF11 Then
        Link_Oggetto = oDoc.IDOggetto
        If Link_Oggetto > 0 Then
            frmChiusuraConferimento.Show vbModal
        End If
    End If
    If KeyCode = vbKeyF4 Then
        If Link_RigaConferimento > 0 Then
            frmRiepilogo.Show vbModal
        End If
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
Private Sub BrwMain_Reposition(ByVal AllColumns As dmtgridctl.dgColumns)
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




Private Sub GrigliaCommissioni_Reposition(ByVal AllColumns As dmtgridctl.dgColumns)
    NuovoRecordComm = 0

    Me.cboCommissioni.WriteOn Me.GrigliaCommissioni("IDRV_POTipoCommissione").Value
    Me.txtPercCommissioni.Value = Me.GrigliaCommissioni("Percentuale").Value
    Me.txtImportoCommissioni.Value = Me.GrigliaCommissioni("Importo").Value
End Sub


Private Sub Label13_Click(Index As Integer)
On Error GoTo ERR_Label13_Click
Dim IDOggetto As Long
Dim IDTipoOggetto As Long
Dim IDValoriOggettoDettaglio As Long

    
    If Index = 9 Then
        IDOggetto = 0
        IDTipoOggetto = 0
        IDValoriOggettoDettaglio = 0
          
        IDOggetto = fnNotNullN(oDoc.Field("RV_POIDOggetto", , sTabellaDettaglio))
        IDTipoOggetto = fnNotNullN(oDoc.Field("RV_POIDTipoOggetto", , sTabellaDettaglio))
        IDValoriOggettoDettaglio = fnNotNullN(oDoc.Field("RV_POIDValoriOggettoDettaglio", , sTabellaDettaglio))
        If IDTipoOggetto > 0 Then
            Me.txtDataLavorazione.Text = GET_VALORE_CAMPO_RIGA_VENDITA(IDTipoOggetto, IDOggetto, IDValoriOggettoDettaglio, "RV_PODataLavorazione", True)
        End If
    End If

Exit Sub
ERR_Label13_Click:
    MsgBox Err.Description, vbCritical, "Label13_Click"
    
End Sub

Private Sub lblCodiceSocio_Click()
    Dim oSearch As dmtFind.Find
    Dim sSQL As String
    Dim oRes As DmtOleDbLib.adoResultset
   
   'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = Cn
    
    'La Caption della finestra di ricerca
    oSearch.Caption = "Soci"
    
    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Comune" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Codice", "Codice", 1
    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1
    
    
    
    'oSearch.AddDisplayField "Listino", "Listino", 0
    'oSearch.AddDisplayField "Sconto", "ScontoPercentuale", 0
    
    
    'Imposta un valore di default per il filtro sul campo Comune
    
    oSearch.Filters.Add "Anagrafica", Me.txtSocio.Text
    oSearch.Filters.Add "Codice", Me.txtCodiceSocio.Text
    
    'Con la query SQL montata sotto l'impostazione di questa proprietà
    'non è necessaria
    'oSearch.IDName = "Comune.IDComune"
 
    'Quando si apre la finestra viene effettuta una ricerca preliminare con il filtro
    'SELECT ... FROM ... WHERE Comune LIKE ‘XXX%’  (essendo TextBox(0).Text = "XXX")
    oSearch.Start = Me.txtCodiceSocio.Text

    'Query SQL con cui effettuare le ricerche in base dati.
    'Attenzione:
    'Il campo chiave primaria (Comune.IDComune in questo caso) deve essere presente
    'nella SELECT
        sSQL = "SELECT  Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.Cap, Comune.Comune, Provincia.Provincia, Anagrafica.Indirizzo, "
        sSQL = sSQL & "Fornitore.Codice "
        sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
        sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
        sSQL = sSQL & "Anagrafica INNER JOIN "
        sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica ON Comune.IDComune = Anagrafica.IDComune "
        sSQL = sSQL & "WHERE ((IDAzienda=" & m_App.IDFirm & ") AND (Anagrafica.IDCategoriaAnagrafica=" & Link_TipoSocio & "))"
        
    
    'Assegnazione della query di ricerca
    oSearch.SQL = fnAnsi2Jet(sSQL)
    
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
   
    
    Set oRes = oSearch.Exec
    
    
    If Not oRes.EOF Then
        If oRes!IDAnagrafica > 0 Then
            Me.txtCodiceSocio.Text = fnNotNull(oRes!Codice)
            Link_Socio = fnNotNullN(oRes!IDAnagrafica)
            Me.lblNomeSocio.Caption = fnNotNull(oRes!Nome)
            Me.txtSocio.Text = fnNotNull(oRes!Anagrafica)
        End If
    End If
            

    
    Set oRes = Nothing
    Set oSearch = Nothing

End Sub




Private Sub lblDocument_Click(Index As Integer)
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtOleDbLib.adoResultset


If Index = 24 Then
    If Me.CDArticolo.Code = "" Then
        'Crea un'istanza dell'oggetto Find
        Set oSearch = New dmtFind.Find
        
        'Assegna la connessione aperta
        oSearch.Database = Cn
        
        'La Caption della finestra di ricerca
        oSearch.Caption = "Articoli"
        
        'Vengono assegnati i campi su cui effettuare la ricerca.
        'Questi campi verranno visualizzati nella tabella della finestra di ricerca
        'ed in quella per l'impostazione del filtro.
        'NOTA:
        'Il primo campo inserito (ovvero "Comune" in questo caso) verrà associato alla combo
        'presente nella finestra di ricerca.
        
        oSearch.AddDisplayField "Articolo", "Articolo", 1
        oSearch.AddDisplayField "Codice Articolo", "CodiceArticolo", 1

        'oSearch.AddDisplayField "Listino", "Listino", 0
        'oSearch.AddDisplayField "Sconto", "ScontoPercentuale", 0
        
        'Imposta un valore di default per il filtro sul campo Comune
        
        oSearch.Filters.Add "Articolo", Me.txtDescrizioneArticolo.Text
        
        'Con la query SQL montata sotto l'impostazione di questa proprietà
        'non è necessaria
        'oSearch.IDName = "Comune.IDComune"
     
        'Quando si apre la finestra viene effettuta una ricerca preliminare con il filtro
        'SELECT ... FROM ... WHERE Comune LIKE ‘XXX%’  (essendo TextBox(0).Text = "XXX")
        oSearch.Start = Me.txtDescrizioneArticolo.Text
    
        'Query SQL con cui effettuare le ricerche in base dati.
        'Attenzione:
        'Il campo chiave primaria (Comune.IDComune in questo caso) deve essere presente
        'nella SELECT
            sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo, IDUnitaDiMisuraAcquisto "
            sSQL = sSQL & "FROM Articolo "
            sSQL = sSQL & "WHERE (((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL)) "

            sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & "))"
            
        
        
        
        'Assegnazione della query di ricerca
        oSearch.SQL = fnAnsi2Jet(sSQL)
        
        
        'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
       
        
        Set oRes = oSearch.Exec
        
        
        If Not oRes.EOF Then
                Me.CDArticolo.Load fnNotNullN(oRes!IDArticolo)

        End If
        Set oRes = Nothing
        Set oSearch = Nothing
        
    End If
    
End If

End Sub

Private Sub lblPianodeiDeiConti_Click()
    Link_PianoDeiConti = GetPianoDeiConti

    SetPDCProperties
End Sub

Private Sub lblSocio_Click()
    Dim oSearch As dmtFind.Find
    Dim sSQL As String
    Dim oRes As DmtOleDbLib.adoResultset
   
   'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = Cn
    
    'La Caption della finestra di ricerca
    oSearch.Caption = "Soci"
    
    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Comune" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.

    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1
    oSearch.AddDisplayField "Codice", "Codice", 1
    
    
    'oSearch.AddDisplayField "Listino", "Listino", 0
    'oSearch.AddDisplayField "Sconto", "ScontoPercentuale", 0
    
    
    'Imposta un valore di default per il filtro sul campo Comune
    
    oSearch.Filters.Add "Anagrafica", Me.txtSocio.Text
    
    'Con la query SQL montata sotto l'impostazione di questa proprietà
    'non è necessaria
    'oSearch.IDName = "Comune.IDComune"
 
    'Quando si apre la finestra viene effettuta una ricerca preliminare con il filtro
    'SELECT ... FROM ... WHERE Comune LIKE ‘XXX%’  (essendo TextBox(0).Text = "XXX")
    oSearch.Start = Me.txtSocio.Text

    'Query SQL con cui effettuare le ricerche in base dati.
    'Attenzione:
    'Il campo chiave primaria (Comune.IDComune in questo caso) deve essere presente
    'nella SELECT
        sSQL = "SELECT  Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.Cap, Comune.Comune, Provincia.Provincia, Anagrafica.Indirizzo, "
        sSQL = sSQL & "Fornitore.Codice "
        sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
        sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
        sSQL = sSQL & "Anagrafica INNER JOIN "
        sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica ON Comune.IDComune = Anagrafica.IDComune "
        sSQL = sSQL & "WHERE ((IDAzienda=" & m_App.IDFirm & ") AND (Anagrafica.IDCategoriaAnagrafica=" & Link_TipoSocio & "))"
        
    
    'Assegnazione della query di ricerca
    oSearch.SQL = fnAnsi2Jet(sSQL)
    
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
   
    
    Set oRes = oSearch.Exec
    
    
    If Not oRes.EOF Then
        If oRes!IDAnagrafica > 0 Then
            Me.txtCodiceSocio.Text = fnNotNull(oRes!Codice)
            Link_Socio = fnNotNullN(oRes!IDAnagrafica)
            Me.lblNomeSocio.Caption = fnNotNull(oRes!Nome)
            Me.txtSocio.Text = fnNotNull(oRes!Anagrafica)
        End If
    End If
            

    
    Set oRes = Nothing
    Set oSearch = Nothing

End Sub

Private Sub lngNumero_Change()
    sbImpostaDatiDocumento
End Sub

Private Sub lngScontoDocPer_LostFocus()
    If (lngScontoDocPer.Value = oDoc.Field("Sco_percentuale_fine_documento", , sTabellaTestata)) Then Exit Sub
    
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub

Private Sub lvwArticoli_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Annulla l'inserimento di un nuovo dettaglio
    
    bVariazioneDettaglio = True
    Var_LostFocus_Colli = 0
    'Va in variazione della riga di dettaglio selezionata
    sbVariazioneRiga
    'Aggiorna il contenuto della listview degli articoli
    CalcolaTotaleRiga
     
    GET_PARAMETRI_LIQUIDAZIONE_ARTICOLO Me.CDArticolo.KeyFieldID
    
    Me.txtDocumentoRiferimento.Text = GET_DOCUMENTO_DI_RIFERIMENTO
    NumeroRecordLista = Me.lvwArticoli.SelectedItem.Index
    cmdElencoProdottiMix.Enabled = Link_TipoUtilizzoProcesso <> 0
    
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

    'oDoc.UpdateOnlyModified = True
    oDoc.UpdateOnlyModified = False
    
    If oDoc.IDOggetto = 0 Then
        NuovoDocumento = 0
        Me.SSTab1.TabEnabled(2) = False
        Me.cboSezionale.Enabled = True
        
        Me.cboIntraDatiRichiesti.Enabled = False
        Me.cboIntraSezDaComp.Enabled = False
        Me.cboIntraMese.Enabled = False
        Me.cboIntraTrasporto.Enabled = False
        Me.cboIntraTrimestre.Enabled = False
        Me.cboIntraProvincia.Enabled = False
        Me.txtIntraAnno.Enabled = False
        Me.chkRiportoIntra_Art.Enabled = False
        Me.chkRiportoIntra_Imb.Enabled = False
    Else
        NuovoDocumento = 1
        Me.SSTab1.TabEnabled(2) = True
        fnGrigliaCommissioni
        Me.cboSezionale.Enabled = False
        If Me.chkCessione.Value = Unchecked Then
            Me.cboIntraDatiRichiesti.Enabled = False
            Me.cboIntraSezDaComp.Enabled = False
            Me.cboIntraMese.Enabled = False
            Me.cboIntraTrasporto.Enabled = False
            Me.cboIntraTrimestre.Enabled = False
            Me.cboIntraProvincia.Enabled = False
            Me.txtIntraAnno.Enabled = False
            Me.chkRiportoIntra_Art.Enabled = False
            Me.chkRiportoIntra_Imb.Enabled = False
        Else
            Me.cboIntraDatiRichiesti.Enabled = True
            Me.cboIntraSezDaComp.Enabled = True
            Me.cboIntraMese.Enabled = True
            Me.cboIntraTrasporto.Enabled = True
            Me.cboIntraTrimestre.Enabled = True
            Me.cboIntraProvincia.Enabled = True
            Me.txtIntraAnno.Enabled = True
            Me.chkRiportoIntra_Art.Enabled = True
            Me.chkRiportoIntra_Imb.Enabled = True
        End If
        
        'fnEliminaDatiTemporanei
        ControllaNumeroRiga
        CONTROLLA_BLOCCHI_INSERIMENTI
        
        'chkCessione_Click
            
        sbPopalaListaArticoli True
        sbPopalaListaScadenze
        sbPopalaListaIva
    
    End If

    
    Me.SSTab1.Tab = 0
    
    
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
        'Ctr.SetFocus
        
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
                oDoc.Field "Doc_data_plafond", Value, sTabellaTestata
            End If
            
        Case "Doc_numero"
            oDoc.Numero = Value
            lngNumero.Value = Value
        Case "Doc_prezzi_lordo_IVA"
            chkLordoIVA.Value = Abs(CLng(Val(Value)))
            'Quando cambia il flag Prezzi lordo Iva ricalcola il documento
            'sbCalcolaDocumento
        Case "Nom_IVA_in_sospensione"
            Me.chkSpospensioneIva.Value = Abs(CLng(Val(Value)))
            'Quando cambia il flag Prezzi lordo Iva ricalcola il documento
            'sbCalcolaDocumento
        Case "Link_Doc_sezionale"
            cboSezionale.WriteOn Val(Value)
            oDoc.IDSezionale = Me.cboSezionale.CurrentID
            oDoc.Field "Link_Doc_sezionale", oDoc.IDSezionale, sTabellaTestata
        Case "Link_Doc_listino"
            cboListino.WriteOn Val(Value)
        Case "Link_Doc_magazzino"
            Me.cboMagazzino.WriteOn Value
        Case "Link_Nom_IVA"
            LINK_CLIENTE_IVA = Val(Value)
            Me.cboIvaCliente.WriteOn Val(Value)
        Case "Link_Doc_pagamento"
            cboPagamento.WriteOn Val(Value)
            'Quando cambia la modalità di pagamento ricalcola il documento
            'sbCalcolaDocumento
        Case "Doc_causale_trasporto"
            Me.txtCausaleDocumento.Text = Value
        Case "Link_Nom_anagrafica"
            cdAnagrafica.Load Val(Value)
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
            'sbCalcolaDocumento
        Case "Spe_trasporto_neutro"
            curSpeseTrasporto.Value = Value
            'Quando cambiano le spese di trasporto ricalcola il documento
            'sbCalcolaDocumento
        Case "Sco_percentuale_fine_documento"
            lngScontoDocPer.Value = Value
            'Quando cambia lo sconto in percentuale di fine documento ricalcola il documento
            'sbCalcolaDocumento
        Case "Sco_ad_importo_fine_documento"
            curScontoDocImp.Value = Value
            'Quando cambia lo sconto ad importo di fine documento ricalcola il documento
            'sbCalcolaDocumento
        Case "Tot_imponibile_corr"
            curTotImponibile.Value = Value
            
        Case "Tot_imposta_corr"
            curTotImposta.Value = Value
            
        Case "Tot_documento_corr"
            curTotDocumento.Value = Value
            
             oDoc.Scadenze.ParametersChanged = True
        Case "Tot_arrotondamenti"
            curTotArrotondamenti.Value = Value
            
        Case "Tot_netto_a_pagare_corr"
            curNettoAPagare.Value = Value
            
        Case "Link_Doc_contratto_bancario_az"
            Me.cboBancaAzienda.WriteOn Value
        Case "Link_Nom_contratto_bancario"
            Me.cboBancaCliente.WriteOn Value
            
        Case "Link_Nom_porto"
            Me.cboPorto.WriteOn Value
        Case "Link_Doc_spedizione"
            Me.cboTrasporto.WriteOn Value
        Case "Link_Vet_vettore"
            Me.cboVettore.WriteOn Value
        Case "Doc_data_inizio_trasporto"
            Me.txtDataTrasporto.Text = Value
        Case "Doc_ora_inizio_trasporto"
            Me.txtOraTrasporto.Text = Value
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
        Case "Nom_ult_sito_referente"
            Me.txtReferenteAltroSito.Text = Value
        Case "Nom_raggruppamento_bolle"
            Me.chkRaggruppBolle.Value = Abs(CLng(Val(Value)))
        Case "Nom_raggruppamento_scadenze"
            Me.chkRaggruppaScadenze.Value = Abs(CLng(Val(Value)))
        Case "RV_POIDProtICE"
            'Me.lblIDProtocolloICE.Caption = Value
        Case "RV_POIndicaProtICE"
            'Me.chkProtocolloICE.Value = Abs(CLng(Val(Value)))
        Case "RV_POProtICE"
           ' Me.txtProtocolloICE.Text = Value
        Case "RV_PONumeroProtICE"
            'Me.txtNumeroProtICE.Value = Value
        Case "Link_Doc_aspetto_esteriore"
            Me.cboAspettoEsteriore.WriteOn Val(Value)
        Case "Doc_annotazioni_variazio"
            Me.txtAnnotazioni.Text = Value
        Case "Doc_intra_cessione"
            Me.chkCessione.Value = Abs(CLng(Val(Value)))
        Case "Link_Doc_intra_dati_richiesti"
            Me.cboIntraDatiRichiesti.WriteOn Value
        Case "Link_Doc_intra_sezione"
            Me.cboIntraSezDaComp.ListIndex = Value
        Case "Link_Doc_intra_mese"
            Me.cboIntraMese.WriteOn Value
        Case "Link_Doc_intra_trimestre"
            Me.cboIntraTrimestre.ListIndex = Value
        Case "Doc_intra_anno"
            Me.txtIntraAnno.Value = Value
        Case "Link_Doc_intra_modo_trasporto"
            Me.cboIntraTrasporto.WriteOn Value
        Case "Link_Doc_intra_provinc_merce"
            Me.cboIntraProvincia.WriteOn Value
        Case "Val_data_cambio"
            Me.txtDataCambio.Text = Value
            'sbCalcolaDocumento
        Case "Val_valore_cambio"
            Me.txtValoreCambioValuta.Value = Value
            'sbCalcolaDocumento
        Case "Link_Val_valuta"
            Me.cboValuta.WriteOn Val(Value)
            'sbCalcolaDocumento
        Case "Link_Val_cambio"
            Me.cboCambioValuta.WriteOn Val(Value)
            'sbCalcolaDocumento
        Case "Tot_netto_a_pagare_naz"
            curNettoAPagare_naz.Value = Value
            
        Case "Link_Doc_intra_naz_pagamento"
            Me.cboIntraNazione.WriteOn Value
        Case "Link_Doc_agente"
            CONTROLLO_CAMBIO_AGENTE Me.CDAgenteRiga.KeyFieldID, fnNotNullN(Value)
            Me.CDAgenteTesta.Load fnNotNullN(Value)
        'Case "RV_POIDSitoPerAnagrafica"
            'Me.cboAltroSito.WriteOn fnNotNullN(Value)
        Case "RV_POIDAnagraficaSocio"
            If Value > 0 Then
                Me.ACSSocioTesta.IDAnagrafica = 0
                Me.ACSSocioTesta.Code = ""
                Me.ACSSocioTesta.Description = ""
                Me.ACSSocioTesta.SecondDescription = ""
                Me.ACSSocioTesta.sbLoadCFByIDAnagrafica 7, Value
            Else
                Me.ACSSocioTesta.IDAnagrafica = 0
                Me.ACSSocioTesta.Code = ""
                Me.ACSSocioTesta.Description = ""
                Me.ACSSocioTesta.SecondDescription = ""
            End If
        Case "Link_Nom_lettera_intento"
            Me.txtIDLetteraIntento.Value = Value
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
        Case "Doc_data_plafond"
            Me.txtDataPlafond.Text = Value
        Case "Nom_IVA_default"
            Me.chkNomIva.Value = Value
        Case "Nom_bollo_esente"
            Me.chkAddebitaBollo.Value = Abs(Value)
        Case "Doc_causale_documento"
            Me.txtCausaleDocumentoEF.Text = Value
        Case "RV_PODataCompetenzaLiq"
            DATA_COMPETENZA_LIQ = Value
        Case "RV_POAnnotazioni1"
            ANNOTAZIONE_01 = Value
        Case "RV_POAnnotazioni2"
            ANNOTAZIONE_02 = Value
        Case "RV_POAnnotazioni3"
            ANNOTAZIONE_03 = Value
        Case "Doc_numero_vs_ordine_di_rifer"
            Me.txtNumeroOrdineCliente.Text = Value
        Case "Doc_data_vs_ordine_di_rifer"
            Me.txtDataOrdineCliente.Text = Value
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
    Dim lRow As Long
    Dim oItem As MSComctlLib.ListItem
    Dim sSQL As String
    
    'Pulisce la listview
    lvwArticoli.ListItems.Clear
    
    With oDoc.Tables(sTabellaDettaglio)
        'Cicla per tutte le righe di dettaglio presenti nel documento
        For lRow = 1 To .NumRetails
            Set oItem = lvwArticoli.ListItems.Add
            
            'Popola l'item della listview
            
            
            oItem.Text = fnNotNullN(.Fields("RV_POLinkRiga").Values(lRow).Value)
            oItem.SubItems(1) = fnNotNullN(.Fields("RV_POTipoRiga").Values(lRow).Value)
            oItem.SubItems(2) = fnNotNullN(fnNotNullN(.Fields("Link_Art_articolo").Values(lRow).Value))
            oItem.SubItems(3) = fnNotNull(.Fields("Art_Codice").Values(lRow).Value)
            oItem.SubItems(4) = fnNotNull(.Fields("Art_descrizione").Values(lRow).Value)
            oItem.SubItems(5) = FormatNumber(fnNotNullN(.Fields("Art_quantita_totale").Values(lRow).Value), 2)
            oItem.SubItems(6) = FormatNumber(fnNotNullN(.Fields("Art_prezzo_unitario_netto_IVA").Values(lRow).Value), 5)
            oItem.SubItems(7) = FormatNumber(fnNotNullN(.Fields("Art_sco_in_percentuale_1").Values(lRow).Value), 2)
            oItem.SubItems(8) = FormatNumber(fnNotNullN(.Fields("Art_sco_in_percentuale_2").Values(lRow).Value), 2)
            If fnNotNullN(.Fields("RV_POTipoRiga").Values(lRow).Value) = 1 Then
                oItem.SubItems(9) = FormatNumber(fnNotNullN(.Fields("Art_pre_uni_net_sco_net_IVA").Values(lRow).Value), 4)
            Else
                oItem.SubItems(9) = FormatNumber(fnNotNullN(.Fields("Art_prezzo_unitario_netto_IVA").Values(lRow).Value), 5)
            End If
            oItem.SubItems(10) = fnNotNullN(.Fields("Link_Art_lotto_articolo").Values(lRow).Value)
            oItem.SubItems(11) = fnNotNull(.Fields("RV_POCodiceLotto").Values(lRow).Value)
            oItem.SubItems(12) = fnNotNull(.Fields("RV_POSocio").Values(lRow).Value)
            oItem.SubItems(13) = fnNotNullN(.Fields("Art_numero_colli").Values(lRow).Value)
            oItem.SubItems(14) = FormatNumber(fnNotNullN(.Fields("Art_quantita_pezzi").Values(lRow).Value), 2)
            oItem.SubItems(15) = FormatNumber(fnNotNullN(.Fields("Art_peso").Values(lRow).Value), 2)
            oItem.SubItems(16) = FormatNumber(fnNotNullN(.Fields("Art_tara").Values(lRow).Value), 2)
            oItem.SubItems(17) = FormatNumber(fnNotNullN(.Fields("Art_volume").Values(lRow).Value), 2)
            oItem.SubItems(18) = fnNotNullN(.Fields("Art_aliquota_IVA").Values(lRow).Value)
            oItem.SubItems(19) = FormatNumber(fnNotNullN(.Fields("Art_importo_totale_netto_IVA").Values(lRow).Value), 2)
            oItem.SubItems(20) = FormatNumber(fnNotNullN(.Fields("Art_importo_totale_lordo_IVA").Values(lRow).Value), 2)
            
            
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
        Tipo_Riga_Sel_EF = 1
        Me.CDArticolo.Load fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))
        Me.txtDescrizioneArticolo.Text = fnNotNull(oDoc.Field("Art_descrizione", , sTabellaDettaglio))
        Me.cboAliquotaArticolo.WriteOn fnNotNullN(oDoc.Field("Link_Art_IVA", , sTabellaDettaglio))
        Me.txtAliquotaArticolo.Value = fnNotNullN(oDoc.Field("Art_aliquota_IVA", , sTabellaDettaglio))
        Me.cboUnitaDiMisura.WriteOn fnNotNullN(oDoc.Field("Link_Art_unita_di_misura", , sTabellaDettaglio))
        Me.txtImportoUnitarioArticolo.Value = fnNotNullN(oDoc.Field("Art_prezzo_unitario_netto_IVA", , sTabellaDettaglio))
        Me.txtSconto1.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_1", , sTabellaDettaglio))
        Me.txtSconto2.Value = fnNotNullN(oDoc.Field("Art_sco_in_percentuale_2", , sTabellaDettaglio))
        Me.txtColli.Value = fnNotNullN(oDoc.Field("Art_numero_colli", , sTabellaDettaglio))
        Me.txtPesoLordo.Value = fnNotNullN(oDoc.Field("Art_peso", , sTabellaDettaglio))
        Me.txtTara.Value = fnNotNullN(oDoc.Field("Art_tara", , sTabellaDettaglio))
        Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
        Me.txtPezzi.Value = fnNotNullN(oDoc.Field("Art_quantita_pezzi", , sTabellaDettaglio))
        Me.txtQta_UM.Value = fnNotNullN(oDoc.Field("Art_quantita_totale", , sTabellaDettaglio))
        
'        Me.txtImponibileUnitario.Value = fnNotNullN(oDoc.Field("Art_Importo_netto_IVA", , sTabellaDettaglio))
'        Me.txtTotaleRiga.Value = fnNotNullN(oDoc.Field("Art_importo_totale_lordo_IVA", , sTabellaDettaglio))
        Me.txtImponibileUnitario.Value = fnNotNullN(oDoc.Field("Art_pre_uni_net_sco_net_IVA", , sTabellaDettaglio))
        Me.txtTotaleRiga.Value = fnNotNullN(oDoc.Field("Art_Importo_netto_IVA", , sTabellaDettaglio))
        
        Me.txtDataConferimento.Text = fnNotNull(oDoc.Field("RV_PODataConferimento", , sTabellaDettaglio))
        Link_RigaConferimento = fnNotNullN(oDoc.Field("RV_POIDConferimentoRighe", , sTabellaDettaglio))
        Link_TipoUtilizzoProcesso = fnNotNullN(oDoc.Field("RV_POIDTipoUtilizzoLinea", , sTabellaDettaglio))
        Link_RigaProcessoIVGamma = fnNotNullN(oDoc.Field("RV_POIDProcessoIVGamma", , sTabellaDettaglio))
        
        Me.txtConferimentoRighe.Text = GET_STRINGA_CONFERIMENTO(Link_RigaConferimento, fnNotNullN(oDoc.Field("RV_POIDAssegnazioneMerce", , sTabellaDettaglio)), fnNotNullN(oDoc.Field("RV_POIDProcessoIVGamma", , sTabellaDettaglio)))
        
        Link_Socio = fnNotNullN(oDoc.Field("RV_POIDSocio", , sTabellaDettaglio))
        Me.txtCodiceSocio.Text = fnNotNull(oDoc.Field("RV_POCodiceSocio", , sTabellaDettaglio))
        Me.txtSocio.Text = fnNotNull(oDoc.Field("RV_POSocio", , sTabellaDettaglio))
        Me.lblNomeSocio.Caption = fnNotNull(oDoc.Field("RV_PONomeSocio", , sTabellaDettaglio))
        Me.CDSocioFatt.Load fnNotNullN(oDoc.Field("RV_POIDAnagraficaFatturazione", , sTabellaDettaglio))

        Me.txtCodiceLottoVendita.Text = fnNotNull(oDoc.Field("RV_POCodiceLotto", , sTabellaDettaglio))
        Me.chkPrezzoMedioInLiq.Value = Abs(fnNotNullN(oDoc.Field("RV_POPrezzoMedioInLiq", , sTabellaDettaglio)))
        
        Me.CDAgenteRiga.Load fnNotNullN(oDoc.Field("Link_Art_agente", , sTabellaDettaglio))
        Me.cboRegolaProvvigione.WriteOn fnNotNullN(oDoc.Field("Link_Art_age_regola_provv", , sTabellaDettaglio))
        Me.txtPercProvv.Value = fnNotNullN(oDoc.Field("Art_age_percentuale_provv", , sTabellaDettaglio))
        Me.txtImportoProvv.Value = fnNotNullN(oDoc.Field("Art_age_importo_provv", , sTabellaDettaglio))
        Me.cboTipoOrdine.WriteOn fnNotNullN(oDoc.Field("Link_Art_age_tipo_ordine", , sTabellaDettaglio))

        Me.chkRiportoIntra_Art.Value = Abs(fnNotNullN(oDoc.Field("Art_intra_non_riporto", , sTabellaDettaglio)))
        Me.CDIntra_art_Nom_Comb.Load fnNotNullN(oDoc.Field("Link_Art_intra_nomenclatura", , sTabellaDettaglio))
        Me.CDIntra_art_Nat_Trans.Load fnNotNullN(oDoc.Field("Link_Art_intra_natura_trans", , sTabellaDettaglio))
        Me.txtIntra_Art_MassaNetta.Value = fnNotNullN(oDoc.Field("Art_intra_qta_tot_massa_netta", , sTabellaDettaglio))
        Me.cboTipoVariazione.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoVariazione", , sTabellaDettaglio))
        Me.cboTipoImportoLiq.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoImportoVenditaLiq", , sTabellaDettaglio))
        
        Me.cboTipoDocumentoCoop.WriteOn fnNotNullN(oDoc.Field("RV_POIDTipoDocumentoCoop", , sTabellaDettaglio))
        Me.txtImportoLiqVarMan.Value = fnNotNullN(oDoc.Field("RV_POVariazionePrezzoManuale", , sTabellaDettaglio))

        Me.txtIDPianoDeiConti.Value = oDoc.Field("Link_Art_IDCContropartita", , sTabellaDettaglio)
        Me.txtCodiceConto.Text = oDoc.Field("Art_CContropartita_codifica", , sTabellaDettaglio)
        Me.txtDescrizioneConto.Text = oDoc.Field("Art_CContropartita_descrizione", , sTabellaDettaglio)
        
        Me.txtDataLavorazione.Text = fnNotNull(oDoc.Field("RV_PODataLavorazione", , sTabellaDettaglio))
        Rif_PA_Riga_Doc_Merce = fnNotNull(oDoc.Field("Art_riferimento_PA", , sTabellaDettaglio))
        Sel_IDArt_dettaglio = fnNotNullN(oDoc.Field("ID_Art_dettaglio_prog", , sTabellaDettaglio))
        ID_ART_PROG_MERCE = fnNotNullN(oDoc.Field("ID_Art_dettaglio_prog", , sTabellaDettaglio))
        Me.chkRiscontroPeso.Value = Abs(fnNotNullN(oDoc.Field("RV_PORigaRiscontroPeso", , sTabellaDettaglio)))
        
    End If
 End Sub

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
    oDoc.Field "Nom_IVA_in_spospensione", Me.chkSpospensioneIva.Value, sTabellaTestata
    oDoc.Field "Link_Doc_listino", cboListino.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_pagamento", cboPagamento.CurrentID, sTabellaTestata
    oDoc.Field "Spe_incasso_neutro", curSpeseIncasso.Value, sTabellaTestata
    oDoc.Field "Spe_trasporto_neutro", curSpeseTrasporto.Value, sTabellaTestata
    oDoc.Field "Sco_percentuale_fine_documento", lngScontoDocPer.Value, sTabellaTestata
    oDoc.Field "Sco_ad_importo_fine_documento", curScontoDocImp.Value, sTabellaTestata
    oDoc.Field "Link_Nom_porto", Me.cboPorto.CurrentID, sTabellaTestata
    oDoc.Field "Link_Vet_vettore", Me.cboVettore.CurrentID, sTabellaTestata
    oDoc.Field "Tot_numero_colli", Me.txtColliTotali.Value, sTabellaTestata
    oDoc.Field "Tot_peso", Me.txtPesoTotale.Value, sTabellaTestata
    oDoc.Field "Link_Doc_magazzino", Me.cboMagazzino.CurrentID, sTabellaTestata
    oDoc.Field "Nom_raggruppamento_bolle", Me.chkRaggruppBolle.Value, sTabellaTestata
    oDoc.Field "Nom_raggruppamento_scadenze", Me.chkRaggruppaScadenze.Value, sTabellaTestata
    oDoc.Field "Doc_causale_trasporto", Me.txtCausaleDocumento.Text, sTabellaTestata
    oDoc.Field "Doc_causale_documento", Me.txtCausaleDocumentoEF.Text, sTabellaTestata
    
    oDoc.Field "Link_Nom_IVA", Me.cboIvaCliente.CurrentID, sTabellaTestata
    oDoc.Field "Link_Val_valuta", Me.cboValuta.CurrentID, sTabellaTestata
    'oDoc.Field "RV_POIDSitoPerAnagrafica", Me.cboAltroSito.CurrentID, sTabellaTestata
    oDoc.Field "RV_POIDAnagraficaSocio", Me.ACSSocioTesta.IDAnagrafica, sTabellaTestata
    oDoc.Field "Link_Nom_lettera_intento", Me.txtIDLetteraIntento.Value, sTabellaTestata
    
    oDoc.Field "Nom_bollo_esente", chkAddebitaBollo.Value, sTabellaTestata
    oDoc.Field "Nom_IVA_default", Me.chkNomIva.Value, sTabellaTestata
    
    'Imposta la valuta se non specificata con la valuta nazionale
    'If fnNotNullN(oDoc.Field("Link_Val_valuta", , sTabellaTestata)) = 0 Then
    '    oDoc.Field "Link_Val_valuta", oDoc.DBDefaults.Link_Val_valuta_nazionale, sTabellaTestata
    '    oDoc.Field "Link_Val_cambio", Null, sTabellaTestata
    'End If


    
    oDoc.Field "Doc_data_inizio_trasporto", Me.txtDataTrasporto.Text, sTabellaTestata
    oDoc.Field "Doc_ora_inizio_trasporto", Me.txtOraTrasporto.Text, sTabellaTestata
    oDoc.Field "Doc_annotazioni_variazio", Me.txtAnnotazioni.Text, sTabellaTestata
    oDoc.Field "Link_Doc_aspetto_esteriore", Me.cboAspettoEsteriore.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_spedizione", Me.cboTrasporto.CurrentID, sTabellaTestata
    
    'DATI INTRA
    oDoc.Field "Doc_intra_cessione", Me.chkCessione.Value, sTabellaTestata
    oDoc.Field "Link_Doc_intra_dati_richiesti", Me.cboIntraDatiRichiesti.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_intra_sezione", Me.cboIntraSezDaComp.ListIndex, sTabellaTestata
    oDoc.Field "Link_Doc_intra_mese", Me.cboIntraMese.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_intra_trimestre", Me.cboIntraTrimestre.ListIndex, sTabellaTestata
    oDoc.Field "Doc_intra_anno", Me.txtIntraAnno.Value, sTabellaTestata
    oDoc.Field "Link_Doc_intra_modo_trasporto", Me.cboIntraTrasporto.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_intra_provinc_merce", Me.cboIntraProvincia.CurrentID, sTabellaTestata
    oDoc.Field "Link_Doc_intra_naz_pagamento", Me.cboIntraNazione.CurrentID, sTabellaTestata


    ''AGENTE
    oDoc.Field "Link_doc_agente", Me.CDAgenteTesta.KeyFieldID, sTabellaTestata
    oDoc.Field "Doc_age_ragione_sociale", Me.CDAgenteTesta.Code, sTabellaTestata
    oDoc.Field "Doc_age_nome", Me.CDAgenteTesta.Description, sTabellaTestata
    oDoc.Field "Doc_age_codice", GET_CODICE_AGENTE(Me.CDAgenteTesta.KeyFieldID), sTabellaTestata

    'fnRecuperaAnnotazioniPerDoc
    
    oDoc.Field "RV_POIDAnagraficaDestinazione", Me.ACSAnaDest.IDAnagrafica, sTabellaTestata

    oDoc.Field "Doc_data_plafond", Me.txtDataPlafond.Text, sTabellaTestata
    'ALTRI DATI
    oDoc.Field "RV_PODataCompetenzaLiq", DATA_COMPETENZA_LIQ, sTabellaTestata
    oDoc.Field "RV_POAnnotazioni1", ANNOTAZIONE_01, sTabellaTestata
    oDoc.Field "RV_POAnnotazioni2", ANNOTAZIONE_02, sTabellaTestata
    oDoc.Field "RV_POAnnotazioni3", ANNOTAZIONE_03, sTabellaTestata

    oDoc.Field "Doc_numero_vs_ordine_di_rifer", Me.txtNumeroOrdineCliente.Text, sTabellaTestata
    oDoc.Field "Doc_data_vs_ordine_di_rifer", Me.txtDataOrdineCliente.Text, sTabellaTestata
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
    If Not bLoading Then
        'Effettuta il calcolo del documento
        On Error Resume Next
        oDoc.PerformDocument Nothing
        
        'Aggiorna il contenuto delle listview delle scadenze e del castelletto Iva
        sbPopalaListaScadenze
        sbPopalaListaIva
    End If
    
    Screen.MousePointer = 0
End Sub

Public Sub ConnessioneDiamanteADO()
On Error GoTo ERR_ConnessioneDiamanteADO
    
    Set Cn = m_App.Database.Connection
    
Exit Sub
ERR_ConnessioneDiamanteADO:
    MsgBox Err.Description, vbCritical, "Connessione Diamante di tipo ADO"
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
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
If Me.SSTab1.Tab = 2 Then
    cmdNuovaCommissione_Click
    TOTALE_MERCE = GET_TOTALE_MERCE_DOCUMENTO
    Me.lblTotaleMercePerComm.Caption = "TOTALE MERCE DOCUMENTO: " & FormatNumber(TOTALE_MERCE, 2)
End If
End Sub

Private Function GET_TOTALE_MERCE_DOCUMENTO()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT SUM(Art_importo_totale_neutro) AS TotaleMerce "
sSQL = sSQL & "FROM " & sTabellaDettaglio
sSQL = sSQL & " WHERE (IDOggetto = " & oDoc.IDOggetto & ") "
'sSQL = sSQL & " AND (RV_POTipoRiga = 1)"


Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_TOTALE_MERCE_DOCUMENTO = 0
Else
    GET_TOTALE_MERCE_DOCUMENTO = fnNotNullN(rs!TotaleMerce)
End If

rs.CloseResultset
Set rs = Nothing

End Function



Private Sub txtAnnotazioni_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtCausaleDocumento_Change()
    sbImpostaDatiDocumento
End Sub







Private Sub txtCausaleDocumentoEF_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtCodiceSocio_LostFocus()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer

If Len(Me.txtCodiceSocio.Text) > 0 Then
    If fnNotNull(oDoc.Field("RV_POCodiceSocio", , sTabellaDettaglio)) <> Me.txtCodiceSocio.Text Then
    
        sSQL = "SELECT  Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.Cap, Comune.Comune, Provincia.Provincia, Anagrafica.Indirizzo, "
        sSQL = sSQL & "Fornitore.Codice "
        sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
        sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
        sSQL = sSQL & "Anagrafica INNER JOIN "
        sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica ON Comune.IDComune = Anagrafica.IDComune "
        sSQL = sSQL & "WHERE (IDAzienda=" & m_App.IDFirm & ") "
        sSQL = sSQL & "AND (Anagrafica.IDCategoriaAnagrafica=" & Link_TipoSocio & ") "
        sSQL = sSQL & "AND (Fornitore.Codice = " & fnNormString(Me.txtCodiceSocio.Text) & ") "
        
        Set rs = Cn.OpenResultset(sSQL, adOpenKeyset)
        
        If rs.EOF Then
            rs.CloseResultset
            Set rs = Nothing
            lblCodiceSocio_Click
        Else
            I = 0
            While Not rs.EOF
                I = I + 1
                If I > 1 Then
                    rs.MoveLast
                End If
            
            rs.MoveNext
            Wend
            
            If I > 1 Then
                rs.CloseResultset
                Set rs = Nothing
                lblCodiceSocio_Click
            Else
                rs.MoveFirst
                
                Me.txtCodiceSocio.Text = fnNotNull(rs!Codice)
                Link_Socio = fnNotNullN(rs!IDAnagrafica)
                Me.lblNomeSocio.Caption = fnNotNull(rs!Nome)
                Me.txtSocio.Text = fnNotNull(rs!Anagrafica)
                        
                rs.CloseResultset
                Set rs = Nothing
            End If
            
        End If
    End If
End If
End Sub



Private Sub txtColliTotali_LostFocus()
    If (txtColliTotali.Value = oDoc.Field("Tot_numero_colli", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub

Private Sub txtDataCompLiq_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtDataTrasporto_LostFocus()
    sbImpostaDatiDocumento
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

Private Sub txtDescrizioneArticolo_LostFocus()
    lblDocument_Click 24
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


sbImpostaDatiDocumento

Exit Sub
ERR_txtIDLetteraIntento_Change:
    MsgBox Err.Description, vbCritical, "txtIDLetteraIntento_Change"
End Sub
Private Sub txtImportoUnitarioArticolo_LostFocus()
    CalcolaImportoScontato
    CalcolaTotaleRiga
End Sub
Private Sub CalcolaImportoScontato()
    Me.txtImponibileUnitario.Value = Me.txtImportoUnitarioArticolo.Value - ((Me.txtImportoUnitarioArticolo.Value / 100)) * (Me.txtSconto1.Value)
    Me.txtImponibileUnitario.Value = Me.txtImponibileUnitario.Value - ((Me.txtImponibileUnitario.Value / 100)) * (Me.txtSconto2.Value)
End Sub
Private Sub CalcolaTotaleRiga()
Dim TOTALE_RIGA As Double
Dim TOTALE_IMPONIBILE_RIGA As Double

Me.txtTotaleRiga.Value = (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value)
End Sub
Private Sub txtIntraAnno_LostFocus()
    sbImpostaDatiDocumento
End Sub

Private Sub txtOraTrasporto_LostFocus()
    sbImpostaDatiDocumento
End Sub
Public Sub SalvataggioRiga()

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
Private Sub NuovaRigaDocumento()
        
    If oDoc.Tables(sTabellaDettaglio).NumRetails = 0 Then
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail 1
    Else
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
    End If

    
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
'    oDoc.Field "Art_Importo_netto_IVA", Me.txtImponibileArticolo.Value, sTabellaDettaglio
'    oDoc.Field "Art_importo_net_sconto_lor_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) + (((Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
'    oDoc.Field "Art_importo_net_sconto_net_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value), sTabellaDettaglio
'    oDoc.Field "Art_importo_sconto_lordo_IVA", ((Me.txtImponibileUnitario.Value / 100) * (Me.txtSconto1.Value + Me.txtSconto2.Value)), sTabellaDettaglio




    oDoc.Field "Art_numero_colli", Me.txtColli.Value, sTabellaDettaglio
    oDoc.Field "Art_Peso", Me.txtPesoLordo.Value, sTabellaDettaglio
    oDoc.Field "Art_tara", Me.txtTara.Value, sTabellaDettaglio
    oDoc.Field "Art_quantita_pezzi", Me.txtPezzi.Value, sTabellaDettaglio
    
    oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
    oDoc.Field "Link_art_IVA", Me.cboAliquotaArticolo.CurrentID, sTabellaDettaglio
    oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaArticolo.Value, sTabellaDettaglio
    
    oDoc.Field "Link_Art_unita_di_misura", Me.cboUnitaDiMisura.CurrentID, sTabellaDettaglio
    oDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(Me.cboUnitaDiMisura.CurrentID), sTabellaDettaglio
    
    oDoc.Field "RV_POLinkRiga", NumeroRiga, sTabellaDettaglio
    oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
    
    oDoc.Field "RV_PODataConferimento", Me.txtDataConferimento.Text, sTabellaDettaglio
    oDoc.Field "RV_POIDConferimentoRighe", Link_RigaConferimento, sTabellaDettaglio
    oDoc.Field "RV_POIDSocio", Link_Socio, sTabellaDettaglio
    oDoc.Field "RV_POCodiceSocio", Me.txtCodiceSocio.Text, sTabellaDettaglio
    oDoc.Field "RV_POSocio", Me.txtSocio.Text, sTabellaDettaglio
    oDoc.Field "RV_PONomeSocio", Me.lblNomeSocio.Caption, sTabellaDettaglio
    oDoc.Field "RV_POIDAnagraficaFatturazione", Me.CDSocioFatt.KeyFieldID, sTabellaDettaglio
    oDoc.Field "RV_POCodiceLotto", Me.txtCodiceLottoVendita.Text, sTabellaDettaglio
    oDoc.Field "RV_POImportoDaLiq", 0, sTabellaDettaglio
    oDoc.Field "RV_POImportoLiq", Me.txtImponibileUnitario.Value, sTabellaDettaglio
    oDoc.Field "RV_POQuantitaLiq", Me.txtQta_UM.Value * Moltiplicatore, sTabellaDettaglio
    oDoc.Field "RV_POPrezzoMedioInLiq", Me.chkPrezzoMedioInLiq.Value, sTabellaDettaglio
    oDoc.Field "RV_POImportoMerceNetta", Me.txtImponibileUnitario.Value, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoImballo", 0, sTabellaDettaglio
    oDoc.Field "RV_POIDIvaImballo", 0, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoImportoVenditaLiq", Me.cboTipoImportoLiq.CurrentID, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoDocumentoCoop", Me.cboTipoDocumentoCoop.CurrentID, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoManuale", Me.txtImportoLiqVarMan.Value, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoVariazione", Me.cboTipoVariazione.CurrentID, sTabellaDettaglio

    oDoc.Field "RV_PODataLavorazione", Me.txtDataLavorazione.Text, sTabellaDettaglio
    oDoc.Field "RV_PORigaRiscontroPeso", Me.chkRiscontroPeso.Value, sTabellaDettaglio
    

    If Link_RigaConferimento > 0 Then
        oDoc.Field "RV_POQuantitaOrigine", Me.txtQuantitaOriginale.Value, sTabellaDettaglio
        oDoc.Field "RV_POPrezzoUnitarioOrigine", Me.txtPrezzoOriginale.Value, sTabellaDettaglio
    Else
        oDoc.Field "RV_POQuantitaOrigine", 0, sTabellaDettaglio
        oDoc.Field "RV_POPrezzoUnitarioOrigine", 0, sTabellaDettaglio
    End If

    oDoc.Field "Link_Art_IDCContropartita", Me.txtIDPianoDeiConti.Value, sTabellaDettaglio
    oDoc.Field "Art_CContropartita_codifica", Me.txtCodiceConto.Text, sTabellaDettaglio
    oDoc.Field "Art_CContropartita_descrizione", Me.txtDescrizioneConto.Text, sTabellaDettaglio


    'INTRA
    oDoc.Field "Art_intra_non_riporto", Me.chkRiportoIntra_Art.Value, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_nomenclatura", Me.CDIntra_art_Nom_Comb.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_natura_trans", Me.CDIntra_art_Nat_Trans.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Art_intra_qta_tot_massa_netta", GET_QUANTITA_MASSA_NETTA_INSTRASTAT(Me.CDArticolo.KeyFieldID, TheApp.IDFirm, Me.txtQta_UM.Value), sTabellaDettaglio


    '''AGENTE
    oDoc.Field "Link_Art_agente", Me.CDAgenteRiga.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Art_age_nome", Me.CDAgenteRiga.Description, sTabellaDettaglio
    oDoc.Field "Art_age_codice", GET_LINK_CODICE_AGE(Me.CDAgenteRiga.KeyFieldID, TheApp.IDFirm), sTabellaDettaglio
    oDoc.Field "Art_age_ragione_sociale", Me.CDAgenteRiga.Code, sTabellaDettaglio
    
    oDoc.Field "Link_Art_age_regola_provv", Me.cboRegolaProvvigione.CurrentID, sTabellaDettaglio
    oDoc.Field "Art_age_regola_provv", Me.cboRegolaProvvigione.Text, sTabellaDettaglio
    oDoc.Field "Link_Art_age_tipo_ordine", Me.cboTipoOrdine.CurrentID, sTabellaDettaglio

    If Me.txtPercProvv.Value > 0 Then
        oDoc.Field "Art_age_percentuale_provv", Me.txtPercProvv.Value, sTabellaDettaglio
    End If
    If Me.txtImportoProvv.Value > 0 Then
        oDoc.Field "Art_age_importo_provv", Me.txtImportoProvv.Value, sTabellaDettaglio
    End If

    'NumeroProgSingolaRiga = NumeroProgSingolaRiga + 1
    oDoc.Field "ID_Art_dettaglio_prog", oDoc.SetIDArtDettaglioProg, sTabellaDettaglio
    oDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(Me.CDArticolo.KeyFieldID, Me.cdAnagrafica.KeyFieldID, Me.cboAltroSito.CurrentID), sTabellaDettaglio
    sbLoadElectronicInvoiceData4Article fnNotNullN(oDoc.Field("ID_Art_dettaglio_prog", , sTabellaDettaglio)), fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))


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

    oDoc.Tables(sTabellaDettaglio).SetActiveRetail lvwArticoli.SelectedItem.Index
    oDoc.Field "Link_Art_articolo", Me.CDArticolo.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Art_codice", Me.CDArticolo.Code, sTabellaDettaglio
    oDoc.Field "Art_descrizione", Me.txtDescrizioneArticolo.Text, sTabellaDettaglio
    
    If Link_RigaConferimento > 0 Then
        oDoc.Field "Art_peso", (oDoc.Field("Art_peso", , sTabellaDettaglio) / oDoc.Field("Art_quantita_totale", , sTabellaDettaglio)) * Me.txtQta_UM.Value, sTabellaDettaglio
    End If
    
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
'    oDoc.Field "Art_Importo_netto_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value), sTabellaDettaglio
'    oDoc.Field "Art_importo_net_sconto_lor_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) + (((Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value) / 100) * Me.txtAliquotaArticolo.Value), sTabellaDettaglio
'    oDoc.Field "Art_importo_net_sconto_net_IVA", (Me.txtImponibileUnitario.Value * Me.txtQta_UM.Value), sTabellaDettaglio
'    oDoc.Field "Art_importo_sconto_lordo_IVA", ((Me.txtImponibileUnitario.Value / 100) * (Me.txtSconto1.Value + Me.txtSconto2.Value)), sTabellaDettaglio


    oDoc.Field "Art_numero_colli", Me.txtColli.Value, sTabellaDettaglio
    oDoc.Field "Art_Peso", Me.txtPesoLordo.Value, sTabellaDettaglio
    oDoc.Field "Art_tara", Me.txtTara.Value, sTabellaDettaglio
    oDoc.Field "Art_quantita_pezzi", Me.txtPezzi.Value, sTabellaDettaglio
    
    oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
    oDoc.Field "Link_art_IVA", Me.cboAliquotaArticolo.CurrentID, sTabellaDettaglio
    oDoc.Field "Art_aliquota_IVA", Me.txtAliquotaArticolo.Value, sTabellaDettaglio
    
    oDoc.Field "Link_Art_unita_di_misura", Me.cboUnitaDiMisura.CurrentID, sTabellaDettaglio
    oDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(Me.cboUnitaDiMisura.CurrentID), sTabellaDettaglio
    
    oDoc.Field "RV_POLinkRiga", NumeroRiga, sTabellaDettaglio
    oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
    
    'oDoc.Field "RV_PODescrizioneDocumento", "Rif. D.d.t. n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
    
    oDoc.Field "RV_PODataConferimento", Me.txtDataConferimento.Text, sTabellaDettaglio
    oDoc.Field "RV_POIDConferimentoRighe", Link_RigaConferimento, sTabellaDettaglio
    oDoc.Field "RV_POIDSocio", Link_Socio, sTabellaDettaglio
    oDoc.Field "RV_POCodiceSocio", Me.txtCodiceSocio.Text, sTabellaDettaglio
    oDoc.Field "RV_POSocio", Me.txtSocio.Text, sTabellaDettaglio
    oDoc.Field "RV_PONomeSocio", Me.lblNomeSocio.Caption, sTabellaDettaglio
    oDoc.Field "RV_POIDAnagraficaFatturazione", Me.CDSocioFatt.KeyFieldID, sTabellaDettaglio

    oDoc.Field "RV_POCodiceLotto", Me.txtCodiceLottoVendita.Text, sTabellaDettaglio
    
    oDoc.Field "RV_POImportoDaLiq", 0, sTabellaDettaglio
    oDoc.Field "RV_POQuantitaLiq", Me.txtQta_UM.Value * Moltiplicatore, sTabellaDettaglio
    oDoc.Field "RV_POPrezzoMedioInLiq", Me.chkPrezzoMedioInLiq.Value, sTabellaDettaglio
    oDoc.Field "RV_POImportoMerceNetta", Me.txtImponibileUnitario.Value, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoImballo", 0, sTabellaDettaglio
    oDoc.Field "RV_POIDIvaImballo", 0, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoImportoVenditaLiq", Me.cboTipoImportoLiq.CurrentID, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoDocumentoCoop", Me.cboTipoDocumentoCoop.CurrentID, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoManuale", Me.txtImportoLiqVarMan.Value, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoVariazione", Me.cboTipoVariazione.CurrentID, sTabellaDettaglio
    
    oDoc.Field "RV_PODataLavorazione", Me.txtDataLavorazione.Text, sTabellaDettaglio
    oDoc.Field "RV_PORigaRiscontroPeso", Me.chkRiscontroPeso.Value, sTabellaDettaglio
    
    If Link_RigaConferimento > 0 Then
        oDoc.Field "RV_POQuantitaOrigine", Me.txtQuantitaOriginale.Value, sTabellaDettaglio
        oDoc.Field "RV_POPrezzoUnitarioOrigine", Me.txtPrezzoOriginale.Value, sTabellaDettaglio
        
    Else
        
        oDoc.Field "RV_POQuantitaOrigine", 0, sTabellaDettaglio
        oDoc.Field "RV_POPrezzoUnitarioOrigine", 0, sTabellaDettaglio
    End If


    oDoc.Field "Link_Art_IDCContropartita", Me.txtIDPianoDeiConti.Value, sTabellaDettaglio
    oDoc.Field "Art_CContropartita_codifica", Me.txtCodiceConto.Text, sTabellaDettaglio
    oDoc.Field "Art_CContropartita_descrizione", Me.txtDescrizioneConto.Text, sTabellaDettaglio


    'INTRA
    oDoc.Field "Art_intra_non_riporto", Me.chkRiportoIntra_Art.Value, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_nomenclatura", Me.CDIntra_art_Nom_Comb.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_natura_trans", Me.CDIntra_art_Nat_Trans.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Art_intra_qta_tot_massa_netta", GET_QUANTITA_MASSA_NETTA_INSTRASTAT(Me.CDArticolo.KeyFieldID, TheApp.IDFirm, Me.txtQta_UM.Value), sTabellaDettaglio


    '''AGENTE
    oDoc.Field "Link_Art_agente", Me.CDAgenteRiga.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Art_age_nome", Me.CDAgenteRiga.Description, sTabellaDettaglio
    oDoc.Field "Art_age_codice", GET_LINK_CODICE_AGE(Me.CDAgenteRiga.KeyFieldID, TheApp.IDFirm), sTabellaDettaglio
    oDoc.Field "Art_age_ragione_sociale", Me.CDAgenteRiga.Code, sTabellaDettaglio
    
    oDoc.Field "Link_Art_age_regola_provv", Me.cboRegolaProvvigione.CurrentID, sTabellaDettaglio
    oDoc.Field "Art_age_regola_provv", Me.cboRegolaProvvigione.Text, sTabellaDettaglio
    oDoc.Field "Link_Art_age_tipo_ordine", Me.cboTipoOrdine.CurrentID, sTabellaDettaglio
    
    If Me.txtPercProvv.Value > 0 Then
        oDoc.Field "Art_age_percentuale_provv", Me.txtPercProvv.Value, sTabellaDettaglio
    End If
    If Me.txtImportoProvv.Value > 0 Then
        oDoc.Field "Art_age_importo_provv", Me.txtImportoProvv.Value, sTabellaDettaglio
    End If
    
    oDoc.Field "Art_riferimento_PA", Rif_PA_Riga_Doc_Merce, sTabellaDettaglio
    oDoc.PerformTable sTabellaDettaglio, True


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

Private Function DispLotto() As Double

End Function
Private Function CalcolaDisponibiltaColli()


End Function
Private Function DispLottoOnLine(IDLotto As Long, DispLottoDiamante As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Integer
sSQL = "SELECT Qta_UM, SegnoDispLotto FROM RV_POTMPQtaLottoDocumento "
sSQL = sSQL & "WHERE IDLottoArticolo=" & IDLotto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = True Then
    DispLottoOnLine = 0
Else
    While Not rs.EOF
       Select Case fnNotNull(rs!SegnoDispLotto)
        
        Case "-"
            DispLottoDiamante = DispLottoDiamante - fnNotNullN(rs!Qta_UM)
        Case "+"
            DispLottoDiamante = DispLottoDiamante + fnNotNullN(rs!Qta_UM)
        Case ""
            DispLottoDiamante = DispLottoDiamante
       End Select
    rs.MoveNext
    Wend
End If
rs.CloseResultset
Set rs = Nothing
DispLottoOnLine = DispLottoDiamante
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

Private Sub fnDeleteTabellaRicorsione(IDUtente As Long, IDTipoOggetto As Long)
On Error GoTo ERR_fnDeleteTabellaRicorsione
Dim sSQL As String
    
    sSQL = "DELETE FROM TabellaRicorsione "
    sSQL = sSQL & "WHERE IDUtente=" & IDUtente
    'sSQL = sSQL & " WHERE IDFunzione=" & oDoc.IDOggetto
    Cn.Execute sSQL
    
    sSQL = "DELETE FROM TabellaRicorsione2 "
    'sSQL = sSQL & "WHERE IDPath1=" & oDoc.IDOggetto
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
End If
oRes.CloseResultset
Set oRes = Nothing
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
        sSQL = sSQL & "RV_POCaricoMerceTesta.Nome, RV_POCaricoMerceRighe.CodiceLotto, RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
        'sSQL = sSQL & "RV_POCaricoMerceTesta.DataConferimento, RV_POCaricoMerceTesta.IDRV_POTipoDocumentoCoop "
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
        Testo = "Processo di IV° Gamma numero " & GET_NUMERO_PROCESSO(IDProcessoIVGamma)
        GET_STRINGA_CONFERIMENTO = Testo
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
Dim sSQL As String
Dim OLDCursor As Long
Dim cl As dgColumnHeader


sSQL = "SELECT RV_POCommissioniPerDoc.IDRV_POCommissioniPerDoc, RV_POCommissioniPerDoc.IDOggetto, "
sSQL = sSQL & "RV_POCommissioniPerDoc.IDRV_POTipoCommissione, RV_POCommissioniPerDoc.Percentuale, RV_POCommissioniPerDoc.Importo, "
sSQL = sSQL & "RV_POCommissioniPerDoc.ImportoRiga , RV_POTipoCommissione.TipoCommissione "
sSQL = sSQL & "FROM RV_POCommissioniPerDoc LEFT OUTER JOIN "
sSQL = sSQL & "RV_POTipoCommissione ON RV_POCommissioniPerDoc.IDRV_POTipoCommissione = RV_POTipoCommissione.IDRV_POTipoCommissione "
sSQL = sSQL & "WHERE RV_POCommissioniPerDoc.IDOggetto=" & oDoc.IDOggetto

    OLDCursor = Cn.CursorLocation
    Cn.CursorLocation = 3
        
        Set rsGriglia = Cn.OpenResultset(sSQL)
            'Set rsEvent = rsGriglia.Data
    
        
    
        With Me.GrigliaCommissioni.ColumnsHeader
            .Clear
                .Add "IDRV_POCommissioniPerDoc", "ID", dgInteger, False, 500, 0, True, True, False
                .Add "IDRV_POTipoCommissione", "IDTipo", dgInteger, False, 500, 0, True, True, False
                .Add "TipoCommissione", "Tipo commissione", dgchar, True, 4500, 0, True, True, False
                Set cl = .Add("Percentuale", "%", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                Set cl = .Add("Importo", "Importo", dgDouble, True, 1500, dgAlignRight)
                    cl.FormatOptions.FormatNumericRegionalSettings = False
                    cl.FormatOptions.UseFormatControlSettings = False
                    cl.FormatOptions.FormatNumericDecSep = ","
                    cl.FormatOptions.FormatNumericDecimals = 2
                    cl.FormatOptions.FormatNumericThousandSep = "."
                'Set cl = .Add("ImportoRiga", "Importo riga", dgDouble, True, 1500, dgAlignRight)
                '    cl.FormatOptions.FormatNumericRegionalSettings = False
                '    cl.FormatOptions.UseFormatControlSettings = False
                '    cl.FormatOptions.FormatNumericDecSep = ","
                '    cl.FormatOptions.FormatNumericDecimals = 2
                '    cl.FormatOptions.FormatNumericThousandSep = "."

            
            Set Me.GrigliaCommissioni.Recordset = rsGriglia.Data
            Me.GrigliaCommissioni.Refresh
        End With
    
    Cn.CursorLocation = OLDCursor
    If (rsGriglia.BOF And rsGriglia.EOF) Then
        cmdNuovaCommissione_Click
    End If
    Me.txtImportoTotaleCommissioni.Value = GET_TOTALE_COMMISSIONI
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
Private Function fnGestioneArticoli(IDTipoDocumentoCoop As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POOperazionePerDoc.GestioneArticoli "
sSQL = sSQL & "FROM RV_POOperazionePerDoc INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_POOperazionePerDoc.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE IDRV_PODocumentoCoop=" & IDTipoDocumentoCoop
sSQL = sSQL & " AND IDFiliale=" & m_App.Branch

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    fnGestioneArticoli = False
Else
    fnGestioneArticoli = fnNormBoolean(rs!GestioneArticoli)
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
Dim rs As DmtOleDbLib.adoResultset
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
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtCodiceSocio.Text), 0)
            Case 2 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtSocio.Text), 1)
            Case 3 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.lblNomeSocio.Caption), 1)
            Case 4 'Giorno conferimento

            Case 5 'Mese del conferimento
            
            Case 6 'Anno del conferimento
            
            Case 7 'Giorno lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("d", Me.dtData.Text)), 0)
            Case 8 'Mese lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("m", Me.dtData.Text)), 0)
            Case 9 'Anno lavorazione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("yyyy", Me.dtData.Text)), 0)
            Case 10 'calibro
            '    StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(Me.cboCalibro.Text)), 1)
            Case 11 'Tipo lavorazione
            '    StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(Me.cboTipoLavorazione.Text)), 1)
            Case 12 'Tipo categoria
            '    StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(Me.cboTipoCategoria.Text)), 1)
            Case 13 'Carattere speciale "_"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("_"), 1)
            Case 14 'Carattere speciale "-"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(rs!LottoComp)), 1)
            Case 15 'Stringa personalizzata
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(rs!Testo)), 1)
            Case 16 'Codice imballo
                'StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(Me.CDImballo.Code)), 1)
            Case 17 'Descrizione imballo
                'StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(Me.txtDescrizioneImballo.Text)), 1)
            Case 18 'Codice pedana
               ' StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull("")), 1)
            Case 19 'Codice articolo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(Me.CDArticolo.Code)), 1)
            Case 20 'Descrizione articolo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(Me.txtDescrizioneArticolo.Text)), 1)
            Case 22 'Numero della settimana
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("ww", Me.dtData.Text)), 0)
            Case 23 'giorno dell'anno
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", Me.dtData.Text)), 0)
            Case 24 'Lotto di campagna
                'StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtLottoCampagna.Text), 1)
            Case 25 'Lotto di conferimento
                StringaElaborata = StringaElaborata
        
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

Private Sub sbElaborazioneRighe()
On Error GoTo ERR_sbElaborazioneRighe
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Quantita_Riga As Double
Dim Quantita_Riga_Conferita As Double
Dim Quantita_Colli As Double
Dim Quantita_PesoLordo As Double
Dim Quantita_Tara As Double
Dim Quantita_PesoNetto As Double
Dim Quantita_Pezzi As Double
Dim TotaleQuantitaColliEla As Long

Screen.MousePointer = 11

sSQL = "SELECT * FROM RV_POTMPArticoliNotaDiCredito "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Selezionato=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

TotaleQuantitaColliEla = 0

While Not rs.EOF
    Quantita_Riga = FormatNumber(((Quantita_da_accreditare / Quantita_Totale_Selezionata) * rs!Quantita), 5)
    Quantita_Riga_Conferita = FormatNumber((fnNotNullN(rs!QuantitaConferita) / Quantita_Totale_Selezionata) * Quantita_Riga)
    
    Quantita_Riga = FormatNumber(((Quantita_da_accreditare / Quantita_Totale_Selezionata) * rs!Quantita), 5)
    
    If Colli_Totali_Selezionati > 0 Then
        Quantita_Colli = FormatNumber(((Colli_da_accreditare / Colli_Totali_Selezionati) * rs!Colli), 0)
    Else
        Quantita_Colli = 0
    End If
    
    If (TotaleQuantitaColliEla + Quantita_Colli) > Colli_da_accreditare Then
        Quantita_Colli = Colli_da_accreditare - TotaleQuantitaColliEla
    Else
        Quantita_Colli = Quantita_Colli
    End If
    
    TotaleQuantitaColliEla = TotaleQuantitaColliEla + Quantita_Colli
    
    If PesoLordo_Totali_Selezionati > 0 Then
        Quantita_PesoLordo = FormatNumber(((PesoLordo_da_accreditare / PesoLordo_Totali_Selezionati) * rs!PesoLordo), 5)
    Else
        Quantita_PesoLordo = 0
    End If
    
    If Tara_Totali_Selezionati > 0 Then
        Quantita_Tara = FormatNumber(((Tara_da_accreditare / Tara_Totali_Selezionati) * rs!Tara), 5)
    Else
        Quantita_Tara = 0
    End If
    
    If PesoNetto_Totali_Selezionati > 0 Then
        Quantita_PesoNetto = FormatNumber(((PesoNetto_da_accreditare / PesoNetto_Totali_Selezionati) * rs!PesoNetto), 5)
    Else
        Quantita_PesoNetto = 0
    End If
    
    If Pezzi_Totali_Selezionati > 0 Then
        Quantita_Pezzi = FormatNumber(((Pezzi_da_accreditare / Pezzi_Totali_Selezionati) * rs!Pezzi), 5)
    Else
        Quantita_Pezzi = 0
    End If
    
    NuovaRigaDocumento_DaElaborazione rs, Quantita_Riga, rs!Quantita, Quantita_Riga_Conferita, Quantita_Colli, Quantita_PesoLordo, Quantita_Tara, Quantita_PesoNetto, Quantita_Pezzi

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Screen.MousePointer = 0

oDoc.PerformTable sTabellaDettaglio, True

'Aggiorna il contenuto della listview degli articoli
sbPopalaListaArticoli False
'Ricalcola il documento
sbCalcolaDocumento
Exit Sub
ERR_sbElaborazioneRighe:
    MsgBox Err.Description, vbCritical, "sbElaborazioneRighe"
    Screen.MousePointer = 0
End Sub
Private Sub NuovaRigaDocumento_DaElaborazione(rs As DmtOleDbLib.adoResultset, Quantita_Riga As Double, QuantitaOriginale As Double, QuantitaConferita As Double, Colli As Double, PesoLordo As Double, Tara As Double, PesoNetto As Double, Pezzi As Double)
Dim ImportoOriginale As Double
Dim Moltiplicatore_Local As Double
Dim IDUMCoop_Local As Long
Dim ImportoScontato As Double

ImportoOriginale = 0

    If oDoc.Tables(sTabellaDettaglio).NumRetails = 0 Then
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails
    Else
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
    End If
    
    oDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDArticolo), sTabellaDettaglio
    oDoc.Field "Art_codice", rs!CodiceArticolo, sTabellaDettaglio
    oDoc.Field "Art_descrizione", rs!DescrizioneArticolo, sTabellaDettaglio
    
    If (rs!RigaRiscontroPeso) = 1 Then
        If (Quantita_Riga) < 0 Then
            Quantita_Riga = Abs(Quantita_Riga)
            Importo_da_accreditare = -1 * Importo_da_accreditare
        End If
    End If
    
    
    oDoc.Field "Art_quantita_totale", Quantita_Riga, sTabellaDettaglio
    oDoc.Field "Art_numero_colli", Colli, sTabellaDettaglio
    oDoc.Field "Art_Peso", PesoLordo, sTabellaDettaglio
    oDoc.Field "Art_tara", Tara, sTabellaDettaglio
    oDoc.Field "Art_quantita_pezzi", Pezzi, sTabellaDettaglio
    
    oDoc.Field "Art_prezzo_unitario_neutro", Importo_da_accreditare, sTabellaDettaglio
    oDoc.Field "Art_sco_in_percentuale_1", rs!Sconto1, sTabellaDettaglio
    oDoc.Field "Art_sco_in_percentuale_2", rs!Sconto2, sTabellaDettaglio
    
    ImportoScontato = Importo_da_accreditare - ((Importo_da_accreditare / 100)) * fnNotNullN(rs!Sconto1)
    ImportoScontato = ImportoScontato - ((ImportoScontato / 100)) * fnNotNullN(rs!Sconto2)
    
    
    
    oDoc.Field "Art_importo_totale_netto_IVA", (ImportoScontato * Quantita_Riga), sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_netto_IVA", ImportoScontato, sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_lordo_IVA", ImportoScontato + ((ImportoScontato / 100) * rs!AliquotaIva), sTabellaDettaglio
    oDoc.Field "Art_Importo_totale_neutro", (ImportoScontato * Quantita_Riga), sTabellaDettaglio
    
    oDoc.Field "Art_Importo_netto_IVA", Importo_da_accreditare, sTabellaDettaglio
    
    oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
    oDoc.Field "Link_art_IVA", rs!IDIvaVendita, sTabellaDettaglio
    oDoc.Field "Art_aliquota_IVA", rs!AliquotaIva, sTabellaDettaglio
    
    oDoc.Field "Link_Art_unita_di_misura", rs!IDUnitaDiMisura, sTabellaDettaglio
    oDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(fnNotNullN(rs!IDUnitaDiMisura)), sTabellaDettaglio
    
'    oDoc.Field "Art_intra_non_riporto", Me.chkRiportoIntra_Art.Value, sTabellaDettaglio
'    oDoc.Field "Link_Art_intra_nomenclatura", Me.CDIntra_art_Nom_Comb.KeyFieldID, sTabellaDettaglio
'    oDoc.Field "Link_Art_intra_natura_trans", Me.CDIntra_art_Nat_Trans.KeyFieldID, sTabellaDettaglio
'    oDoc.Field "Art_intra_qta_tot_massa_netta", Me.txtIntra_Art_MassaNetta.Value, sTabellaDettaglio
    
    If (rs!RigaRiscontroPeso) = 1 Then
        If (Quantita_Riga) < 0 Then
            Importo_da_accreditare = Abs(Importo_da_accreditare)
            Quantita_Riga = -Quantita_Riga
        End If
    End If
    
    'oDoc.Field "RV_POLinkRiga", NumeroRiga, sTabellaDettaglio
    'oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
    'oDoc.Field "RV_PODescrizioneDocumento", "Rif. D.d.t. n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
    
    oDoc.Field "RV_PODescrizioneDocumento", rs!DescrizioneDocumento, sTabellaDettaglio
    oDoc.Field "RV_PODataConferimento", rs!DataConferimento, sTabellaDettaglio
    oDoc.Field "RV_POIDConferimentoRighe", rs!IDCollegamentoConferimento, sTabellaDettaglio
    oDoc.Field "RV_POIDAssegnazioneMerce", rs!IDCollegamentoAssegnazioneMerce, sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoIVGamma", rs!IDCollegamentoProcessoLavorazione, sTabellaDettaglio
    oDoc.Field "RV_POIDSocio", fnNotNullN(rs!IDSocio), sTabellaDettaglio
    oDoc.Field "RV_POCodiceSocio", fnNotNull(rs!CodiceSocio), sTabellaDettaglio
    oDoc.Field "RV_POSocio", fnNotNull(rs!Socio), sTabellaDettaglio
    oDoc.Field "RV_PONomeSocio", fnNotNull(rs!NomeSocio), sTabellaDettaglio
    oDoc.Field "RV_POCodiceLotto", fnNotNull(rs!CodiceLotto), sTabellaDettaglio
    oDoc.Field "RV_POIDAnagraficaFatturazione", fnNotNullN(rs!IDAnagraficaFatturazione), sTabellaDettaglio
    
    oDoc.Field "RV_POIDTipoOggetto", rs!IDTipoOggetto, sTabellaDettaglio
    oDoc.Field "RV_POIDOggetto", rs!IDOggetto, sTabellaDettaglio
    oDoc.Field "RV_POIDValoriOggettoDettaglio", rs!IDValoriOggettoDettaglio, sTabellaDettaglio
    oDoc.Field "RV_POIDOggetto_Collegato", rs!IDOggettoCollegato, sTabellaDettaglio
        
    oDoc.Field "RV_POIDTipoLavorazione", rs!IDTipoLavorazione, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoCategoria", rs!IDTipoCategoria, sTabellaDettaglio
    oDoc.Field "RV_POIDCalibro", rs!IDcalibro, sTabellaDettaglio
    oDoc.Field "RV_POQuantitaOrigine", QuantitaOriginale, sTabellaDettaglio
    ImportoOriginale = GET_RIGA_IMPORTO_ORIGINALE(fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), 0, fnNotNull(rs!CodiceLotto))
    oDoc.Field "RV_POPrezzoUnitarioOrigine", ImportoOriginale, sTabellaDettaglio
    
    oDoc.Field "RV_POImportoLiq", Importo_da_accreditare, sTabellaDettaglio
    oDoc.Field "RV_POImportoDaLiq", 0, sTabellaDettaglio
    Moltiplicatore_Local = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!IDArticolo))
    oDoc.Field "RV_POQuantitaLiq", Quantita_Riga * Moltiplicatore_Local, sTabellaDettaglio
    IDUMCoop_Local = GET_UM_COOP_ARTICOLO(fnNotNullN(rs!IDArticolo))
    If (IDUMCoop_Local > 0) Then
        Select Case (IDUMCoop_Local)
            Case 1
                oDoc.Field "RV_POQuantitaLiq", Colli * Moltiplicatore_Local, sTabellaDettaglio
            Case 2
                oDoc.Field "RV_POQuantitaLiq", PesoLordo * Moltiplicatore_Local, sTabellaDettaglio
            Case 3
                oDoc.Field "RV_POQuantitaLiq", Tara * Moltiplicatore_Local, sTabellaDettaglio
            Case 4
                oDoc.Field "RV_POQuantitaLiq", PesoNetto * Moltiplicatore_Local, sTabellaDettaglio
            Case 5
                oDoc.Field "RV_POQuantitaLiq", Pezzi * Moltiplicatore_Local, sTabellaDettaglio
        End Select
    End If
    oDoc.Field "RV_POPrezzoMedioInLiq", GET_VALORE_PREZZO_MEDIO_RIGA_ORIGINALE(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNull(rs!CodiceLotto)), sTabellaDettaglio
    oDoc.Field "RV_POImportoMerceNetta", Importo_da_accreditare, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoImballo", 0, sTabellaDettaglio
    oDoc.Field "RV_POIDIvaImballo", 0, sTabellaDettaglio
    
    'oDoc.Field "RV_POIDTipoVariazione", GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO(ImportoOriginale, QuantitaOriginale, Importo_da_accreditare, Quantita_Riga), sTabellaDettaglio
    oDoc.Field "RV_PORigaRiscontroPeso", fnNotNullN(rs!RigaRiscontroPeso), sTabellaDettaglio
    If (TIPO_VARIAZIONE_DA_WIZARD = 0) Then
        oDoc.Field "RV_POIDTipoVariazione", GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO(ImportoOriginale, QuantitaOriginale, Importo_da_accreditare, Quantita_Riga), sTabellaDettaglio
    Else
        oDoc.Field "RV_POIDTipoVariazione", TIPO_VARIAZIONE_DA_WIZARD, sTabellaDettaglio
    End If
    
    oDoc.Field "RV_POIDTipoImportoVenditaLiq", GET_VALORE_FORZA_TIPO_PREZZO_LIQ(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNull(rs!CodiceLotto)), sTabellaDettaglio
    
    'INSERIRE LA PEDANA E IL PESO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    oDoc.Field "RV_PODataLavorazione", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_PODataLavorazione", True), sTabellaDettaglio
    oDoc.Field "RV_POCodicePedana", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_POCodicePedana", True), sTabellaDettaglio
    oDoc.Field "RV_POIDTipoPedana", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_POIDTipoPedana", False), sTabellaDettaglio
    oDoc.Field "RV_POPesoPedana", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_POPesoPedana", False), sTabellaDettaglio
    oDoc.Field "RV_POLottoCampagna", GET_LOTTO_PRODUZIONE(fnNotNullN(rs!IDCollegamentoConferimento)), sTabellaDettaglio
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    oDoc.Field "RV_POID_Art_dettaglio_prog", fnNotNullN(rs!ID_Art_dettaglio_prog), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazione", fnNotNullN(rs!IDRV_POProcessoLavorazione), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazioneRighe", fnNotNullN(rs!IDRV_POProcessoLavorazioneRighe), sTabellaDettaglio
    oDoc.Field "RV_POIDLineaProduzione", fnNotNullN(rs!IDRV_POLineaProduzione), sTabellaDettaglio
    oDoc.Field "RV_POIDTipoUtilizzoLinea", fnNotNullN(rs!IDRV_POTipoUtilizzoLinea), sTabellaDettaglio

    GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDArticolo), TheApp.IDFirm, TheApp.Branch, oDoc.Field("Art_quantita_totale", , sTabellaDettaglio)
    SCRIVI_AGENTE_DA_RIGA_DOC fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio)
    SCRIVI_PDC_DA_RIGA_DOC fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio)

    NumeroProgSingolaRiga = NumeroProgSingolaRiga + 1
    
    oDoc.Field "RV_POLinkRiga", NumeroProgSingolaRiga, sTabellaDettaglio
    oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
    oDoc.Field "RV_POID_Art_dettaglio_prog", fnNotNullN(rs!ID_Art_dettaglio_prog), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazione", fnNotNullN(rs!IDRV_POProcessoLavorazione), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazioneRighe", fnNotNullN(rs!IDRV_POProcessoLavorazioneRighe), sTabellaDettaglio
    oDoc.Field "RV_POIDLineaProduzione", fnNotNullN(rs!IDRV_POLineaProduzione), sTabellaDettaglio
    oDoc.Field "RV_POIDTipoUtilizzoLinea", fnNotNullN(rs!IDRV_POTipoUtilizzoLinea), sTabellaDettaglio
    
    oDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(Me.CDArticolo.KeyFieldID, Me.cdAnagrafica.KeyFieldID, Me.cboAltroSito.CurrentID), sTabellaDettaglio
    sbLoadElectronicInvoiceData4Article fnNotNullN(oDoc.Field("ID_Art_dettaglio_prog", , sTabellaDettaglio)), fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))
    SET_INTRASTAT_RIGA_DOCUMENTO fnNotNullN(rs!IDArticolo), Quantita_Riga
        
End Sub
Private Function GET_VALORE_PREZZO_MEDIO_RIGA_ORIGINALE(IDTipoOggetto As Long, IDOggetto As Long, CodiceLottoVandita As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POPrezzoMedioInLiq "
Select Case IDTipoOggetto
    
    Case 114
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
    Case 2
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
End Select

sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVandita)
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_VALORE_PREZZO_MEDIO_RIGA_ORIGINALE = 0
Else
    GET_VALORE_PREZZO_MEDIO_RIGA_ORIGINALE = Abs(fnNotNullN(rs!RV_POPrezzoMedioInLiq))
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


Set rs = Cn.OpenResultset(sSQL)
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

Private Function GET_NUMERO_COPIE() As Long

End Function
Private Function GET_ORIENTAMENTO() As Long

End Function
Private Function GET_DOCUMENTO_DI_RIFERIMENTO() As String
On Error GoTo ERR_GET_DOCUMENTO_DI_RIFERIMENTO
Dim IDOggettoCollegato As Long
Dim IDTipoOggetto As Long
Dim IDValoriOggettoDettaglio As Long
Dim CodiceLottoVendita As String

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

IDOggettoCollegato = fnNotNullN(oDoc.Field("RV_POIDOggetto", , sTabellaDettaglio))
IDTipoOggetto = fnNotNullN(oDoc.Field("RV_POIDTipoOggetto", , sTabellaDettaglio))
IDValoriOggettoDettaglio = fnNotNullN(oDoc.Field("RV_POIDValoriOggettoDettaglio", , sTabellaDettaglio))
CodiceLottoVendita = fnNotNull(oDoc.Field("RV_POCodiceLotto", , sTabellaDettaglio))

If IDTipoOggetto = 0 Then
    GET_DOCUMENTO_DI_RIFERIMENTO = ""
    Me.txtQuantitaOriginale.Value = 0
    Me.txtPrezzoOriginale.Value = 0
    Me.txtColliOriginali.Value = 0
    Me.txtPesoLordoOriginale.Value = 0
    Me.txtTaraOriginale.Value = 0
    Me.txtPezziOriginali.Value = 0
    Me.txtPesoNettoOriginale.Value = 0
    Exit Function
End If

Select Case IDTipoOggetto
    Case 114
        sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0072.Doc_prefisso "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0001 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0001.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0001.Link_Art_IVA "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0001.IDOggetto=" & IDOggettoCollegato
    Case 2
        sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0002.Doc_prefisso "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0004 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0004.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0004.Link_Art_IVA "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0004.IDOggetto=" & IDOggettoCollegato
    Case 8
        sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica, ValoriOggettoPerTipo0008.Doc_prefisso "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0034 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0034.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0034.Link_Art_IVA "
        sSQL = sSQL & "WHERE ValoriOggettoPerTipo0008.IDOggetto=" & IDOggettoCollegato
End Select

sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
sSQL = sSQL & " AND RV_POTipoRiga=1 "

If COLLEGAMENTO_NOTA_PER_LOTTO = 0 Then
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
End If

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DOCUMENTO_DI_RIFERIMENTO = ""
    Me.txtQuantitaOriginale.Value = 0
    Me.txtPrezzoOriginale.Value = 0
    Me.txtColliOriginali.Value = 0
    Me.txtPesoLordoOriginale.Value = 0
    Me.txtTaraOriginale.Value = 0
    Me.txtPezziOriginali.Value = 0
    Me.txtPesoNettoOriginale.Value = 0
    If IDValoriOggettoDettaglio > 0 Then
        GET_DOCUMENTO_DI_RIFERIMENTO = "!!! ERRORE COLLEGAMENTO DDT/FA INESISTENTE !!!"
    End If
Else
    
    Select Case IDTipoOggetto
        Case 114
            
            GET_DOCUMENTO_DI_RIFERIMENTO = "Rif. Fattura accompagnatoria n° " & Trim(fnNotNull(rs!Doc_prefisso)) & "/" & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_data)
        Case 2
            GET_DOCUMENTO_DI_RIFERIMENTO = "Rif. Documento di trasporto n° " & Trim(fnNotNull(rs!Doc_prefisso)) & "/" & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_data)
        Case 8
            GET_DOCUMENTO_DI_RIFERIMENTO = "Rif. Corrispettivo n° " & Trim(fnNotNull(rs!Doc_prefisso)) & "/" & fnNotNull(rs!Doc_Numero) & " del " & fnNotNull(rs!Doc_data)
    End Select
    
    
    Me.txtQuantitaOriginale.Value = fnNotNullN(rs!Art_quantita_totale)
    Me.txtPrezzoOriginale.Value = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    Me.txtColliOriginali.Value = fnNotNullN(rs!Art_numero_colli)
    Me.txtPesoLordoOriginale.Value = fnNotNullN(rs!Art_peso)
    Me.txtTaraOriginale.Value = fnNotNullN(rs!Art_tara)
    Me.txtPezziOriginali.Value = fnNotNullN(rs!Art_quantita_pezzi)
    Me.txtSconto1Ori.Value = fnNotNullN(rs!Art_sco_in_percentuale_1)
    Me.txtSconto2Ori.Value = fnNotNullN(rs!Art_sco_in_percentuale_2)
    Me.txtPesoNettoOriginale.Value = Me.txtPesoLordoOriginale.Value - Me.txtTaraOriginale.Value
    
End If

rs.CloseResultset
Set rs = Nothing

Exit Function

ERR_GET_DOCUMENTO_DI_RIFERIMENTO:
    MsgBox Err.Description, vbCritical, "GET_DOCUMENTO_DI_RIFERIMENTO"
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
    
    'On Error GoTo BarMenu_ClickError
        
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
Private Sub SendDocument(ByVal Appl As Long, Optional InvioEmail As Long = 1)
    On Error GoTo errHandler
    
    Dim OLDCursor As Integer
    Dim SExt As String
    Dim DataDocumento As String
    Dim NomeCartella As String
    Dim NomeFile As String
    Dim InvioEmailPersonalizzata As Boolean
    
    InvioEmailPersonalizzata = False
    
    If InvioEmail = 1 Then
        InvioEmailPersonalizzata = GET_INVIO_EMAIL_PERSONALIZZATA(TheApp.IDUser)
    End If
    
    OLDCursor = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
   
   
    Select Case Appl
        Case 0
            SExt = ".xls"
        Case 1
            SExt = ".doc"
        Case 2
            SExt = ".html"
        Case 3
            SExt = ".pdf"
    End Select
   
    DataDocumento = Replace(oDoc.DataEmissione, "/", "-")
    NomeCartella = TrovaCartella(CSIDL_COMMON_APPDATA)
    
    
    NomeFile = GET_NOME_FILE_DOCUMENTO '& SExt
    
    
    If (InvioEmail = 1) Then
        oReport.ShowExportFile = False
        If (InvioEmailPersonalizzata = False) Then
            NomeFile = NomeFile & SExt
        End If
    Else
        oReport.ForceOpenCmdDlg = True
    End If

    oReport.ExportFileName = NomeCartella & NomeFile
    
    oReport.Export Appl
    
    
    If (InvioEmail = 1) Then
        If (InvioEmailPersonalizzata = False) Then
            PREPARAZIONE_EMAIL oReport.ExportFileName, Me.cdAnagrafica.KeyFieldID, Me.cboAltroSito.CurrentID, Me.lngNumero.Value, Me.dtData.Text, oDoc.Descrizione
        Else
            oReport.SendMailTo recPDF, GET_INDIRIZZO_EMAIL_CLIENTE(Me.cdAnagrafica.KeyFieldID)
        End If
    End If
    
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

Private Sub AGGIORNA_MOVIMENTI_MAGAZZINO_CONFERIMENTI(IDTipoOggetto As Long, IDOggetto As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim QuantitaMovimento As Double
Dim Link_Unita_di_Misura_Conferimeto As Long
Dim TipoNota As Long

Screen.MousePointer = 11
Me.Caption = "AGGIORNAMENTO MOVIMENTI....................."
'Me.lblInfoTesta.Font.Bold = True
'Me.lblInfoTesta.ForeColor = vbBlue
DoEvents

Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection

sSQL = "SELECT " & sTabellaDettaglio & ".IDValoriOggettoDettaglio, "
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
sSQL = sSQL & sTabellaDettaglio & ".Link_art_Articolo, "
sSQL = sSQL & sTabellaDettaglio & ".Art_descrizione, "
sSQL = sSQL & sTabellaDettaglio & ".Art_prezzo_unitario_neutro, "
sSQL = sSQL & sTabellaDettaglio & ".Art_Quantita_totale, "
sSQL = sSQL & sTabellaDettaglio & ".Art_Importo_totale_neutro, "
sSQL = sSQL & sTabellaDettaglio & ".Link_Art_unita_di_misura, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDOggetto, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoOggetto, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDValoriOggettoDettaglio, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POQuantitaOrigine, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POPrezzoMedioInLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoMerceNetta, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDIvaImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POVariazionePrezzoImballo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoImballoInArticolo, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POPrezzoUnitarioOrigine, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoVariazione, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDAnagraficaFatturazione, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoImportoVenditaLiq, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POVariazionePrezzoManuale, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoDocumentoCoop, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POImportoRigaCommissioni, "
sSQL = sSQL & sTabellaDettaglio & ".RV_PODataLavorazione, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoLavorazione, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoCategoria, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDCalibro, "

sSQL = sSQL & sTabellaDettaglio & ".RV_POIDPedana, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDTipoPedana, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POCodicePedana, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POPesoPedana, "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDLottoCampagnaLavorazione, "
sSQL = sSQL & sTabellaDettaglio & ".Art_sco_in_percentuale_1, "
sSQL = sSQL & sTabellaDettaglio & ".Art_sco_in_percentuale_2, "

sSQL = sSQL & "RV_POCaricoMerceRighe.IDArticolo, RV_POCaricoMerceRighe.Articolo, RV_POCaricoMerceTesta.NumeroDocumento, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDUnitaDiMisura, RV_POCaricoMerceRighe.CodiceLotto, "
sSQL = sSQL & "RV_POCaricoMerceTesta.IDMagazzinoConferimento, RV_POCaricoMerceRighe.IDUnitaDiMisuraDiamante, "
sSQL = sSQL & "RV_POCaricoMerceRighe.IDRV_POTipoLavorazione AS IDTipoLavorazioneConf, RV_POCaricoMerceRighe.PrezzoMedio AS PrezzoMedioConf "
sSQL = sSQL & "FROM " & sTabellaDettaglio & " LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceRighe ON "
sSQL = sSQL & sTabellaDettaglio & ".RV_POIDConferimentoRighe = RV_POCaricoMerceRighe.IDRV_POCaricoMerceRighe LEFT OUTER JOIN "
sSQL = sSQL & "RV_POCaricoMerceTesta ON RV_POCaricoMerceRighe.IDRV_POCaricoMerceTesta = RV_POCaricoMerceTesta.IDRV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE " & sTabellaDettaglio & ".IDOggetto=" & oDoc.IDOggetto
'sSQL = sSQL & " AND " & sTabellaDettaglio & ".RV_POTipoRiga=1"

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection
    
While Not rs.EOF
    LINK_LOTTO_PROD_DA_LAV_MOV = fnNotNullN(rs!RV_POIDLottoCampagnaLavorazione)
    'If fnNotNullN(rs!RV_POIDConferimentoRighe) > 0 Then
    If ((fnNotNullN(rs!RV_POIDSocio) > 0) Or (fnNotNullN(rs!RV_POIDAssegnazioneMerce) > 0)) Then
        'TipoNota = GET_RIGA_A_QUANTITA(fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Art_prezzo_unitario_neutro), fnNotNull(rs!RV_POCodiceLotto))
        'Select Case TipoNota
        Select Case fnNotNullN(rs!RV_POIDTipoVariazione)
            Case 3 'VARIAZIONE TOTALE DI QUANTITA
                CARICA_MOVIMENTO_DOCUMENTO_DERIVATO rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNullN(rs!NumeroDocumento), _
                fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
                fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Link_art_articolo), _
                fnNotNull(rs!art_descrizione), fnNotNullN(rs!Link_Art_unita_di_misura), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Art_prezzo_unitario_neutro), _
                fnNotNullN(rs!Art_Importo_totale_neutro), fnNotNullN(rs!Art_sco_in_percentuale_1), fnNotNullN(rs!Art_sco_in_percentuale_2), 19, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaOrigine), fnNotNullN(rs!RV_POPrezzoUnitarioOrigine), fnNotNullN(rs!RV_POPrezzoMedioInLiq), _
                fnNotNullN(rs!RV_POIDTipoVariazione), fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POIDAnagraficaFatturazione), fnNotNullN(rs!RV_POIDTipoImportoVenditaLiq), _
                fnNotNullN(rs!RV_POIDTipoDocumentoCoop), fnNotNullN(rs!RV_POVariazionePrezzoManuale), fnNotNullN(rs!RV_POImportoRigaCommissioni), fnNotNull(rs!RV_PODataLavorazione), _
                fnNotNullN(rs!RV_POIDTipoLavorazione), fnNotNullN(rs!RV_POIDTipoCategoria), fnNotNullN(rs!RV_POIDCalibro), fnNotNullN(rs!IDTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), _
                fnNotNullN(rs!RV_POIDPedana), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNull(rs!RV_POCodicePedana), fnNotNullN(rs!RV_POPesoPedana)
                
                'GeneraMovimentoConferimento fnNotNullN(rs!IDMagazzinoConferimento), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!IDArticolo), fnNotNull(rs!Articolo), fnNotNullN(rs!IDUnitaDiMisuraDiamante), fnNotNullN(rs!Art_peso)
            
            Case 2 'VARIAZIONE PARZIALE DI QUANTITA
                CARICA_MOVIMENTO_DOCUMENTO_DERIVATO rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNullN(rs!NumeroDocumento), _
                fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
                fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Link_art_articolo), _
                fnNotNull(rs!art_descrizione), fnNotNullN(rs!Link_Art_unita_di_misura), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Art_prezzo_unitario_neutro), _
                fnNotNullN(rs!Art_Importo_totale_neutro), fnNotNullN(rs!Art_sco_in_percentuale_1), fnNotNullN(rs!Art_sco_in_percentuale_2), 19, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaOrigine), fnNotNullN(rs!RV_POPrezzoUnitarioOrigine), fnNotNullN(rs!RV_POPrezzoMedioInLiq), _
                fnNotNullN(rs!RV_POIDTipoVariazione), fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POIDAnagraficaFatturazione), fnNotNullN(rs!RV_POIDTipoImportoVenditaLiq), _
                fnNotNullN(rs!RV_POIDTipoDocumentoCoop), fnNotNullN(rs!RV_POVariazionePrezzoManuale), fnNotNullN(rs!RV_POImportoRigaCommissioni), fnNotNull(rs!RV_PODataLavorazione), _
                fnNotNullN(rs!RV_POIDTipoLavorazione), fnNotNullN(rs!RV_POIDTipoCategoria), fnNotNullN(rs!RV_POIDCalibro), fnNotNullN(rs!IDTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), _
                fnNotNullN(rs!RV_POIDPedana), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNull(rs!RV_POCodicePedana), fnNotNullN(rs!RV_POPesoPedana)
                
                'GeneraMovimentoConferimento fnNotNullN(rs!IDMagazzinoConferimento), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!IDArticolo), fnNotNull(rs!Articolo), fnNotNullN(rs!IDUnitaDiMisuraDiamante), fnNotNullN(rs!Art_peso)
            
            Case 4 'VARIAZIONE TOTALE DI VALORE
                CARICA_MOVIMENTO_DOCUMENTO_DERIVATO rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNullN(rs!NumeroDocumento), _
                fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
                fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), 0, fnNotNullN(rs!Link_art_articolo), _
                fnNotNull(rs!art_descrizione), fnNotNullN(rs!Link_Art_unita_di_misura), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Art_prezzo_unitario_neutro), _
                fnNotNullN(rs!Art_Importo_totale_neutro), fnNotNullN(rs!Art_sco_in_percentuale_1), fnNotNullN(rs!Art_sco_in_percentuale_2), 18, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaOrigine), fnNotNullN(rs!RV_POPrezzoUnitarioOrigine), fnNotNullN(rs!RV_POPrezzoMedioInLiq), _
                fnNotNullN(rs!RV_POIDTipoVariazione), fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POIDAnagraficaFatturazione), fnNotNullN(rs!RV_POIDTipoImportoVenditaLiq), _
                fnNotNullN(rs!RV_POIDTipoDocumentoCoop), fnNotNullN(rs!RV_POVariazionePrezzoManuale), fnNotNullN(rs!RV_POImportoRigaCommissioni), fnNotNull(rs!RV_PODataLavorazione), _
                fnNotNullN(rs!RV_POIDTipoLavorazione), fnNotNullN(rs!RV_POIDTipoCategoria), fnNotNullN(rs!RV_POIDCalibro), fnNotNullN(rs!IDTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), _
                fnNotNullN(rs!RV_POIDPedana), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNull(rs!RV_POCodicePedana), fnNotNullN(rs!RV_POPesoPedana)
        
            Case 1 'VARIAZIONE PARZIALE DI VALORE e DI QUANTITA'
                CARICA_MOVIMENTO_DOCUMENTO_DERIVATO rs!IDValoriOggettoDettaglio, fnNotNullN(rs!RV_POIDConferimentoRighe), fnNotNullN(rs!RV_POIDAssegnazioneMerce), fnNotNullN(rs!RV_POIDProcessoIVGamma), fnNotNullN(rs!RV_POIDSocio), fnNotNull(rs!RV_PODataConferimento), fnNotNullN(rs!NumeroDocumento), _
                fnNotNull(rs!CodiceLotto), fnNotNull(rs!RV_POLottoCampagna), fnNotNull(rs!RV_POCodiceLotto), _
                fnNotNullN(rs!RV_POQuantitaLiq), fnNotNullN(rs!RV_POImportoDaLiq), fnNotNullN(rs!RV_POImportoLiq), 0, fnNotNullN(rs!Link_art_articolo), _
                fnNotNull(rs!art_descrizione), fnNotNullN(rs!Link_Art_unita_di_misura), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Art_prezzo_unitario_neutro), _
                fnNotNullN(rs!Art_Importo_totale_neutro), fnNotNullN(rs!Art_sco_in_percentuale_1), fnNotNullN(rs!Art_sco_in_percentuale_2), 18, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNullN(rs!RV_POQuantitaOrigine), fnNotNullN(rs!RV_POPrezzoUnitarioOrigine), fnNotNullN(rs!RV_POPrezzoMedioInLiq), _
                fnNotNullN(rs!RV_POIDTipoVariazione), fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi), fnNotNullN(rs!RV_POIDAnagraficaFatturazione), fnNotNullN(rs!RV_POIDTipoImportoVenditaLiq), _
                fnNotNullN(rs!RV_POIDTipoDocumentoCoop), fnNotNullN(rs!RV_POVariazionePrezzoManuale), fnNotNullN(rs!RV_POImportoRigaCommissioni), fnNotNull(rs!RV_PODataLavorazione), _
                fnNotNullN(rs!RV_POIDTipoLavorazione), fnNotNullN(rs!RV_POIDTipoCategoria), fnNotNullN(rs!RV_POIDCalibro), fnNotNullN(rs!IDTipoLavorazioneConf), fnNotNullN(rs!PrezzoMedioConf), _
                fnNotNullN(rs!RV_POIDPedana), fnNotNullN(rs!RV_POIDTipoPedana), fnNotNull(rs!RV_POCodicePedana), fnNotNullN(rs!RV_POPesoPedana)
        End Select
    Else
        If fnNotNullN(rs!Link_art_articolo) > 0 Then
            CARICA_MOVIMENTO_DOCUMENTO fnNotNullN(rs!Link_art_articolo), _
            fnNotNull(rs!art_descrizione), fnNotNullN(rs!Link_Art_unita_di_misura), fnNotNullN(rs!Art_quantita_totale), fnNotNullN(rs!Art_prezzo_unitario_neutro), _
            fnNotNullN(rs!Art_Importo_totale_neutro), fnNotNullN(rs!Art_sco_in_percentuale_1), fnNotNullN(rs!Art_sco_in_percentuale_2), fnNotNullN(rs!Art_numero_colli), fnNotNullN(rs!Art_peso), fnNotNullN(rs!Art_tara), fnNotNullN(rs!Art_quantita_pezzi)
        End If
    End If
        
rs.MoveNext
Wend

rs.Close
Set rs = Nothing

Set Mov = Nothing

Me.lblInfoTesta.Font.Bold = False
Me.lblInfoTesta.ForeColor = vbBlack
Screen.MousePointer = 0
End Sub
Private Function CARICA_MOVIMENTO_DOCUMENTO_DERIVATO(IDRiga As Long, IDRigaConferimento, IDAssegnazione As Long, IDProcessoIVGamma As Long, IDSocio As Long, _
DataConferimento As String, NumeroConferimento As Long, _
CodiceLottoEntrata As String, CodiceLottoCampagna As String, CodiceLottoVendita As String, _
QuantitaLiquidazione As Double, ImportoInclusoImballo As Double, ImportoLiquidazione As Double, _
QuantitaMovimentata As Double, IDArticoloMov As Long, ArticoloMov As String, IDUMMov As Long, QuantitaMov As Double, _
PrezzoUnitarioMov As Double, PrezzoImponibileMov As Double, Sconto1Mov As Double, Sconto2Mov As Double, IDTipoDocumentoCoop As Long, _
IDOggettoCollegato As Long, IDTipoOggettoCollegato As Long, IDValoriOggettoDettaglioCollegato As Long, _
QuantitaOrigine As Double, ImportoUnitarioOrigine As Double, PrezzoMedioInLiquidazione As Long, _
IDTipoVariazione As Long, Colli As Long, PesoLordo As Double, Tara As Double, Pezzi As Double, IDAnagraficaFatturazione As Long, IDTipoImportoLiq As Long, _
IDTipoDocumentoCoopDoc As Long, VarImpLiqMan As Double, ImportoRigaCommissioni As Double, _
DataLavorazione As String, IDTipoLavorazione As Long, IDTipoCategoria As Long, IDcalibro As Long, IDTipoLavorazioneConf As Long, PrezzoMedioConf As Long, _
IDPedana As Long, IDTipoPedana As Long, CodicePedana As String, PesoPedana As Double) As Long


Dim Prezzo As Double
Dim PrezzoScontato As Double

Dim Moltiplicatore As Double

'Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(IDArticoloMov)

Mov.DataMovimento = Me.dtData.Text
Mov.FattoreDiConversione = Null
Mov.GestioneMatricole = False
Mov.IDEsercizio = oDoc.IDEsercizio
Mov.IDTipoOggetto = oDoc.IDTipoOggetto
Mov.IDOggetto = oDoc.IDOggetto
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(IDTipoDocumentoCoop, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Me.cboMagazzino.CurrentID
Mov.IDMagazzinoUscita = Me.cboMagazzino.CurrentID
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Me.cdAnagrafica.KeyFieldID
Mov.Field "IDTipoAnagrafica", 2
Mov.Field "IDArticolo", IDArticoloMov
Mov.Field "IDUnitaDiMisura", IDUMMov
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", ArticoloMov
Mov.Field "QuantitaTotale", QuantitaMov
Mov.Field "NumeroDocumento", lngNumero.Value
Mov.Field "DataDocumento", Me.dtData.Text
Mov.Field "Importo", PrezzoImponibileMov
Mov.Field "PrezzoUnitario", PrezzoUnitarioMov
Mov.Field "IDTipoMovimento", 1
Mov.Field "TipoRiga", trcNessuno

Mov.Field "ScontoPerc1", Sconto1Mov
Mov.Field "ScontoPerc2", Sconto2Mov
PrezzoScontato = PrezzoUnitarioMov - ((PrezzoUnitarioMov / 100) * Sconto1Mov)
PrezzoScontato = PrezzoScontato - ((PrezzoScontato / 100) * Sconto2Mov)
Mov.Field "PrezzoScontato", PrezzoScontato
'DATI DI CONFERIMENTO
Mov.Field "RV_POTipoRiga", 1
Mov.Field "IDValoriOggettoDettaglio", IDRiga
Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
Mov.Field "RV_POIDAssegnazioneMerce", IDAssegnazione
Mov.Field "RV_POIDProcessoIVGamma", IDProcessoIVGamma
Mov.Field "RV_POIDAnagraficaSocio", IDSocio
Mov.Field "RV_PODataConferimento", DataConferimento
Mov.Field "RV_PONumeroConferimento", NumeroConferimento
Mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
Mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
Mov.Field "RV_POCodiceLottoVendita", CodiceLottoVendita
Mov.Field "RV_POQuantitaLiquidazione", QuantitaLiquidazione
Mov.Field "RV_POImportoInclusoImballo", ImportoInclusoImballo
Mov.Field "RV_POPrezzoMedioInLiq", PrezzoMedioInLiquidazione

Mov.Field "RV_PODataLavorazione", DataLavorazione
Mov.Field "RV_POIDTipoLavorazione", IDTipoLavorazione
Mov.Field "RV_POIDCalibro", IDcalibro
Mov.Field "RV_POIDTipoCategoria", IDTipoCategoria
Mov.Field "RV_POIDTipoLavorazioneConf", IDTipoLavorazioneConf
Mov.Field "RV_POPrezzoMedioConf", PrezzoMedioConf


Mov.Field "RV_POIDPedana", IDPedana
Mov.Field "RV_POIDTipoPedana", IDTipoPedana
Mov.Field "RV_POCodicePedana", CodicePedana
Mov.Field "RV_POPesoPedana", PesoPedana

'CALCOLO DEL PREZZO DI LIQUIDAZIONE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Prezzo = PrezzoUnitarioMov

'Prezzo = Prezzo / Moltiplicatore

'If ImportoInclusoImballo > 0 Then
'    Prezzo = Prezzo + ImportoInclusoImballo
'Else
'    Prezzo = Prezzo - Abs(ImportoInclusoImballo)
'End If

'If (IDTipoVariazione = 3) Or (IDTipoVariazione = 2) Then
'    Prezzo = GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo, IDOggettoCollegato, IDTipoOggettoCollegato, IDValoriOggettoDettaglioCollegato, CodiceLottoVendita)
'    Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, IDOggettoCollegato, IDTipoOggettoCollegato, Moltiplicatore)
'End If

'Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, oDoc.IDOggetto, oDoc.IDTipoOggetto, Moltiplicatore)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Mov.Field "RV_POImportoLiquidazione", ImportoLiquidazione 'Prezzo
Mov.Field "RV_POQuantitaMovimentata", QuantitaMovimentata 'Quantita di conferimento movimentata
Mov.Field "RV_PONumeroColli", Colli
Mov.Field "RV_POPesoLordo", PesoLordo
Mov.Field "RV_POTara", Tara
Mov.Field "RV_POQuantitaPezzi", Pezzi
Mov.Field "RV_POPesoNetto", PesoLordo - Tara
Mov.Field "RV_POIDOggettoCollegato", IDOggettoCollegato
Mov.Field "RV_POIDTipoOggettoCollegato", IDTipoOggettoCollegato
Mov.Field "RV_POIDValoriOggettoDettaglioCollegato", IDValoriOggettoDettaglioCollegato
Mov.Field "RV_POQuantitaOriginale", QuantitaOrigine
Mov.Field "RV_POImportoOriginale", ImportoUnitarioOrigine
Mov.Field "RV_POTipoRigaCollegata", IDTipoVariazione
Mov.Field "RV_POPrezzoMerceNetta", PrezzoUnitarioMov
Mov.Field "RV_POVariazionePrezzoImballo", 0
Mov.Field "RV_POQuantitaLiqPerPrezzoMedio", QuantitaLiquidazione
Mov.Field "RV_POIDSitoPerAnagrafica", Me.cboAltroSito.CurrentID
Mov.Field "RV_POIDAnagraficaFatturazione", IDAnagraficaFatturazione
Mov.Field "RV_POMerceInclusaImballo", 0
Mov.Field "RV_POIDTipoImportoVenditaLiq", IDTipoImportoLiq
Mov.Field "RV_POVariazionePrezzoManuale", VarImpLiqMan
Mov.Field "RV_POIDTipoDocumentoCoop", IDTipoDocumentoCoopDoc
Mov.Field "RV_POImportoRigaCommissioni", ImportoRigaCommissioni
Mov.Field "RV_POIDLottoCampagnaLavorazione", LINK_LOTTO_PROD_DA_LAV_MOV
Mov.Field "RV_POImportoLiqDoc", ImportoLiquidazione
Select Case IDTipoVariazione
    Case 1 'Parziale valore
        Mov.Field "Oggetto", "Variazione parziale di valore con " & GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto) & " del " & Me.dtData.Text & " numero " & Me.lngNumero.Value
        Mov.Field "RV_POQuantitaLiqPerPrezzoMedio", 0
    Case 2 'Parziale quantità
        Mov.Field "Oggetto", "Variazione parziale di quantità con " & GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto) & " del " & Me.dtData.Text & " numero " & Me.lngNumero.Value
    Case 3 'Totale quantità
        Mov.Field "Oggetto", "Variazione totale di quantità con " & GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto) & " del " & Me.dtData.Text & " numero " & Me.lngNumero.Value
    Case 4 'Totale valore
        Mov.Field "Oggetto", "Variazione totale di valore con " & GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto) & " del " & Me.dtData.Text & " numero " & Me.lngNumero.Value
        Mov.Field "RV_POQuantitaLiqPerPrezzoMedio", 0
End Select
Mov.Field "RV_PODataCompetenzaLiq", DATA_COMPETENZA_LIQ

CARICA_MOVIMENTO_DOCUMENTO_DERIVATO = Mov.Insert

End Function
Private Function GET_COMMISSIONI_DOCUMENTO(PrezzoLiquidazione As Double, IDOggetto As Long, IDTipoOggetto As Long, Moltiplicatore As Double, rstmp As ADODB.Recordset, ImportoMerceRiga As Double, PrezzoLiquidazioneScontato As Double, Optional RigaCommDaDocColl As Double) As Double
Dim sSQL As String
Dim rsComm As DmtOleDbLib.adoResultset
Dim ImportoRigaCommissioni As Double

GET_COMMISSIONI_DOCUMENTO = 0
ImportoRigaCommissioni = 0

sSQL = "SELECT RV_POCommissioniPerDoc.IDRV_POCommissioniPerDoc, RV_POCommissioniPerDoc.IDOggetto, RV_POCommissioniPerDoc.IDRV_POTipoCommissione, RV_POCommissioniPerDoc.Percentuale, "
sSQL = sSQL & "RV_POCommissioniPerDoc.Importo, RV_POCommissioniPerDoc.ImportoRiga, RV_POCommissioniPerDoc.Quantita, RV_POCommissioniPerDoc.APercentuale, RV_POCommissioniPerDoc.ImportoTotale,"
sSQL = sSQL & "RV_POCommissioniPerDoc.IDRV_POTipoPedana , RV_POCommissioniPerDoc.PercentualeDaCommissione, RV_POCommissioniPerDoc.IDArticoloImballo, RV_POTipoCommissione.IDRV_POTipoValoreDocumento "
sSQL = sSQL & "FROM RV_POCommissioniPerDoc INNER JOIN "
sSQL = sSQL & "RV_POTipoCommissione ON RV_POCommissioniPerDoc.IDRV_POTipoCommissione = RV_POTipoCommissione.IDRV_POTipoCommissione "
sSQL = sSQL & "Where IDOggetto = " & IDOggetto
Set rsComm = Cn.OpenResultset(sSQL)

While Not rsComm.EOF
    If GET_CONTROLLO_TIPO_COMM_ND_NC(fnNotNullN(rsComm!IDRV_POTipoCommissione), IDTipoOggetto) = 0 Then
        If (fnNotNullN(rsComm!IDRV_POTipoValoreDocumento) < 5) Then
            ImportoRigaCommissioni = ImportoRigaCommissioni + ((ImportoMerceRiga / 100) * fnNotNullN(rsComm!Percentuale))
            ImportoRigaCommissioni = ImportoRigaCommissioni + (fnNotNullN(rsComm!Importo) / Moltiplicatore)
        
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + ((ImportoMerceRiga / 100) * fnNotNullN(rsComm!Percentuale))
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + (fnNotNullN(rsComm!Importo) / Moltiplicatore)
        Else
            ImportoRigaCommissioni = ImportoRigaCommissioni + ((PrezzoLiquidazioneScontato / 100) * fnNotNullN(rsComm!Percentuale))
            ImportoRigaCommissioni = ImportoRigaCommissioni + (fnNotNullN(rsComm!Importo) / Moltiplicatore)
        
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + ((PrezzoLiquidazioneScontato / 100) * fnNotNullN(rsComm!Percentuale))
            GET_COMMISSIONI_DOCUMENTO = GET_COMMISSIONI_DOCUMENTO + (fnNotNullN(rsComm!Importo) / Moltiplicatore)
        End If
    End If
rsComm.MoveNext
Wend

rsComm.CloseResultset
Set rsComm = Nothing

rstmp!RV_POImportoRigaCommissioni = ImportoRigaCommissioni + RigaCommDaDocColl

GET_COMMISSIONI_DOCUMENTO = PrezzoLiquidazione - GET_COMMISSIONI_DOCUMENTO

End Function

Private Function GET_MOLTIPLICATORE_ARTICOLO(IDArticolo As Long) As Double
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

Private Function CARICA_MOVIMENTO_DOCUMENTO(IDArticoloMov As Long, ArticoloMov As String, IDUMMov As Long, QuantitaMov As Double, _
PrezzoUnitarioMov As Double, PrezzoImponibileMov As Double, Sconto1Mov As Double, Sconto2Mov As Double, Colli As Long, PesoLordo As Double, Tara As Double, Pezzi As Double) As Long


Mov.DataMovimento = Me.dtData.Text
Mov.FattoreDiConversione = Null
Mov.GestioneMatricole = False
Mov.IDEsercizio = oDoc.IDEsercizio
Mov.IDTipoOggetto = oDoc.IDTipoOggetto
Mov.IDOggetto = oDoc.IDOggetto
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(IDDocumento, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Me.cboMagazzino.CurrentID
Mov.IDMagazzinoUscita = Me.cboMagazzino.CurrentID
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Me.cdAnagrafica.KeyFieldID
Mov.Field "IDTipoAnagrafica", 2
Mov.Field "IDArticolo", IDArticoloMov
Mov.Field "IDUnitaDiMisura", IDUMMov
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", ArticoloMov
Mov.Field "QuantitaTotale", QuantitaMov
Mov.Field "DataDocumento", Me.dtData.Text
Mov.Field "NumeroDocumento", Me.lngNumero.Value
Mov.Field "Importo", PrezzoImponibileMov
Mov.Field "PrezzoUnitario", PrezzoUnitarioMov
Mov.Field "ScontoPerc1", Sconto1Mov
Mov.Field "ScontoPerc2", Sconto2Mov
Mov.Field "Oggetto", GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto) & " del " & Me.dtData.Text & " numero " & Me.lngNumero.Value
Mov.Field "IDTipoMovimento", 1
Mov.Field "RV_POIDSitoPerAnagrafica", Me.cboAltroSito.CurrentID
Mov.Field "TipoRiga", trcNessuno

'DATI DI CONFERIMENTO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Mov.Field "RV_POTipoRiga", 1
Mov.Field "IDValoriOggettoDettaglio", 0
Mov.Field "RV_POIDCaricoMerceRighe", 0
Mov.Field "RV_POIDAssegnazioneMerce", 0
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", 0
Mov.Field "RV_PODataConferimento", 0
Mov.Field "RV_PONumeroConferimento", 0
Mov.Field "RV_POCodiceLotto", ""
Mov.Field "RV_POCodiceLottoCampagna", ""
Mov.Field "RV_POCodiceLottoVendita", ""
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POPrezzoMedioInLiq", 0

Mov.Field "RV_PODataLavorazione", Null
Mov.Field "RV_POIDTipoLavorazione", 0
Mov.Field "RV_POIDCalibro", 0
Mov.Field "RV_POIDTipoCategoria", 0
Mov.Field "RV_POIDTipoLavorazioneConf", 0
Mov.Field "RV_POPrezzoMedioConf", 0


Mov.Field "RV_POIDPedana", 0
Mov.Field "RV_POIDTipoPedana", 0
Mov.Field "RV_POCodicePedana", ""
Mov.Field "RV_POPesoPedana", 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POQuantitaMovimentata", 0 'Quantita di conferimento movimentata
Mov.Field "RV_PONumeroColli", 0
Mov.Field "RV_POPesoLordo", 0
Mov.Field "RV_POTara", 0
Mov.Field "RV_POQuantitaPezzi", 0
Mov.Field "RV_POPesoNetto", 0
Mov.Field "RV_POIDOggettoCollegato", 0
Mov.Field "RV_POIDTipoOggettoCollegato", 0
Mov.Field "RV_POIDValoriOggettoDettaglioCollegato", 0
Mov.Field "RV_POQuantitaOriginale", 0
Mov.Field "RV_POImportoOriginale", 0
Mov.Field "RV_POTipoRigaCollegata", 0
Mov.Field "RV_POPrezzoMerceNetta", 0
Mov.Field "RV_POVariazionePrezzoImballo", 0
Mov.Field "RV_POQuantitaLiqPerPrezzoMedio", 0

Mov.Field "RV_POIDAnagraficaFatturazione", 0
Mov.Field "RV_POMerceInclusaImballo", 0
Mov.Field "RV_POIDTipoImportoVenditaLiq", 0
Mov.Field "RV_POVariazionePrezzoManuale", 0
Mov.Field "RV_POIDTipoDocumentoCoop", 0
Mov.Field "RV_POImportoRigaCommissioni", 0
Mov.Field "RV_POQuantitaLiqPerPrezzoMedio", 0
Mov.Field "RV_POIDLottoCampagnaLavorazione", 0
Mov.Field "RV_POImportoLiqDoc", 0
Mov.Field "RV_PODataCompetenzaLiq", Null


CARICA_MOVIMENTO_DOCUMENTO = Mov.Insert

'AggiornamentoProgressivoArticolo oDoc.IDEsercizio, Me.cboMagazzino.CurrentID, IDArticoloMov
        
        
End Function
Private Function GeneraMovimentoConferimento(IDMagazzino As Long, IDRiga As Long, IDRigaConferimento As Long, IDArticoloConferito As Long, DescrizioneArticolo As String, IDUnitaDiMisura As Long, Quantita As Double) As Boolean
Dim Mov As DmtMovim.cMovimentazione
Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection

Mov.DataMovimento = Date
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = oDoc.IDEsercizio
Mov.IDTipoOggetto = oDoc.IDTipoOggetto
Mov.IDOggetto = oDoc.IDOggetto
Mov.IDFunzione = oDoc.IDFunzione
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = IDMagazzino
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
Mov.Field "Oggetto", GET_DESCRIZIONE_TIPOOGGETTO(oDoc.IDTipoOggetto) & " del " & Me.dtData.Text & " numero " & Me.lngNumero.Value
Mov.Field "IDTipoMovimento", 1
Mov.Field "TipoRiga", trcNessuno

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDRiga



GeneraMovimentoConferimento = Mov.Insert

Set Mov = Nothing

AggiornamentoProgressivoArticolo oDoc.IDEsercizio, IDMagazzino, IDArticoloConferito

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

sSQL = "SELECT AttivazioneNuovoMetodoCalcolo, RIferimentiDettagliatiInND "
sSQL = sSQL & " FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    ATTIVAZIONE_NUOVO_CALCOLO = fnNotNullN(rs!AttivazioneNuovoMetodoCalcolo)
    RIPORTA_RIF_DETTAGLIATI = fnNotNullN(rs!RIferimentiDettagliatiInND)
Else
    ATTIVAZIONE_NUOVO_CALCOLO = 0
    RIPORTA_RIF_DETTAGLIATI = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_RIGA_IMPORTO_ORIGINALE(IDOggettoCollegato As Long, IDTipoOggettoCollegato As Long, IDValoriOggettoDettaglioCollegato As Long, CodiceLottoVendita As String) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_RIGA_IMPORTO_ORIGINALE = 0

Select Case IDTipoOggettoCollegato
    Case 114
        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "

    Case 2
        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "

    Case 8
        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
End Select

sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoCollegato
sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIGA_IMPORTO_ORIGINALE = FormatNumber(fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), 3)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_RIGA_A_QUANTITA(IDOggettoCollegato As Long, IDTipoOggettoCollegato As Long, IDValoriOggettoDettaglioCollegato As Long, QuantitaMov As Double, PrezzoUnitario As Double, CodiceLottoVendita As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_RIGA_A_QUANTITA = 0

Select Case IDTipoOggettoCollegato
    Case 114
        sSQL = "SELECT ValoriOggettoDettaglio0001.*, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0001 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0001.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0001.Link_Art_IVA "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0001.IDOggetto=" & IDOggettoCollegato
    Case 2
        sSQL = "SELECT ValoriOggettoDettaglio0004.*, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0004 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0004.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0004.Link_Art_IVA "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0004.IDOggetto=" & IDOggettoCollegato

    Case 8
        sSQL = "SELECT ValoriOggettoDettaglio0034.*, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, Iva.Iva AS IvaVendita, "
        sSQL = sSQL & "Iva.Codice AS CodiceIvaVendita, Iva_1.IDIva AS IDIvaArticolo, Iva_1.Iva AS IvaArticolo, Iva_1.AliquotaIva AS AliquotaIvaArticolo, "
        sSQL = sSQL & "Iva_1.Codice AS CodiceIvaArticolo, Articolo.IDCategoriaMerceologica "
        sSQL = sSQL & "FROM Iva RIGHT OUTER JOIN "
        sSQL = sSQL & "Iva AS Iva_1 RIGHT OUTER JOIN "
        sSQL = sSQL & "ValoriOggettoDettaglio0034 INNER JOIN "
        sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto LEFT OUTER JOIN "
        sSQL = sSQL & "Articolo ON ValoriOggettoDettaglio0034.Link_Art_articolo = Articolo.IDArticolo ON Iva_1.IDIva = Articolo.IDIvaVendita ON "
        sSQL = sSQL & "Iva.IDIva = ValoriOggettoDettaglio0034.Link_Art_IVA "
        sSQL = sSQL & "WHERE ValoriOggettoDettaglio0034.IDOggetto=" & IDOggettoCollegato

End Select

sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)


Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If QuantitaMov - fnNotNullN(rs!Art_quantita_totale) = 0 Then
        If FormatNumber(fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), 3) = FormatNumber(PrezzoUnitario, 3) Then
            GET_RIGA_A_QUANTITA = 3
        Else
            GET_RIGA_A_QUANTITA = 4
        End If
    Else
        If FormatNumber(fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA), 3) = FormatNumber(PrezzoUnitario, 3) Then
            GET_RIGA_A_QUANTITA = 2
        Else
            GET_RIGA_A_QUANTITA = 1
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
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
Private Sub AggiornamentoProgressivoArticolo(IDEsercizio As Long, IDMagazzino As Long, IDArticolo As Long)
Dim prg As dmtProg.cRicostruzione


Set prg = New dmtProg.cRicostruzione

Set prg.Connessione = TheApp.Database.Connection
    prg.Filtri.Add m_App.IDFirm, "IDAzienda"

    prg.Filtri.Add IDEsercizio, "IDEsercizio"

    prg.Filtri.Add IDMagazzino, "IDMagazzino"

    prg.Where = "Articolo.IDArticolo = " & fnNormString(IDArticolo)
    'Vengono ricostruiti in questo caso le giacenze dei lotti degli articoli che hanno codice compreso tra 'A000' e 'A999', di esercizio, azienda e magazzino specificati
    
    prg.RicostruzioneProgressivi
    
End Sub
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
    If TypeOf Controls(I) Is DmtSearchACS2 Then
        Controls(I).Enabled = Abilitato
    End If
    
Next

'''''''''CONTROLLI STANDARD SEMPRE ABILITATI O NON ABILITATI
'TESTA DEL DOCUMENTO
Me.txtNLetteraIntento.Enabled = False
Me.txtDataLetteraIntento.Enabled = False
If oDoc.IDOggetto > 0 Then
    Me.cboSezionale.Enabled = False
Else
    Me.cboSezionale.Enabled = True
End If

'PIEDE DEL DOCUMENTO
Me.curTotArrotondamenti.Enabled = False
Me.curTotImposta.Enabled = False
Me.curTotImponibile.Enabled = False
Me.curTotDocumento.Enabled = False
Me.curNettoAPagare.Enabled = False
Me.curNettoAPagare_naz.Enabled = False


Me.chkCessione.Enabled = True
Me.txtDataCambio.Enabled = False
Me.txtValoreCambioValuta.Enabled = False

'CORPO DOCUMENTO
Me.txtImponibileUnitario.Enabled = False
Me.txtTotaleRiga.Enabled = False
Me.cmdSalva.Enabled = True
Me.chkPrezzoMedioInLiq.Enabled = True
Me.txtIDTipoVariazione.Enabled = False
Me.txtPrezzoOriginale.Enabled = False
Me.txtQuantitaOriginale.Enabled = False
Me.txtColliOriginali.Enabled = False
Me.txtPesoLordoOriginale.Enabled = False
Me.txtTaraOriginale.Enabled = False
Me.txtPezziOriginali.Enabled = False
Me.txtPesoNettoOriginale.Enabled = False
Me.txtConferimentoRighe.Enabled = False
Me.txtDataConferimento.Enabled = False
Me.txtCodiceSocio.Enabled = False
Me.txtSocio.Enabled = False
Me.CDSocioFatt.Enabled = False
Me.cmdAgenteRiga.Enabled = True
Me.cboTipoDocumentoCoop.Enabled = False
Me.cmdAgenteRiga.Enabled = True
Me.txtDataLavorazione.Enabled = False
Me.chkRiscontroPeso.Enabled = False
Me.txtSconto1Ori.Enabled = False
Me.txtSconto2Ori.Enabled = False
Me.cboUnitaDiMisura.Enabled = False
'COMMISSIONI
Me.cboCommissioni.Enabled = True
Me.txtPercCommissioni.Enabled = True
Me.txtImportoRigaComm.Enabled = True
Me.txtImportoCommissioni.Enabled = True
Me.cmdNuovaCommissione.Enabled = True
Me.cmdSalvaCommissione.Enabled = True
Me.cmdEliminaCommissione.Enabled = True

chkCessione_Click
End Sub
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
Private Function GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO(PrezzoOriginale As Double, Quantita_Originale As Double, Prezzo_Riga_Doc As Double, Quantita_Riga_Doc As Double) As Long
    
If Quantita_Riga_Doc - Quantita_Originale = 0 Then
    If FormatNumber(Prezzo_Riga_Doc, 3) = FormatNumber(PrezzoOriginale, 3) Then
        GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO = 3
    Else
        GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO = 4
    End If
Else
    If FormatNumber(Prezzo_Riga_Doc, 3) = FormatNumber(PrezzoOriginale, 3) Then
        GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO = 2
    Else
        GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO = 1
    End If
                    
End If

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
    rs.CloseResultset
    Set rs = Nothing
    
    sSQL = "SELECT IDListinoDiBase "
    sSQL = sSQL & "FROM ConfigurazioneVendite "
    sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    
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
        
        Set rs = Cn.OpenResultset(sSQL)
        
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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = 0
Else
    GET_LISTINO_DEFAULT_PARAMETRI_AZIENDA = fnNotNullN(rs!IDListinoImballiDefault)
    
End If
rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_RIGHE_DOCUMENTO(NomeTabella As String)
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
sSQL = sSQL & "RV_POImportoDaLiq, RV_POLinkRiga, RV_POImportoImballoInArticolo, Art_numero_colli, Art_quantita_Totale, Link_art_articolo, "
sSQL = sSQL & "RV_POIDOggetto, RV_POIDTipoOggetto, RV_POIDValoriOggettoDettaglio, RV_POCodiceLotto, "
sSQL = sSQL & "RV_POVariazionePrezzoManuale, RV_POImportoRigaCommissioni,  "
sSQL = sSQL & "RV_POIDConferimentoRighe, RV_POIDAssegnazioneMerce, RV_POIDProcessoIVGamma, RV_PODataLavorazione, "
sSQL = sSQL & "RV_POIDPedana, RV_POCodicePedana, RV_POIDTipoPedana, RV_POPesoPedana, RV_PORigaRiscontroPeso, RV_POQuantitaLiq ,  "
sSQL = sSQL & "Art_numero_colli,Art_Peso, Art_Tara, Art_quantita_pezzi, RV_POImportoLiqDoc, Art_prezzo_unitario_neutro, "
sSQL = sSQL & "Art_sco_in_percentuale_1,Art_sco_in_percentuale_2 "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND RV_POIDValoriOggettoDettaglio>0 "

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF

    Moltiplicatore = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!Link_art_articolo))
    Prezzo = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_1))
    Prezzo = Prezzo - ((Prezzo / 100) * fnNotNullN(rs!Art_sco_in_percentuale_2))
    
    IDUMCoop_ArtVenduto = GET_LINK_UM_LIQUIDAZIONE_ARTICOLO(fnNotNullN(rs!Link_art_articolo))
    
    If IDUMCoop_ArtVenduto > 0 Then
        rs!RV_POQuantitaLiq = fnNotNullN(rs!Art_quantita_totale) * Moltiplicatore
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
        
        If fnNotNullN(rs!RV_POQuantitaLiq) > 0 Then
            Prezzo = (fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA) * fnNotNullN(rs!Art_quantita_totale)) / fnNotNullN(rs!RV_POQuantitaLiq)
        Else
            Prezzo = Prezzo / Moltiplicatore
        End If
    Else
        Prezzo = Prezzo / Moltiplicatore
    End If

    
    PrezzoScontato = Prezzo

    If (fnNotNullN(rs!RV_PORigaRiscontroPeso) = 1) Then
        Prezzo = Abs(Prezzo)
    End If
    

    If fnNotNullN(rs!RV_POImportoImballoInArticolo) = 0 Then
        rs!RV_POVariazionePrezzoImballo = 0
        rs!RV_POImportoMerceNetta = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
    Else
        rs!RV_POVariazionePrezzoImballo = ((Abs(fnNotNullN(rs!RV_POImportoDaLiq)) * fnNotNullN(rs!RV_POQuantitaLiq)) / fnNotNullN(rs!Art_quantita_totale))
        rs!RV_POImportoMerceNetta = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA) - rs!RV_POVariazionePrezzoImballo
    End If

    MerceNettaPerLiquidazione = fnNotNullN(rs!RV_POImportoMerceNetta) '/ Moltiplicatore
    
    rs!RV_POImportoRigaCommissioni = 0
    rs!RV_POImportoDaLiq = 0
    
    If (fnNotNullN(rs!RV_POIDTipoVariazione) = 2) Or (fnNotNullN(rs!RV_POIDTipoVariazione) = 3) Then
        If (NON_CALC_INCIDENZA_IMB = 0) Then Prezzo = GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), fnNotNullN(rs!RV_POIDValoriOggettoDettaglio), fnNotNull(rs!RV_POCodiceLotto), rs)
    End If
    
    If (NON_CALC_COMM = 0) Then Prezzo = GET_COMMISSIONI_DOCUMENTO(Prezzo, fnNotNullN(rs!RV_POIDOggetto), fnNotNullN(rs!RV_POIDTipoOggetto), Moltiplicatore, rs, Prezzo, PrezzoScontato, 0)
   
    
    If fnNotNullN(rs!RV_POVariazionePrezzoManuale) >= 0 Then
        Prezzo = Prezzo + fnNotNullN(rs!RV_POVariazionePrezzoManuale)
    Else
        Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POVariazionePrezzoManuale))
    End If
    
    rs!RV_POImportoLiq = Prezzo
    rs!RV_POImportoLiqDoc = Prezzo
    
    If fnNotNullN(rs!RV_POIDAssegnazioneMerce) > 0 Then
        'rs!RV_PODataLavorazione = GET_DATI_LAVORAZIONE(fnNotNullN(rs!RV_POIDAssegnazioneMerce), "DataDocumento")
    End If
    
    
    rs.Update
    
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub

Private Function GET_CESSIONE_INTRA_CLIENTE(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'sSQL = "SELECT ContoPDC.IDClassificazioneIva, Cliente.IDAnagrafica, ContoPDC.ContoPDC, Cliente.IDAzienda "
'sSQL = sSQL & "FROM Cliente INNER JOIN "
'sSQL = sSQL & "ContoPDC ON Cliente.IDAnagrafica = ContoPDC.IDAnagrafica "
'sSQL = sSQL & "WHERE Cliente.IDAzienda =" & TheApp.IDFirm
'sSQL = sSQL & " AND Cliente.IDAnagrafica=" & IDAnagraficaCliente

sSQL = "SELECT IDClassificazioneIva FROM Cliente "
sSQL = sSQL & "WHERE Cliente.IDAzienda =" & TheApp.IDFirm
sSQL = sSQL & " AND Cliente.IDAnagrafica=" & IDAnagraficaCliente

Set rs = Cn.OpenResultset(sSQL)


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


End Function
Private Function GET_CAMPI_INSTRASTAT_AZIENDA(IDAzienda) As Long
On Error GoTo ERR_GET_CAMPI_INSTRASTAT_AZIENDA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDComune, IDNazione FROM RepAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
Set rs = Cn.OpenResultset(sSQL)



If rs.EOF Then
    Me.cboIntraProvincia.WriteOn oDoc.Field("Link_Doc_intra_provinc_merce", 0, sTabellaTestata)
    Me.cboIntraNazione.WriteOn oDoc.Field("Link_Doc_intra_naz_pagamento", 0, sTabellaTestata)
Else
    Me.cboIntraProvincia.WriteOn oDoc.Field("Link_Doc_intra_provinc_merce", GET_LINK_PROVINCIA_INTRA(TheApp.Branch), sTabellaTestata)
    
    If Me.cboIntraProvincia.CurrentID = 0 Then
        Me.cboIntraProvincia.WriteOn oDoc.Field("Link_Doc_intra_provinc_merce", GET_LINK_PROVINCIA(fnNotNullN(rs!IDComune)), sTabellaTestata)
    End If
    Me.cboIntraNazione.WriteOn oDoc.Field("Link_Doc_intra_naz_pagamento", fnNotNullN(rs!IDNazione), sTabellaTestata)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_CAMPI_INSTRASTAT_AZIENDA:
    MsgBox Err.Description, vbCritical, "GET_CAMPI_INSTRASTAT_AZIENDA"
End Function
Private Function GET_LINK_PROVINCIA(IDComune As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDProvincia FROM Comune "
sSQL = sSQL & "WHERE IDComune=" & IDComune

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    oDoc.Field "Art_intra_non_riporto", 0, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_nomenclatura", 0, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_natura_trans", 0, sTabellaDettaglio
    oDoc.Field "Art_intra_qta_tot_massa_netta", 0, sTabellaDettaglio

Else
    If Abs(fnNotNullN(rs!NonRiportoIntrastat)) = 1 Then
        oDoc.Field "Art_intra_non_riporto", fnNotNullN(rs!NonRiportoIntrastat), sTabellaDettaglio
        oDoc.Field "Link_Art_intra_nomenclatura", 0, sTabellaDettaglio
        oDoc.Field "Link_Art_intra_natura_trans", 0, sTabellaDettaglio
        oDoc.Field "Art_intra_qta_tot_massa_netta", 0, sTabellaDettaglio
    Else
        oDoc.Field "Art_intra_non_riporto", fnNotNullN(rs!NonRiportoIntrastat), sTabellaDettaglio
        oDoc.Field "Link_Art_intra_nomenclatura", fnNotNullN(rs!IDNomenclaturaCombinata), sTabellaDettaglio
        oDoc.Field "Link_Art_intra_natura_trans", 2 'GET_LINK_NATURA_TRANSAZIONE(IDArticolo, IDAzienda, IDFiliale), sTabellaDettaglio
        oDoc.Field "Art_intra_qta_tot_massa_netta", GET_QUANTITA_MASSA_NETTA_INSTRASTAT(IDArticolo, IDAzienda, Quantita), sTabellaDettaglio
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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_CODICE_AGE = ""
Else
    GET_LINK_CODICE_AGE = fnNotNull(rs!Codice)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_DATI_RICHIESTI_INTRA(IDAzienda As Long, DataDocumento As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Periodo_IVA As Long
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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_DATI_RICHIESTI_INTRA = 1
Else
    Select Case fnNotNullN(rs!IDPeriodicitaCessione)
        Case 1
            GET_LINK_DATI_RICHIESTI_INTRA = 1
        Case 2
            GET_LINK_DATI_RICHIESTI_INTRA = 1
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
Private Sub SCRIVI_AGENTE_DA_RIGA_DOC(IDTipoOggettoCollegato As Long, IDOggettoCollegato As Long, IDValoriOggettoCollegato As Long)
Dim sSQL As String
Dim rsAge As DmtOleDbLib.adoResultset

Select Case IDTipoOggettoCollegato

    Case 114
        sSQL = "SELECT Art_age_codice, Art_age_importo_provv, Art_age_nome, Art_age_percentuale_provv, "
        sSQL = sSQL & "Art_age_ragione_sociale, Art_age_regola_provv, Link_Art_age_regola_provv, Link_Art_agente, Link_Art_age_tipo_ordine "
        sSQL = sSQL & " FROM ValoriOggettoDettaglio0001 "
        sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoCollegato
        sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoCollegato
    Case 2
        sSQL = "SELECT Art_age_codice, Art_age_importo_provv, Art_age_nome, Art_age_percentuale_provv, "
        sSQL = sSQL & "Art_age_ragione_sociale, Art_age_regola_provv, Link_Art_age_regola_provv, Link_Art_agente, Link_Art_age_tipo_ordine "
        sSQL = sSQL & " FROM ValoriOggettoDettaglio0004 "
        sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoCollegato
        sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoCollegato
    Case Else
        Exit Sub
End Select

Set rsAge = Cn.OpenResultset(sSQL)

If Not rsAge.EOF Then
 
    '''AGENTE
    oDoc.Field "Link_Art_agente", fnNotNullN(rsAge!Link_Art_agente), sTabellaDettaglio
    oDoc.Field "Art_age_nome", fnNotNull(rsAge!Art_age_nome), sTabellaDettaglio
    oDoc.Field "Art_age_codice", fnNotNull(rsAge!Art_age_codice), sTabellaDettaglio
    oDoc.Field "Art_age_ragione_sociale", fnNotNull(rsAge!Art_age_ragione_sociale), sTabellaDettaglio
    
    oDoc.Field "Link_Art_age_regola_provv", fnNotNullN(rsAge!Link_Art_age_regola_provv), sTabellaDettaglio
    oDoc.Field "Art_age_regola_provv", fnNotNull(rsAge!Art_age_regola_provv), sTabellaDettaglio
    oDoc.Field "Link_Art_age_tipo_ordine", fnNotNull(rsAge!Link_Art_age_tipo_ordine), sTabellaDettaglio

    If fnNotNullN(rsAge!Art_age_percentuale_provv) > 0 Then
        oDoc.Field "Art_age_percentuale_provv", fnNotNullN(rsAge!Art_age_percentuale_provv), sTabellaDettaglio
    End If
    If fnNotNullN(rsAge!Art_age_importo_provv) > 0 Then
        oDoc.Field "Art_age_importo_provv", fnNotNullN(rsAge!Art_age_importo_provv), sTabellaDettaglio
    End If

End If
rsAge.CloseResultset
Set rsAge = Nothing


End Sub
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

Private Function GET_PREFISSO_SEZ(IDSezionale As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Prefisso FROM Sezionale "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDSezionale=" & IDSezionale

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREFISSO_SEZ = ""
Else
    GET_PREFISSO_SEZ = fnNotNull(rs!Prefisso)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub CREA_PROVV_AGENTI(NuovoDocumento As Boolean)
Dim oAgents As DmtAgentsLib.Agents
Dim oAgent As DmtAgentsLib.Agent

Set oAgManager = New DmtAgentsLib.AgentManager
If oAgManager.Database Is Nothing Then
    Set oAgManager.Database = TheApp.Database
End If

             
oAgManager.IDBranch = oDoc.IDFiliale        'Identificativo della filiale
oAgManager.IDFirm = oDoc.IDAzienda         'Identificativo dell'azienda
oAgManager.IDUser = oDoc.IDUtente           'Identificativo dell'utente corrente
oAgManager.Password = TheApp.Password                           'Password dell'utente corrente
oAgManager.UserName = TheApp.User                 'Nome utente dell'utente corrente
             
Set oAgent = oAgManager.CreateDocument
oAgent.OpenBySelectedID 1  '  Da lasciare fisso così
      
Set oAgents = New DmtAgentsLib.Agents
             
'Verifica se per l'azienda corrente è prevista la generazione
' automatica dei movimenti di provvigione e in tal caso avvia il calcolo
If oAgents.CheckCreateMov(oDoc.IDAzienda) Then
    If NuovoDocumento = False Then
    'Se si è in variazione
    'Se il documento è in variazione è necessario avviare
    'l'annullamento dell'eventuali provvigioni precedentemente calcolate
        If Not oAgents.CtrlExistCommLiquidated(oDoc.IDOggetto, oDoc.IDTipoOggetto) Then
            oAgents.AbortCalcCommissions oDoc.IDOggetto, oDoc.IDTipoOggetto
            'Riesegue il calcolo dei movimenti
                If oDoc.IDTipoOggetto <> 11 Then   'Se non si tratta di una Nota di credito
                'Calcolo e creazione dei movimenti di provvigione
                    oAgents.CreateMovements oDoc.IDTipoOggetto, oDoc.IDOggetto, oDoc.Field("Doc_data")
                Else         'Se si tratta di una Nota di credito
                'Calcolo e creazione dei movimenti di provvigione
                    oAgents.CreateAdjustmentMov oDoc.IDTipoOggetto, oDoc.IDOggetto, oDoc.Field("Doc_data")
                End If
        End If
    Else                   'Se si è in inserimento
        If oDoc.IDTipoOggetto <> 11 Then '  Se non si tratta di una Nota di credito
        'Calcolo e creazione dei movimenti di provvigione
            oAgents.CreateMovements oDoc.IDTipoOggetto, oDoc.IDOggetto, oDoc.Field("Doc_data")
        Else
            '  Calcolo e creazione dei movimenti di provvigione
            oAgents.CreateAdjustmentMov oDoc.IDTipoOggetto, oDoc.IDOggetto, oDoc.Field("Doc_data")
        End If
    End If
End If
oAgent.Destroy
Set oAgent = Nothing
Set oAgents = Nothing
'Set oAgManager = Nothing

If oAgManager.Database Is Nothing Then
    Set oAgManager.Database = TheApp.Database
End If
End Sub

Private Function GET_VARIAZIONE_LIQ_DOC_ORI(Prezzo As Double, IDOggetto As Long, IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long, CodiceLottoVendita As String, rstmp As ADODB.Recordset) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoIncidenzaImballo As Double

sSQL = "SELECT IDMovimento, RV_POImportoInclusoImballo "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
If (COLLEGAMENTO_NOTA_PER_LOTTO = 1) Then
    sSQL = sSQL & " AND RV_POCodiceLottoVendita=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
End If
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    ImportoIncidenzaImballo = fnNotNullN(rs!RV_POImportoInclusoImballo)
    If fnNotNullN(rs!RV_POImportoInclusoImballo) > 0 Then
        Prezzo = Prezzo + fnNotNullN(rs!RV_POImportoInclusoImballo)
    Else
        Prezzo = Prezzo - Abs(fnNotNullN(rs!RV_POImportoInclusoImballo))
    End If
End If

rs.CloseResultset
Set rs = Nothing

rstmp!RV_POImportoDaLiq = ImportoIncidenzaImballo

GET_VARIAZIONE_LIQ_DOC_ORI = Prezzo
End Function
Private Sub ParametroNumeroDecimali()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoDecimaliPesiVendita FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Numero_Decimali_Pesi = fnNotNullN(rs!IDRV_POTipoDecimaliPesiVendita) - 1
    If Numero_Decimali_Pesi < 0 Then
        Numero_Decimali_Pesi = 2
    End If
Else
    Numero_Decimali_Pesi = 2
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_PREZZO_MEDIO_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT NonCalcolarePrezzoMedio "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

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
Private Function GET_FORZATURA_PREZZO_LIQ_CLIENTE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoImportoVenditaLiq "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente"
sSQL = sSQL & " WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = 0
Else
    GET_FORZATURA_PREZZO_LIQ_CLIENTE = fnNotNullN(rs!IDRV_POTipoImportoVenditaLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function

Private Function GET_VALORE_FORZA_TIPO_PREZZO_LIQ(IDTipoOggetto As Long, IDOggetto As Long, CodiceLottoVandita As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POIDTipoImportoVenditaLiq "
Select Case IDTipoOggetto
    
    Case 114
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
    Case 2
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
End Select

sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVandita)
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_VALORE_FORZA_TIPO_PREZZO_LIQ = 0
Else
    GET_VALORE_FORZA_TIPO_PREZZO_LIQ = Abs(fnNotNullN(rs!RV_POIDTipoImportoVenditaLiq))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_ALTRO_SITO(IDAnagrafica As Long) As Long
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT IDSitoPerAnagrafica FROM Cliente "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ALTRO_SITO = 0
Else
    GET_LINK_ALTRO_SITO = fnNotNullN(rs!IDSitoPerAnagrafica)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub SCRIVI_PDC_DA_RIGA_DOC(IDTipoOggettoCollegato As Long, IDOggettoCollegato As Long, IDValoriOggettoCollegato As Long)
Dim sSQL As String
Dim rsAge As DmtOleDbLib.adoResultset

Select Case IDTipoOggettoCollegato

    Case 114
        sSQL = "SELECT Link_Art_IDCContropartita, Art_CContropartita_codifica, Art_CContropartita_descrizione, RV_POIDTipoDocumentoCoop "
        sSQL = sSQL & " FROM ValoriOggettoDettaglio0001 "
        sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoCollegato
        sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoCollegato
    Case 2
        sSQL = "SELECT Link_Art_IDCContropartita, Art_CContropartita_codifica, Art_CContropartita_descrizione, RV_POIDTipoDocumentoCoop "
        sSQL = sSQL & " FROM ValoriOggettoDettaglio0004 "
        sSQL = sSQL & " WHERE IDOggetto=" & IDOggettoCollegato
        sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoCollegato
    Case Else
        Exit Sub
End Select

Set rsAge = Cn.OpenResultset(sSQL)

If Not rsAge.EOF Then
    oDoc.Field "Link_Art_IDCContropartita", fnNotNullN(rsAge!Link_Art_IDCContropartita), sTabellaDettaglio
    oDoc.Field "Art_CContropartita_codifica", fnNotNull(rsAge!Art_CContropartita_codifica), sTabellaDettaglio
    oDoc.Field "Art_CContropartita_descrizione", fnNotNull(rsAge!Art_CContropartita_descrizione), sTabellaDettaglio
    oDoc.Field "RV_POIDTipoDocumentoCoop", fnNotNull(rsAge!RV_POIDTipoDocumentoCoop), sTabellaDettaglio
End If
rsAge.CloseResultset
Set rsAge = Nothing


End Sub
Private Sub SetPDCProperties()
Dim oNode As DmtPDC.INode
Dim oBranch As DmtPDC.Branch

    Set oPDC = New DmtPDC.PDCServices
    'Imposta le proprietà dell'oggetto PDCServices
    With oPDC
        'Viene fornita al controllo la connessione al database DMT.
        'La connessione è di tipo ADO.Connection quindi viene
        'passata la proprietà InternalConnection dell'oggetto Database
        Set .Connection = Cn.InternalConnection
        
        'Indica l'identificativo del Piano dei conti da visualizzare
        .IDPDC = Link_PianoDeiConti
        .HideAccounts = False
        .BranchType = btcAllBranchs
        '.BranchType = .BranchType + btcRevenuesBranch
        .AccountType = atcAllAccounts
        
        If Len(Me.txtCodiceConto.Text) > 0 Then
            Set oNode = .SearchNodeExtended(Me.txtCodiceConto.Text)
        ElseIf Len(Me.txtDescrizioneConto.Text) > 0 Then
            Set oNode = .SearchNodeExtended(, Me.txtDescrizioneConto.Text)
        Else
            Set oNode = .SearchNodeExtended
        End If
        
        If .RecordFounded = 1 Then
                If TypeName(oNode) = "Account" Then
                
                    Me.txtIDPianoDeiConti.Value = oNode.ID
                    'Codifica completa del Conto o del Ramo
                    Me.txtCodiceConto = oNode.CompletedCode
                    Me.txtDescrizioneConto.Text = oNode.Description
                
                Else
                
                    If Len(Me.txtCodiceConto.Text) > 0 Then
                        .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
                    Else
                        .SelectedNode.Description = Me.txtDescrizioneConto.Text
                    End If
                        
                    .ShowSearchDialog
                    
                    ShowNodeProperties oPDC.SelectedNode
                End If
            
        ElseIf .RecordFounded > 1 Then
            If Len(Me.txtCodiceConto.Text) > 0 Then
                .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
            Else
                .SelectedNode.Description = Me.txtDescrizioneConto.Text
            End If

            .ShowSearchDialog
            
            ShowNodeProperties oPDC.SelectedNode
            
        Else
            If Len(Me.txtCodiceConto.Text) > 0 Then
                .SelectedNode.CompletedCode = Me.txtCodiceConto.Text
            Else
                .SelectedNode.Description = Me.txtDescrizioneConto.Text
            End If
        
        
            .ShowSearchDialog
            
            ShowNodeProperties oPDC.SelectedNode
            
        End If
        
    End With
    
    Set oPDC = Nothing
End Sub
Private Sub ShowNodeProperties(ByVal oNode As DmtPDC.INode)
Dim oAccount As DmtPDC.Account
Dim oBranch As DmtPDC.Branch

If Not oNode Is Nothing Then
    Me.txtIDPianoDeiConti.Value = oNode.ID
    Me.txtCodiceConto = oNode.CompletedCode
    Me.txtDescrizioneConto.Text = oNode.Description
End If
End Sub

Private Function GetPianoDeiConti() As Long
'On Error GoTo ERR_GetPianoDeiConti
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT IDPianoDeiConti FROM PianoDeiConti WHERE ("
    sSQL = sSQL & "(IDAzienda = " & TheApp.IDFirm & ") AND "
    sSQL = sSQL & "(TipoPDC = " & 1 & ") AND "
    sSQL = sSQL & "(IDEsercizio= " & oDoc.IDEsercizio & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        GetPianoDeiConti = fnNotNullN(rs!IDPianoDeiConti)
    Else
        GetPianoDeiConti = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Exit Function
ERR_GetPianoDeiConti:
    MsgBox Err.Description, vbCritical, "Errore piano dei conti"
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
Private Function PREPARAZIONE_EMAIL(NomeFile As String, IDAnagrafica As Long, IDDestinazione As Long, NumeroDocumento As String, DataDocumento As String, DescrizioneOggetto As String) As Boolean
On Error GoTo ERR_PREPARAZIONE_EMAIL
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

PREPARAZIONE_EMAIL = False
'On Error Resume Next
'''''''VARIABILI PER OUTLOOK''''''''''''''''''''
Dim Out 'As Outlook.Application
Dim Name 'As Outlook.NameSpace
Dim f 'As Outlook.MAPIFolder 'MAIN MAPIFOLDER
Dim G 'As Outlook.MAPIFolder 'CONTATCT MAPIFOLDER
Dim H 'As Outlook.MAPIFolder 'CARTELLA POINT MAPIFOLDER
Dim Rubrica ' As Outlook.AddressList
Dim Cont 'As Outlook.ContactItem
Dim RCP 'As Outlook.Recipient
Dim Lista 'As Outlook.DistListItem
Dim olMail 'As Outlook.MailItem
Dim oAttach 'As Outlook.Attachment
''''''VARIABILI GENERALI''''''''''''''''''''''''''''
Dim percorso As String
Dim StringaBody As String


    Set Out = CreateObject("Outlook.Application")
    
    Set Name = Out.GetNamespace("MAPI")
    Name.Logon
    
    'Set olMail = Out.CreateItem(olMailItem)
    Set olMail = Out.CreateItem(0)
    olMail.To = GET_EMAIL_ANAGRAFICA(IDAnagrafica)
    If IDDestinazione > 0 Then
        olMail.cc = GET_EMAIL_ANAGRAFICA_DEST(IDDestinazione)
    End If
    olMail.Subject = DescrizioneOggetto & " n° " & NumeroDocumento & " del " & DataDocumento
    
    ''''RECUPERO DAI DEL CORPO DEL MESSAGGIO
    StringaBody = DescrizioneOggetto & " n° " & NumeroDocumento & " del " & DataDocumento
    olMail.Body = StringaBody
    
    olMail.Attachments.Add NomeFile
    
    olMail.Display
    
    
    Name.Logoff
    Set Name = Nothing
    Set olMail = Nothing
    Set Out = Nothing
    PREPARAZIONE_EMAIL = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Function
ERR_PREPARAZIONE_EMAIL:
    MsgBox Err.Description, vbCritical, "Preparazione e-mail"
End Function

Public Function TrovaCartella(IDLCartella As Long) As String

    TrovaCartella = String$(MAX_PATH, 0)
    
    Call SHGetSpecialFolderPath(ByVal 0&, TrovaCartella, IDLCartella, ByVal 0&)
    
    TrovaCartella = Left$(TrovaCartella, InStr(1, TrovaCartella, Chr$(0)) - 1)
    
    If Len(TrovaCartella) > 0 And Right$(TrovaCartella, 1) <> "\" Then TrovaCartella = TrovaCartella & "\"
End Function
Private Function GET_EMAIL_ANAGRAFICA(IDAnagrafica As Long) As String
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT EmailInternet FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_EMAIL_ANAGRAFICA = ""
Else
    GET_EMAIL_ANAGRAFICA = fnNotNull(rs!EmailInternet)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_EMAIL_ANAGRAFICA_DEST(IDSitoPerAnagrafica As Long) As String
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Email FROM SitoPerAnagrafica "
sSQL = sSQL & "WHERE IDSitoPerAnagrafica=" & IDSitoPerAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_EMAIL_ANAGRAFICA_DEST = ""
Else
    GET_EMAIL_ANAGRAFICA_DEST = fnNotNull(rs!Email)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub ParametroPrezzoMedioAutomatico()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT PrezzoMedioDaConf FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    PREZZO_MEDIO_AUT = fnNotNullN(rs!PrezzoMedioDaConf)
Else
    PREZZO_MEDIO_AUT = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub ParametroAggiornaPrezzoMedioDaConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AggiornaPrezzoMedioDaConf FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    AGGIORNA_PREZZO_MEDIO = fnNotNullN(rs!AggiornaPrezzoMedioDaConf)
Else
    AGGIORNA_PREZZO_MEDIO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroAggiornaTipoLavorazioneDaConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AggiornaTipoLavDaConf FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    AGGIORNA_TIPO_LAVORAZIONE = fnNotNullN(rs!AggiornaTipoLavDaConf)
Else
    AGGIORNA_TIPO_LAVORAZIONE = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_DATI_RIGA_CONFERIMENTO(IDRigaConferimento As Long, NomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DATI_RIGA_CONFERIMENTO = 0
Else
    GET_DATI_RIGA_CONFERIMENTO = fnNotNullN(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DATI_LAVORAZIONE(IDLavorazione As Long, NomeCampo As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POAssegnazioneMerce=" & IDLavorazione

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DATI_LAVORAZIONE = ""
Else
    GET_DATI_LAVORAZIONE = fnNotNull(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_LINK_TIPO_PEDANA(IDPedana As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPedana FROM RV_POPedana "
sSQL = sSQL & "WHERE IDRV_POPedana=" & IDPedana

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PESO_PEDANA = 0
Else
    GET_PESO_PEDANA = fnNotNullN(rs!PesoPedana)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_VALORE_CAMPO_RIGA_VENDITA(IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long, NomeCampo As String, StringaReturn As Boolean)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT  " & NomeCampo
Select Case IDTipoOggetto
    
    Case 114
        sSQL = sSQL & " FROM ValoriOggettoDettaglio0001 "
    Case 2
        sSQL = sSQL & " FROM ValoriOggettoDettaglio0004 "
End Select

sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    If StringaReturn = True Then
        GET_VALORE_CAMPO_RIGA_VENDITA = ""
    Else
        GET_VALORE_CAMPO_RIGA_VENDITA = 0
    End If
Else
    If StringaReturn = True Then
        GET_VALORE_CAMPO_RIGA_VENDITA = fnNotNull(rs.adoColumns(NomeCampo).Value)
    Else
        GET_VALORE_CAMPO_RIGA_VENDITA = fnNotNullN(rs.adoColumns(NomeCampo).Value)
    End If
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
Private Sub CONTROLLO_CAMBIO_AGENTE(IDAgenteOLD As Long, IDAgente As Long)
Dim AvviaAggiornamento As Boolean
Dim Testo As String

    AvviaAggiornamento = False
    
    
    If LOADING_NEW_DOC = True Then Exit Sub
    
    
    If IDAgente <> IDAgenteOLD Then
        If oDoc.Tables(sTabellaDettaglio).NumRetails > 0 Then
            AvviaAggiornamento = True
        End If
    End If
    
    Me.CDAgenteTesta.Load IDAgente
    
    'sbImpostaDatiDocumento
    
    'If AvviaAggiornamento = True Then
    '    Testo = "Vuoi aggiornare tutte le righe del documento con l'agente selezionato?"
    '    If MsgBox(Testo, vbQuestion + vbYesNo, "Cambio agente") = vbNo Then Exit Sub
    '
    '    AGGIORNAMENTO_RIGHE_PER_AGENTE Me.CDAgenteTesta.KeyFieldID, Me.CDAgenteTesta.Description, CDAgenteTesta.Code
    'End If
    
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
Private Function GET_SIGLA_UM(IDUM As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ImportoListinoImballo As Double

GET_SIGLA_UM = ""

sSQL = "SELECT * FROM UnitaDiMisura "
sSQL = sSQL & " WHERE IDUnitaDiMisura = " & IDUM

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    
    GET_SIGLA_UM = fnNotNull(rs!DescrizioneFattura)
    
End If

rs.CloseResultset
Set rs = Nothing


End Function
Private Sub RIPRISTINA_GRIGLIA_CONDIZIONI()
Dim I As Long
Dim X As String

Me.BrwMain.Conditions.ClearValues

Me.BrwMain.Conditions.Item("Doc_numero").FromValue = oDoc.Numero
Me.BrwMain.Conditions.Item("Doc_data").FromValue = oDoc.DataEmissione
Me.BrwMain.Conditions.Item("Doc_data").ToValue = oDoc.DataEmissione

BrwMain.ApplyFilter

If (BrwMain.Recordset.EOF = False) Then
    brwMain_DblClick
End If

End Sub

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

Private Sub ACSAnaDest_LostFocus()

    If (ACSAnaDest.IDAnagrafica = fnNotNullN(oDoc.Field("RV_POIDAnagraficaDestinazione", , sTabellaTestata))) Then Exit Sub
    
    sbImpostaDatiDocumento
End Sub

Private Sub txtPesoTotale_LostFocus()
    If (txtPesoTotale.Value = oDoc.Field("Tot_peso", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub
Private Sub txtDataPlafond_LostFocus()
    If (txtDataPlafond.Value = oDoc.Field("Doc_data_plafond", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
End Sub

Private Sub GET_STAMPA_DOCUMENTI_SEL()
On Error GoTo ERR_GET_STAMPA_DOCUMENTI_SEL
Dim Cond As dmtgridctl.dgCondition
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
Private Function REC_DATI_LIQ_FILIALE_LONG(NomeCampo As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo & " FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDSocio=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    REC_DATI_LIQ_FILIALE_LONG = 0
Else
    REC_DATI_LIQ_FILIALE_LONG = fnNotNullN(rs.adoColumns(NomeCampo).Value)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_LINK_PROVINCIA_INTRA(IDFiliale) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDProvinciaIntra FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & IDFiliale & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_LINK_PROVINCIA_INTRA = fnNotNullN(rs!IDProvinciaIntra)
Else
    GET_LINK_PROVINCIA_INTRA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub sbElaborazioneRigheAll()
On Error GoTo ERR_sbElaborazioneRighe
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Quantita_Riga As Double
Dim Quantita_Riga_Conferita As Double
Dim Quantita_Colli As Double
Dim Quantita_PesoLordo As Double
Dim Quantita_Tara As Double
Dim Quantita_PesoNetto As Double
Dim Quantita_Pezzi As Double
Dim TotaleQuantitaColliEla As Long


Screen.MousePointer = 11

sSQL = "SELECT * FROM RV_POTMPArticoliNotaDiCredito "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
'sSQL = sSQL & " AND Selezionato=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

TotaleQuantitaColliEla = 0
While Not rs.EOF
    Quantita_Riga = fnNotNullN(rs!Quantita) ' FormatNumber(((Quantita_da_accreditare / Quantita_Totale_Selezionata) * rs!Quantita), 5)
    Quantita_Riga_Conferita = fnNotNullN(rs!QuantitaConferita) ' FormatNumber((fnNotNullN(rs!QuantitaConferita) / Quantita_Totale_Selezionata) * Quantita_Riga)
    
    'Quantita_Riga = FormatNumber(((Quantita_da_accreditare / Quantita_Totale_Selezionata) * rs!Quantita), 5)
    
    'If Colli_Totali_Selezionati > 0 Then
    '    Quantita_Colli = FormatNumber(((Colli_da_accreditare / Colli_Totali_Selezionati) * rs!Colli), 0)
    'Else
    '    Quantita_Colli = 0
    'End If
    Quantita_Colli = fnNotNullN(rs!Colli) ' FormatNumber(((Colli_da_accreditare / Colli_Totali_Selezionati) * rs!Colli), 0)
    
    
'    If (TotaleQuantitaColliEla + Quantita_Colli) > Colli_da_accreditare Then
'        Quantita_Colli = Colli_da_accreditare - TotaleQuantitaColliEla
'    Else
'        Quantita_Colli = Quantita_Colli
'    End If
'    TotaleQuantitaColliEla = TotaleQuantitaColliEla + Quantita_Colli
    
    
    
'    If PesoLordo_Totali_Selezionati > 0 Then
'        Quantita_PesoLordo = FormatNumber(((PesoLordo_da_accreditare / PesoLordo_Totali_Selezionati) * rs!PesoLordo), 5)
'    Else
'        Quantita_PesoLordo = 0
'    End If
    Quantita_PesoLordo = fnNotNullN(rs!PesoLordo)
    
'    If Tara_Totali_Selezionati > 0 Then
'        Quantita_Tara = FormatNumber(((Tara_da_accreditare / Tara_Totali_Selezionati) * rs!Tara), 5)
'    Else
'        Quantita_Tara = 0
'    End If
    Quantita_Tara = fnNotNullN(rs!Tara)
    
'    If PesoNetto_Totali_Selezionati > 0 Then
'        Quantita_PesoNetto = FormatNumber(((PesoNetto_da_accreditare / PesoNetto_Totali_Selezionati) * rs!PesoNetto), 5)
'    Else
'        Quantita_PesoNetto = 0
'    End If
    Quantita_PesoNetto = fnNotNullN(rs!PesoNetto)
    
'    If Pezzi_Totali_Selezionati > 0 Then
'        Quantita_Pezzi = FormatNumber(((Pezzi_da_accreditare / Pezzi_Totali_Selezionati) * rs!Pezzi), 5)
'    Else
'        Quantita_Pezzi = 0
'    End If
    Quantita_Pezzi = fnNotNullN(rs!Pezzi)
    
    NuovaRigaDocumento_DaElaborazioneAll rs, Quantita_Riga, rs!Quantita, Quantita_Riga_Conferita, Quantita_Colli, Quantita_PesoLordo, Quantita_Tara, Quantita_PesoNetto, Quantita_Pezzi
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Screen.MousePointer = 0

oDoc.PerformTable sTabellaDettaglio, True

'Aggiorna il contenuto della listview degli articoli
sbPopalaListaArticoli False
'Ricalcola il documento
sbCalcolaDocumento

Exit Sub
ERR_sbElaborazioneRighe:
    MsgBox Err.Description, vbCritical, "sbElaborazioneRighe"
    Screen.MousePointer = 0
End Sub

Private Sub NuovaRigaDocumento_DaElaborazioneAll(rs As DmtOleDbLib.adoResultset, Quantita_Riga As Double, QuantitaOriginale As Double, QuantitaConferita As Double, Colli As Double, PesoLordo As Double, Tara As Double, PesoNetto As Double, Pezzi As Double)
Dim ImportoOriginale As Double
Dim Moltiplicatore_Local As Double
Dim IDUMCoop_Local As Long
Dim Sconto1 As Double
Dim Sconto2 As Double
Dim ImportoScontato As Double
ImportoOriginale = 0

    If oDoc.Tables(sTabellaDettaglio).NumRetails = 0 Then
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails
    Else
        oDoc.Tables(sTabellaDettaglio).SetActiveRetail oDoc.Tables(sTabellaDettaglio).NumRetails + 1
    End If
    
    Sconto1 = 0
    Sconto2 = 0
    ImportoOriginale = GET_RIGA_IMPORTO_ORIGINALE_ALL(fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), fnNotNull(rs!CodiceLotto), Sconto1, Sconto2)
    
    oDoc.Field "RV_POPrezzoUnitarioOrigine", ImportoOriginale, sTabellaDettaglio
    
    
    oDoc.Field "Link_Art_articolo", fnNotNullN(rs!IDArticolo), sTabellaDettaglio
    oDoc.Field "Art_codice", rs!CodiceArticolo, sTabellaDettaglio
    oDoc.Field "Art_descrizione", rs!DescrizioneArticolo, sTabellaDettaglio
    
    
    If (rs!RigaRiscontroPeso) = 1 Then
        If (Quantita_Riga) < 0 Then
            Quantita_Riga = Abs(Quantita_Riga)
            ImportoOriginale = -1 * ImportoOriginale
        End If
    End If
    
    oDoc.Field "Art_quantita_totale", Quantita_Riga, sTabellaDettaglio
    oDoc.Field "Art_numero_colli", Colli, sTabellaDettaglio
    oDoc.Field "Art_Peso", PesoLordo, sTabellaDettaglio
    oDoc.Field "Art_tara", Tara, sTabellaDettaglio
    oDoc.Field "Art_quantita_pezzi", Pezzi, sTabellaDettaglio
    
    oDoc.Field "Art_prezzo_unitario_neutro", ImportoOriginale, sTabellaDettaglio
    oDoc.Field "Art_sco_in_percentuale_1", Sconto1, sTabellaDettaglio
    oDoc.Field "Art_sco_in_percentuale_2", Sconto2, sTabellaDettaglio
    
    ImportoScontato = ImportoOriginale - ((ImportoOriginale / 100) * Sconto1)
    ImportoScontato = ImportoScontato - ((ImportoScontato / 100) * Sconto2)
    
    
    
    oDoc.Field "Art_importo_totale_netto_IVA", (ImportoScontato * Quantita_Riga), sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_netto_IVA", ImportoScontato, sTabellaDettaglio
    oDoc.Field "Art_prezzo_unitario_lordo_IVA", ImportoScontato + ((ImportoScontato / 100) * rs!AliquotaIva), sTabellaDettaglio
    oDoc.Field "Art_Importo_totale_neutro", (ImportoScontato * Quantita_Riga), sTabellaDettaglio
    
    oDoc.Field "Art_Importo_netto_IVA", ImportoScontato, sTabellaDettaglio
    
    oDoc.Field "Link_Art_Magazzino", Me.cboMagazzino.CurrentID, sTabellaDettaglio
    oDoc.Field "Link_art_IVA", rs!IDIvaVendita, sTabellaDettaglio
    oDoc.Field "Art_aliquota_IVA", rs!AliquotaIva, sTabellaDettaglio
    
    oDoc.Field "Link_Art_unita_di_misura", rs!IDUnitaDiMisura, sTabellaDettaglio
    oDoc.Field "Art_sigla_unita_di_misura", GET_SIGLA_UM(fnNotNullN(rs!IDUnitaDiMisura)), sTabellaDettaglio
    
    oDoc.Field "Art_intra_non_riporto", Me.chkRiportoIntra_Art.Value, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_nomenclatura", Me.CDIntra_art_Nom_Comb.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_natura_trans", Me.CDIntra_art_Nat_Trans.KeyFieldID, sTabellaDettaglio
    oDoc.Field "Art_intra_qta_tot_massa_netta", Me.txtIntra_Art_MassaNetta.Value, sTabellaDettaglio
    
    If (rs!RigaRiscontroPeso) = 1 Then
        If (Quantita_Riga) < 0 Then
            Quantita_Riga = -Quantita_Riga
            ImportoOriginale = Abs(ImportoOriginale)
            oDoc.Field "Art_intra_qta_tot_massa_netta", -1 * Me.txtIntra_Art_MassaNetta.Value, sTabellaDettaglio
        End If
    End If

    'oDoc.Field "RV_PODescrizioneDocumento", "Rif. D.d.t. n: /" & GetNumeroDocumentoModificato(Me.lngNumero.Value) & " del " & Me.dtData.Text, sTabellaDettaglio
    
    oDoc.Field "RV_PODescrizioneDocumento", rs!DescrizioneDocumento, sTabellaDettaglio
    oDoc.Field "RV_PODataConferimento", rs!DataConferimento, sTabellaDettaglio
    oDoc.Field "RV_POIDConferimentoRighe", rs!IDCollegamentoConferimento, sTabellaDettaglio
    oDoc.Field "RV_POIDAssegnazioneMerce", rs!IDCollegamentoAssegnazioneMerce, sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoIVGamma", rs!IDCollegamentoProcessoLavorazione, sTabellaDettaglio
    oDoc.Field "RV_POIDSocio", fnNotNullN(rs!IDSocio), sTabellaDettaglio
    oDoc.Field "RV_POCodiceSocio", fnNotNull(rs!CodiceSocio), sTabellaDettaglio
    oDoc.Field "RV_POSocio", fnNotNull(rs!Socio), sTabellaDettaglio
    oDoc.Field "RV_PONomeSocio", fnNotNull(rs!NomeSocio), sTabellaDettaglio
    oDoc.Field "RV_POCodiceLotto", fnNotNull(rs!CodiceLotto), sTabellaDettaglio
    oDoc.Field "RV_POIDAnagraficaFatturazione", fnNotNullN(rs!IDAnagraficaFatturazione), sTabellaDettaglio
    'oDoc.Field "RV_PODataLavorazione", fnNotNullN(rs!RV_PODataLavorazione), sTabellaDettaglio

    oDoc.Field "RV_POIDTipoOggetto", rs!IDTipoOggetto, sTabellaDettaglio
    oDoc.Field "RV_POIDOggetto", rs!IDOggetto, sTabellaDettaglio
    oDoc.Field "RV_POIDValoriOggettoDettaglio", rs!IDValoriOggettoDettaglio, sTabellaDettaglio
    oDoc.Field "RV_POIDOggetto_Collegato", rs!IDOggettoCollegato, sTabellaDettaglio
        
    oDoc.Field "RV_POIDTipoLavorazione", rs!IDTipoLavorazione, sTabellaDettaglio
    oDoc.Field "RV_POIDTipoCategoria", rs!IDTipoCategoria, sTabellaDettaglio
    oDoc.Field "RV_POIDCalibro", rs!IDcalibro, sTabellaDettaglio
    oDoc.Field "RV_POQuantitaOrigine", QuantitaOriginale, sTabellaDettaglio
    
    oDoc.Field "RV_POImportoLiq", rs!ImportoUnitario, sTabellaDettaglio

    oDoc.Field "RV_POImportoDaLiq", 0, sTabellaDettaglio
    Moltiplicatore_Local = GET_MOLTIPLICATORE_ARTICOLO(fnNotNullN(rs!IDArticolo))
    oDoc.Field "RV_POQuantitaLiq", Quantita_Riga * Moltiplicatore_Local, sTabellaDettaglio
    IDUMCoop_Local = GET_UM_COOP_ARTICOLO(fnNotNullN(rs!IDArticolo))
    If (IDUMCoop_Local > 0) Then
        Select Case (IDUMCoop_Local)
            Case 1
                oDoc.Field "RV_POQuantitaLiq", Colli * Moltiplicatore_Local, sTabellaDettaglio
            Case 2
                oDoc.Field "RV_POQuantitaLiq", PesoLordo * Moltiplicatore_Local, sTabellaDettaglio
            Case 3
                oDoc.Field "RV_POQuantitaLiq", Tara * Moltiplicatore_Local, sTabellaDettaglio
            Case 4
                oDoc.Field "RV_POQuantitaLiq", PesoNetto * Moltiplicatore_Local, sTabellaDettaglio
            Case 5
                oDoc.Field "RV_POQuantitaLiq", Pezzi * Moltiplicatore_Local, sTabellaDettaglio
        End Select
    End If
    oDoc.Field "RV_POPrezzoMedioInLiq", GET_VALORE_PREZZO_MEDIO_RIGA_ORIGINALE(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNull(rs!CodiceLotto)), sTabellaDettaglio
    oDoc.Field "RV_POImportoMerceNetta", ImportoOriginale, sTabellaDettaglio
    oDoc.Field "RV_POVariazionePrezzoImballo", 0, sTabellaDettaglio
    oDoc.Field "RV_POIDIvaImballo", 0, sTabellaDettaglio
    oDoc.Field "RV_PORigaRiscontroPeso", fnNotNullN(rs!RigaRiscontroPeso), sTabellaDettaglio
    If (TIPO_VARIAZIONE_DA_WIZARD = 0) Then
        oDoc.Field "RV_POIDTipoVariazione", GET_TIPO_VARIAZIONE_RIGA_DOCUMENTO(ImportoOriginale, QuantitaOriginale, Importo_da_accreditare, Quantita_Riga), sTabellaDettaglio
    Else
        oDoc.Field "RV_POIDTipoVariazione", TIPO_VARIAZIONE_DA_WIZARD, sTabellaDettaglio
    End If
    oDoc.Field "RV_POIDTipoImportoVenditaLiq", GET_VALORE_FORZA_TIPO_PREZZO_LIQ(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNull(rs!CodiceLotto)), sTabellaDettaglio
    
    'INSERIRE LA PEDANA E IL PESO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    oDoc.Field "RV_PODataLavorazione", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_PODataLavorazione", True), sTabellaDettaglio
    oDoc.Field "RV_POCodicePedana", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_POCodicePedana", True), sTabellaDettaglio
    oDoc.Field "RV_POIDTipoPedana", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_POIDTipoPedana", False), sTabellaDettaglio
    oDoc.Field "RV_POPesoPedana", GET_VALORE_CAMPO_RIGA_VENDITA(fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio), "RV_POPesoPedana", False), sTabellaDettaglio
    oDoc.Field "RV_POLottoCampagna", GET_LOTTO_PRODUZIONE(fnNotNullN(rs!IDCollegamentoConferimento)), sTabellaDettaglio
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    oDoc.Field "RV_POID_Art_dettaglio_prog", fnNotNullN(rs!ID_Art_dettaglio_prog), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazione", fnNotNullN(rs!IDRV_POProcessoLavorazione), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazioneRighe", fnNotNullN(rs!IDRV_POProcessoLavorazioneRighe), sTabellaDettaglio
    oDoc.Field "RV_POIDLineaProduzione", fnNotNullN(rs!IDRV_POLineaProduzione), sTabellaDettaglio
    oDoc.Field "RV_POIDTipoUtilizzoLinea", fnNotNullN(rs!IDRV_POTipoUtilizzoLinea), sTabellaDettaglio

    GET_INTRAST_RIGA_ARTICOLO fnNotNullN(rs!IDArticolo), TheApp.IDFirm, TheApp.Branch, oDoc.Field("Art_quantita_totale", , sTabellaDettaglio)
    SCRIVI_AGENTE_DA_RIGA_DOC fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio)
    SCRIVI_PDC_DA_RIGA_DOC fnNotNullN(rs!IDTipoOggetto), fnNotNullN(rs!IDOggetto), fnNotNullN(rs!IDValoriOggettoDettaglio)

    NumeroProgSingolaRiga = NumeroProgSingolaRiga + 1
    
    oDoc.Field "RV_POLinkRiga", NumeroProgSingolaRiga, sTabellaDettaglio
    oDoc.Field "RV_POTipoRiga", 1, sTabellaDettaglio
    oDoc.Field "RV_POID_Art_dettaglio_prog", fnNotNullN(rs!ID_Art_dettaglio_prog), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazione", fnNotNullN(rs!IDRV_POProcessoLavorazione), sTabellaDettaglio
    oDoc.Field "RV_POIDProcessoLavorazioneRighe", fnNotNullN(rs!IDRV_POProcessoLavorazioneRighe), sTabellaDettaglio
    oDoc.Field "RV_POIDLineaProduzione", fnNotNullN(rs!IDRV_POLineaProduzione), sTabellaDettaglio
    oDoc.Field "RV_POIDTipoUtilizzoLinea", fnNotNullN(rs!IDRV_POTipoUtilizzoLinea), sTabellaDettaglio
    
    oDoc.Field "Art_riferimento_PA", GET_RIF_PA_ARTICOLO(Me.CDArticolo.KeyFieldID, Me.cdAnagrafica.KeyFieldID, Me.cboAltroSito.CurrentID), sTabellaDettaglio
    sbLoadElectronicInvoiceData4Article fnNotNullN(oDoc.Field("ID_Art_dettaglio_prog", , sTabellaDettaglio)), fnNotNullN(oDoc.Field("Link_Art_articolo", , sTabellaDettaglio))
    SET_INTRASTAT_RIGA_DOCUMENTO fnNotNullN(rs!IDArticolo), Quantita_Riga
End Sub

Private Function GET_RIGA_IMPORTO_ORIGINALE_ALL(IDOggettoCollegato As Long, IDTipoOggettoCollegato As Long, IDValoriOggettoDettaglioCollegato As Long, CodiceLottoVendita As String, Sconto1 As Double, Sconto2 As Double) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_RIGA_IMPORTO_ORIGINALE_ALL = 0

Select Case IDTipoOggettoCollegato
'    Case 114
'        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA "
'        sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
'
'    Case 2
'        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA "
'        sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
'
'    Case 8
'        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA "
'        sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
    
    Case 114
        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA, Art_sco_in_percentuale_1, Art_sco_in_percentuale_2, Art_prezzo_unitario_neutro "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "

    Case 2
        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA, Art_sco_in_percentuale_1, Art_sco_in_percentuale_2, Art_prezzo_unitario_neutro "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "

    Case 8
        sSQL = "SELECT Art_pre_uni_net_sco_net_IVA, Art_sco_in_percentuale_1, Art_sco_in_percentuale_2, Art_prezzo_unitario_neutro "
        sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
End Select

sSQL = sSQL & "WHERE IDOggetto=" & IDOggettoCollegato

If COLLEGAMENTO_NOTA_PER_LOTTO = 1 Then
    sSQL = sSQL & " AND RV_POCodiceLotto=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglioCollegato
End If

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    'GET_RIGA_IMPORTO_ORIGINALE_ALL = fnNotNullN(rs!Art_pre_uni_net_sco_net_IVA)
    GET_RIGA_IMPORTO_ORIGINALE_ALL = fnNotNullN(rs!Art_prezzo_unitario_neutro)
    Sconto1 = fnNotNullN(rs!Art_sco_in_percentuale_1)
    Sconto2 = fnNotNullN(rs!Art_sco_in_percentuale_2)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_CONTROLLO_TIPO_COMM_ND_NC(IDTipoCommissione As Long, IDTipoOggetto As Long) As Long
On Error GoTo ERR_GET_CONTROLLO_TIPO_COMM_ND_NC
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_TIPO_COMM_ND_NC = 0
If IDTipoOggetto = 2 Then Exit Function
If IDTipoOggetto = 8 Then Exit Function
If IDTipoOggetto = 114 Then Exit Function

sSQL = "SELECT * FROM RV_POTipoCommissione "
sSQL = sSQL & "WHERE IDRV_POTipoCommissione=" & IDTipoCommissione

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CONTROLLO_TIPO_COMM_ND_NC = fnNotNullN(rs!NonCalcolareCommissioniInNDeNC)
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_CONTROLLO_TIPO_COMM_ND_NC:
    MsgBox Err.Description, vbCritical, "GET_CONTROLLO_TIPO_COMM_ND_NC"
End Function
Private Sub chkNomIva_Click()
    If (chkNomIva.Value = oDoc.Field("Nom_IVA_default", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub
Private Sub chkAddebitaBollo_Click()
    If (chkAddebitaBollo.Value = oDoc.Field("Nom_bollo_esente", , sTabellaTestata)) Then Exit Sub
    sbImpostaDatiDocumento
    sbCalcolaDocumento
End Sub
Private Function GET_LOTTO_PRODUZIONE(IDRigaConferimento As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_LOTTO_PRODUZIONE = ""

sSQL = "SELECT IDRV_POCaricoMerceRighe, LottoDiConferimento"
sSQL = sSQL & " FROM RV_POCaricoMerceRighe "
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LOTTO_PRODUZIONE = fnNotNull(rs!LottoDiConferimento)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub sbLoadElectronicInvoiceData4Article(ByVal lID_Art_dettaglio_prog As Long, ByVal lIDArticle As Long)
On Error GoTo ERR_sbLoadElectronicInvoiceData4Article
    Dim oRs As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFilter As String

    'DATI
    'leggo i dati legati all'articolo con IDArticolo richiesto
    Set oRs = oDoc.ElectronicInvoiceAdditionalData.GetDataFromArticle(lIDArticle)
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then
            If oRs.RecordCount > 0 Then
                With oDoc.ElectronicInvoiceAdditionalData.AdditionalData
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
                        
                        .Fields("IDOggetto").Value = oDoc.IDOggetto
                        .Fields("IDTipoOggetto").Value = oDoc.IDTipoOggetto
                        .Fields("ID_Art_dettaglio_prog").Value = lID_Art_dettaglio_prog
                        .Fields("Eliminato").Value = False
                        'Impostare "Temporaneo" a False se i codici vengono immediatamente legati al dettaglio,
                        'a True se questi codici rimangono sospesi in attesa di un ulteriore conferma
                        '(Salva di dettaglio, ad es., i cui Temporaneo verrà finalmente posto a False)
                        .Fields("Temporaneo").Value = False
                        
                        oRs.MoveNext
                    Wend
                    oDoc.ElectronicInvoiceAdditionalData.Changed = True
                End With
            End If
            oRs.Close
        End If
    End If
    Set oRs = Nothing
    
    'CODICI
    'leggo i codici legati all'articolo con IDArticolo richiesto
    Set oRs = oDoc.ElectronicInvoiceAdditionalData.GetCodesFromArticle(lIDArticle)
    If Not oRs Is Nothing Then
        If oRs.State = adStateOpen Then
            If oRs.RecordCount > 0 Then
                With oDoc.ElectronicInvoiceAdditionalData.AdditionalCodes
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
                        
                        .Fields("IDOggetto").Value = oDoc.IDOggetto
                        .Fields("IDTipoOggetto").Value = oDoc.IDTipoOggetto
                        .Fields("ID_Art_dettaglio_prog").Value = lID_Art_dettaglio_prog
                        .Fields("Eliminato").Value = False
                        'Impostare "Temporaneo" a False se i codici vengono immediatamente legati al dettaglio,
                        'a True se questi codici rimangono sospesi in attesa di un ulteriore conferma
                        '(Salva di dettaglio, ad es., i cui Temporaneo verrà finalmente posto a False)
                        .Fields("Temporaneo").Value = False
                        
                        oRs.MoveNext
                    Wend
                    oDoc.ElectronicInvoiceAdditionalData.Changed = True
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
Private Sub ConsolidaDettaglioFatturaElettronica()
On Error GoTo ERR_ConsolidaDettaglioFatturaElettronica
    Dim oRs As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFilter As String
    
    'DATI
    'leggo i dati legati al documento
    With oDoc.ElectronicInvoiceAdditionalData.AdditionalData
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
    With oDoc.ElectronicInvoiceAdditionalData.AdditionalCodes
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
    
    oDoc.ElectronicInvoiceAdditionalData.Changed = True

Exit Sub
ERR_ConsolidaDettaglioFatturaElettronica:
    MsgBox Err.Description, vbCritical, "ConsolidaDettaglioFatturaElettronica"
End Sub
Private Sub Command3_Click()
On Error GoTo ERR_Command3_Click
    If BrwMain.Visible = True Then Exit Sub
    
    With oDoc.ElectronicInvoiceAdditionalDataHeader
        .Show
        If .Changed Then
            If Not (BrwMain.Visible) Then Change
        End If
    End With
Exit Sub
ERR_Command3_Click:
    MsgBox Err.Description, vbCritical, "Dati e-fattura documento"
End Sub

Private Sub Command4_Click()
On Error GoTo ERR_Command3_Click
Dim sArt_riferimentoPA As String

If BrwMain.Visible = True Then Exit Sub
If Tipo_Riga_Sel_EF = 1 Then sArt_riferimentoPA = Rif_PA_Riga_Doc_Merce
If Tipo_Riga_Sel_EF = 2 Then sArt_riferimentoPA = Rif_PA_Riga_Doc_Imballo

With oDoc.ElectronicInvoiceAdditionalData
    .Show Sel_IDArt_dettaglio, sArt_riferimentoPA
    If .Changed Then
        If Tipo_Riga_Sel_EF = 1 Then Rif_PA_Riga_Doc_Merce = sArt_riferimentoPA
        If Tipo_Riga_Sel_EF = 2 Then Rif_PA_Riga_Doc_Imballo = sArt_riferimentoPA
    End If
End With
Exit Sub
ERR_Command3_Click:
    MsgBox Err.Description, vbCritical, "Dati e-fattura riga documento"
End Sub
Private Function GET_RIF_PA_ARTICOLO(IDArticolo As Long, IDCliente As Long, IDDestinazione As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_RIF_PA_ARTICOLO = ""

'Articolo - Cliente
sSQL = "SELECT RiferimentoPACliente FROM ClientePerArticolo "
sSQL = sSQL & " WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAnagrafica=" & IDCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPACliente)
End If
rs.CloseResultset
Set rs = Nothing

If Len(Trim(GET_RIF_PA_ARTICOLO)) > 0 Then Exit Function

'Destinazione
sSQL = "SELECT RiferimentoPAArticolo FROM SitoPerAnagrafica "
sSQL = sSQL & " WHERE IDSitoPerAnagrafica=" & IDDestinazione

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

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

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_RIF_PA_ARTICOLO = fnNotNull(rs!RiferimentoPA)
End If
rs.CloseResultset
Set rs = Nothing

End Function
Private Sub EliminaDettaglioFatturaElettronica(IDArtProg As Long)
On Error GoTo ERR_ConsolidaDettaglioFatturaElettronica
    Dim oRs As ADODB.Recordset
    Dim oField As ADODB.Field
    Dim sFilter As String
    
    'DATI
    'leggo i dati legati al documento
    With oDoc.ElectronicInvoiceAdditionalData.AdditionalData
        'conservo il filtro per rimetterlo dopo
        If Len(.Filter) > 0 Then sFilter = .Filter
        
        .Filter = "ID_Art_dettaglio_prog=" & IDArtProg
        If .RecordCount > 0 Then
            While Not .EOF
                If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                    .Fields("Eliminato").Value = True
                End If
                .MoveNext
            Wend
        End If
        .Filter = sFilter
    End With
    'DATI
    'leggo i dati legati al documento
    With oDoc.ElectronicInvoiceAdditionalData.AdditionalCodes
        'conservo il filtro per rimetterlo dopo
        If Len(.Filter) > 0 Then sFilter = .Filter
        
        .Filter = "ID_Art_dettaglio_prog=" & IDArtProg
        If .RecordCount > 0 Then
            While Not .EOF
                If fnNotNullL(.Fields("Eliminato").Value) = 0 Then
                    .Fields("Eliminato").Value = True
                End If
                .MoveNext
            Wend
        End If
        .Filter = sFilter
    End With
    
    oDoc.ElectronicInvoiceAdditionalData.Changed = True

Exit Sub
ERR_ConsolidaDettaglioFatturaElettronica:
    MsgBox Err.Description, vbCritical, "ConsolidaDettaglioFatturaElettronica"
End Sub

Private Function GET_INVIO_EMAIL_PERSONALIZZATA(IDUtente As Long) As Boolean
On Error GoTo ERR_GET_INVIO_EMAIL_PERSONALIZZATA
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUtente, eMail_ServerSMTP "
sSQL = sSQL & "FROM Utente "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND eMail_ProtocolloInvio=1"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_INVIO_EMAIL_PERSONALIZZATA = False
Else
    If (Len(fnNotNull(rs!eMail_ServerSMTP)) > 0) Then
    GET_INVIO_EMAIL_PERSONALIZZATA = True
    Else
        GET_INVIO_EMAIL_PERSONALIZZATA = False
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_INVIO_EMAIL_PERSONALIZZATA:
    MsgBox Err.Description, vbCritical, "GET_INVIO_EMAIL_PERSONALIZZATA"
End Function
Private Function GET_INDIRIZZO_EMAIL_CLIENTE(IDAnagrafica As Long) As String
On Error GoTo ERR_GET_INDIRIZZO_EMAIL_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagrafica, EmailInternet "
sSQL = sSQL & "FROM Anagrafica "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_INDIRIZZO_EMAIL_CLIENTE = ""
Else
    GET_INDIRIZZO_EMAIL_CLIENTE = fnNotNull(rs!EmailInternet)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_INDIRIZZO_EMAIL_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_INDIRIZZO_EMAIL_CLIENTE"
End Function
Private Function GET_NOME_FILE_DOCUMENTO() As String

Dim NomeFile As String

GET_NOME_FILE_DOCUMENTO = ""

GET_NOME_FILE_DOCUMENTO = oDoc.Descrizione & ""
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " " & fnNotNull(oDoc.Field("Nom_ragione_sociale_o_cognome", , sTabellaTestata)) & fnNotNull(oDoc.Field("Nom_nome", , sTabellaTestata))
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " (" & fnNotNull(oDoc.Field("Nom_codice", , sTabellaTestata)) & ")"
GET_NOME_FILE_DOCUMENTO = GET_NOME_FILE_DOCUMENTO & " [" & GET_DATA_FORMATTATA(oDoc.DataEmissione) & "]"
    


End Function
Private Function GET_DATA_FORMATTATA(DataF As String) As String
On Error GoTo ERR_GET_DATA_FORMATTATA
Dim Anno As String
Dim mese As String
Dim giorno As String

GET_DATA_FORMATTATA = ""

Anno = Year(DataF)
mese = Month(DataF)
giorno = Day(DataF)

If Len(mese) = 1 Then mese = "0" & mese
If Len(giorno) = 1 Then giorno = "0" & giorno

GET_DATA_FORMATTATA = Anno & "-" & mese & "-" & giorno

GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & " n. "

If Len(Trim(fnNotNull(oDoc.Field("Doc_prefisso", , sTabellaTestata)))) > 0 Then
    GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & Trim(fnNotNull(oDoc.Field("Doc_prefisso", , sTabellaTestata))) & "-"
End If

GET_DATA_FORMATTATA = GET_DATA_FORMATTATA & oDoc.Numero

Exit Function
ERR_GET_DATA_FORMATTATA:
End Function

Private Sub SET_INTRASTAT_RIGA_DOCUMENTO(IDArticolo As Long, Quantita As Double)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDArticolo,  "
sSQL = sSQL & "NonRiportoIntrastat, MassaNettaInKg, IDNomenclaturaCombinata "
sSQL = sSQL & "FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & "Iva ON Articolo.IDIvaVendita = Iva.IDIva "
sSQL = sSQL & "WHERE IDArticolo = " & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    oDoc.Field "Art_intra_non_riporto", rs!NonRiportoIntrastat, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_nomenclatura", rs!IDNomenclaturaCombinata, sTabellaDettaglio
    oDoc.Field "Link_Art_intra_natura_trans", 2, sTabellaDettaglio
    oDoc.Field "Art_intra_qta_tot_massa_netta", Quantita * fnNotNullN(rs!MassaNettaInKg), sTabellaDettaglio
End If

rs.CloseResultset
Set rs = Nothing


End Sub
Private Function GET_UM_COOP_ARTICOLO(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
GET_UM_COOP_ARTICOLO = 0

sSQL = "SELECT RV_POIDUnitaDiMisuraLiq FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_UM_COOP_ARTICOLO = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
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
        
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_LINK_UM_LIQUIDAZIONE_ARTICOLO = fnNotNullN(rs!RV_POIDUnitaDiMisuraLiq)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function SET_VARIAZIONE_LIQ_DOC_ORI(IDOggetto As Long, IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long, CodiceLottoVendita As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

SET_VARIAZIONE_LIQ_DOC_ORI = 0

sSQL = "SELECT IDMovimento, RV_POMerceInclusaImballo "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto
If (COLLEGAMENTO_NOTA_PER_LOTTO = 1) Then
    sSQL = sSQL & " AND RV_POCodiceLottoVendita=" & fnNormString(CodiceLottoVendita)
Else
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
End If
sSQL = sSQL & " AND RV_POTipoRiga=1"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    SET_VARIAZIONE_LIQ_DOC_ORI = fnNotNullN(rs!RV_POMerceInclusaImballo)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_NUMERO_PROCESSO(IDProcessoIVGamma As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT NumeroProcesso, AnnoProcesso "
sSQL = sSQL & "FROM RV_POProcessoIVGamma "
sSQL = sSQL & "WHERE IDRV_POProcessoIVGamma=" & IDProcessoIVGamma

Set rs = Cn.OpenResultset(sSQL)


If rs.EOF Then
    GET_NUMERO_PROCESSO = ""
Else
    GET_NUMERO_PROCESSO = fnNotNull(rs!AnnoProcesso) & "-" & fnNotNull(rs!NumeroProcesso)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_GESTIONE_INCASSI(IDOggetto As Long, IDTipoOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDOggettoScadenza As Long

sSQL = "SELECT FlussoOggettiCollegati.IDFlussoFunzione, FlussoOggettiCollegati.IDTipoOggetto, FlussoOggettiCollegati.IDOggetto, "
sSQL = sSQL & "FlussoOggettiCollegati.IDTipoOggettoCollegato, FlussoOggettiCollegati.IDOggettoCollegato, "
sSQL = sSQL & "FlussoOggettiCollegati_1.IDOggettoCollegato AS IDOggettoCollegatoIncasso, "
sSQL = sSQL & "FlussoOggettiCollegati_1.IDTipoOggettoCollegato AS IDTipoOggettoCollegatoIncasso "
sSQL = sSQL & "FROM FlussoOggettiCollegati INNER JOIN "
sSQL = sSQL & "FlussoOggettiCollegati AS FlussoOggettiCollegati_1 ON FlussoOggettiCollegati.IDOggettoCollegato = FlussoOggettiCollegati_1.IDOggetto AND "
sSQL = sSQL & "FlussoOggettiCollegati.IDTipoOggettoCollegato = FlussoOggettiCollegati_1.IDTipoOggetto "
sSQL = sSQL & "WHERE FlussoOggettiCollegati.IDOggetto = " & IDOggetto
sSQL = sSQL & " AND FlussoOggettiCollegati.IDTipoOggetto = " & IDTipoOggetto
sSQL = sSQL & " AND FlussoOggettiCollegati.IDTipoOggettoCollegato = 131 "

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_GESTIONE_INCASSI = False
Else
    GET_CONTROLLO_GESTIONE_INCASSI = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub txtSconto1_Change()
    CalcolaImportoScontato
    CalcolaTotaleRiga
End Sub

Private Sub txtSconto1_LostFocus()
    'Me.txtPrezzoNettoMerce.Value = Me.txtImponibileUnitario.Value - Me.txtVariazionePrezzoImballo.Value
End Sub

Private Sub txtSconto2_Change()
    CalcolaImportoScontato
    CalcolaTotaleRiga
End Sub

Private Sub txtSconto2_LostFocus()
    'Me.txtPrezzoNettoMerce.Value = Me.txtImponibileUnitario.Value - Me.txtVariazionePrezzoImballo.Value
End Sub
Private Sub Gestione_Altri_dati()

frmAltriDati.Show vbModal

If CONFERMA_ALTRI_DATI = 1 Then sbImpostaDatiDocumento

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

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_NOTA_DOCUMENTO = fnNotNull(rs!Annotazione)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_NOTA_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "GET_NOTA_DOCUMENTO"
End Function
Private Sub SCRIVI_RIFERIMENTI(NomeTabella As String)
On Error GoTo ERR_SCRIVI_RIFERIMENTI
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsOrdini As ADODB.Recordset
Dim rsFatture As ADODB.Recordset
Dim rsDDT As ADODB.Recordset
Dim SplitRifDocumento() As String
Dim NumeroDocumento As String
Dim DataDocumento As String
Dim rsDDTSel As DmtOleDbLib.adoResultset

Set rsOrdini = New ADODB.Recordset
rsOrdini.CursorLocation = adUseClient
rsOrdini.Fields.Append "NumeroDocumento", adVarChar, 50, adFldIsNullable
rsOrdini.Fields.Append "DataDocumento", adDBDate, , adFldIsNullable
'rsOrdini.Fields.Append "DataDocumento", adVarChar, 50, adFldIsNullable
rsOrdini.Fields.Append "NumeroRiferimento", adInteger, , adFldIsNullable
rsOrdini.Fields.Append "IDRiferimento", adInteger, , adFldIsNullable
rsOrdini.Open , , adOpenKeyset, adLockBatchOptimistic

Set rsFatture = New ADODB.Recordset
rsFatture.CursorLocation = adUseClient
rsFatture.Fields.Append "NumeroDocumento", adVarChar, 50, adFldIsNullable
rsFatture.Fields.Append "DataDocumento", adDBDate, , adFldIsNullable
rsFatture.Fields.Append "NumeroRiferimento", adInteger, , adFldIsNullable
rsFatture.Fields.Append "IDRiferimento", adInteger, , adFldIsNullable

rsFatture.Open , , adOpenKeyset, adLockBatchOptimistic

Set rsDDT = New ADODB.Recordset
rsDDT.CursorLocation = adUseClient
rsDDT.Fields.Append "NumeroDocumento", adVarChar, 50, adFldIsNullable
rsDDT.Fields.Append "DataDocumento", adDBDate, , adFldIsNullable
rsDDT.Fields.Append "NumeroRiferimento", adInteger, , adFldIsNullable
rsDDT.Fields.Append "IDRiferimento", adInteger, , adFldIsNullable

rsDDT.Open , , adOpenKeyset, adLockBatchOptimistic


sSQL = "SELECT IDValoriOggettoDettaglio, IDOggetto, IDTipoOggetto, "
sSQL = sSQL & "RV_POIDOggetto, RV_POIDTipoOggetto, RV_POIDValoriOggettoDettaglio, RV_PODescrizioneDocumento, RV_POOggettoOrdineCliente, "
sSQL = sSQL & "ID_Art_dettaglio_prog "
sSQL = sSQL & " FROM " & NomeTabella
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    NumeroDocumento = ""
    DataDocumento = ""
    '''FATTURE COLLEGATE
    If Len(Trim(fnNotNull(rs!RV_PODescrizioneDocumento))) > 0 Then
        SplitRifDocumento = Split(Trim(fnNotNull(rs!RV_PODescrizioneDocumento)), "n:")
        If (UBound(SplitRifDocumento) > 0) Then
            NumeroDocumento = ""
            DataDocumento = ""
            
            SplitRifDocumento = Split(SplitRifDocumento(1), "del")
            NumeroDocumento = Trim(SplitRifDocumento(0))
            If UBound(SplitRifDocumento) > 0 Then DataDocumento = Trim(SplitRifDocumento(1))
                        
            If Len(NumeroDocumento) > 0 Then
                If RIPORTA_RIF_DETTAGLIATI = 0 Then
                    rsFatture.Filter = "NumeroDocumento=" & fnNormString(NumeroDocumento)
                    If rsFatture.EOF Then
                        rsFatture.AddNew
                        rsFatture!NumeroDocumento = NumeroDocumento
                        rsFatture!DataDocumento = DataDocumento
                        rsFatture.Update
                    End If
                    rsFatture.Filter = vbNullString
                Else
                    rsFatture.AddNew
                    rsFatture!NumeroDocumento = NumeroDocumento
                    rsFatture!DataDocumento = DataDocumento
                    rsFatture!NumeroRiferimento = fnNotNullN(rs!ID_Art_dettaglio_prog)
                    rsFatture!IDRiferimento = 0
                    rsFatture.Update
                End If
            End If
        End If
    End If
    
    'RIGA COLLEGATA AD UN DOCUMENTO DI TRASPORTO
    If fnNotNullN(rs!RV_POIDTipoOggetto = 2) Then
        sSQL = "SELECT IDOggetto, Doc_data, Doc_prefisso, Doc_numero, Doc_numero_vs_ordine_di_rifer, Doc_data_vs_ordine_di_rifer "
        sSQL = sSQL & "FROM ValoriOggettoPerTipo0002 "
        sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!RV_POIDOggetto)
        
        Set rsDDTSel = Cn.OpenResultset(sSQL)
        
        If Not rsDDTSel.EOF Then
            'DOCUMENTI DI TRASPORTO COLLEGATI
            If RIPORTA_RIF_DETTAGLIATI = 0 Then
                rsDDT.Filter = "NumeroDocumento=" & fnNormString(rsDDTSel!Doc_Numero)
                If rsDDT.EOF Then
                    rsDDT.AddNew
                    rsDDT!NumeroDocumento = rsDDTSel!Doc_Numero
                    If Len(Trim(rsDDTSel!Doc_prefisso)) > 0 Then
                        rsDDT!NumeroDocumento = Trim(rsDDTSel!Doc_prefisso) & "/" & rsDDT!NumeroDocumento
                    End If
                    rsDDT!DataDocumento = rsDDTSel!Doc_data
                    rsDDT.Update
                End If
                rsDDT.Filter = vbNullString
            Else
                rsDDT.AddNew
                rsDDT!NumeroDocumento = rsDDTSel!Doc_Numero
                If Len(Trim(rsDDTSel!Doc_prefisso)) > 0 Then
                    rsDDT!NumeroDocumento = Trim(rsDDTSel!Doc_prefisso) & "/" & rsDDT!NumeroDocumento
                End If
                rsDDT!DataDocumento = rsDDTSel!Doc_data
                rsDDT!NumeroRiferimento = fnNotNullN(rs!ID_Art_dettaglio_prog)
                rsDDT!IDRiferimento = 0
                rsDDT.Update
            End If
            'ORDINI COLLEGATI
            If Len(fnNotNull(rsDDTSel!Doc_numero_vs_ordine_di_rifer)) > 0 Then
                If RIPORTA_RIF_DETTAGLIATI = 0 Then
                    rsOrdini.Filter = "NumeroDocumento=" & fnNormString(rsDDTSel!Doc_numero_vs_ordine_di_rifer)
                    If rsDDT.EOF Then
                        rsOrdini.AddNew
                        rsOrdini!NumeroDocumento = rsDDTSel!Doc_numero_vs_ordine_di_rifer
                        rsOrdini!DataDocumento = rsDDTSel!Doc_data_vs_ordine_di_rifer
                        rsOrdini.Update
                    End If
                    rsOrdini.Filter = vbNullString
                Else
                    rsOrdini.AddNew
                    rsOrdini!NumeroDocumento = rsDDTSel!Doc_numero_vs_ordine_di_rifer
                    rsOrdini!DataDocumento = rsDDTSel!Doc_data_vs_ordine_di_rifer
                    rsOrdini!NumeroRiferimento = fnNotNullN(rs!ID_Art_dettaglio_prog)
                    rsOrdini!IDRiferimento = 0
                    rsOrdini.Update
                End If
            End If
        End If
        
        rsDDTSel.CloseResultset
        Set rsDDTSel = Nothing
    End If
    
    'RIGA COLLEGATA AD UNA FATTURA ACCOMPAGNATORIA
    If fnNotNullN(rs!RV_POIDTipoOggetto = 114) Then
        sSQL = "SELECT IDOggetto, Doc_data, Doc_prefisso, Doc_numero, Doc_numero_vs_ordine_di_rifer, Doc_data_vs_ordine_di_rifer "
        sSQL = sSQL & "FROM ValoriOggettoPerTipo0072 "
        sSQL = sSQL & "WHERE IDOggetto=" & fnNotNullN(rs!RV_POIDOggetto)
        
        Set rsDDTSel = Cn.OpenResultset(sSQL)
        
        If Not rsDDTSel.EOF Then
            'ORDINI COLLEGATI
            If Len(fnNotNull(rsDDTSel!Doc_numero_vs_ordine_di_rifer)) > 0 Then
                If RIPORTA_RIF_DETTAGLIATI = 0 Then
                    rsOrdini.Filter = "NumeroDocumento=" & fnNormString(rsDDTSel!Doc_numero_vs_ordine_di_rifer)
                    If rsDDT.EOF Then
                        rsOrdini.AddNew
                        rsOrdini!NumeroDocumento = rsDDTSel!Doc_numero_vs_ordine_di_rifer
                        rsOrdini!DataDocumento = rsDDTSel!Doc_data_vs_ordine_di_rifer
                        rsOrdini.Update
                    End If
                    rsOrdini.Filter = vbNullString
                Else
                    rsOrdini.AddNew
                    rsOrdini!NumeroDocumento = rsDDTSel!Doc_numero_vs_ordine_di_rifer
                    rsOrdini!DataDocumento = rsDDTSel!Doc_data_vs_ordine_di_rifer
                    rsOrdini!NumeroRiferimento = fnNotNullN(rs!ID_Art_dettaglio_prog)
                    rsOrdini!IDRiferimento = 0
                    rsOrdini.Update
                End If
            End If
        End If
        
        rsDDTSel.CloseResultset
        Set rsDDTSel = Nothing
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

If Not ((rsFatture.EOF) And (rsFatture.BOF)) Then
    rsFatture.MoveFirst
    While Not rsFatture.EOF
        SCRIVI_ORD_CLI_RIF_XML fnNotNull(rsFatture!DataDocumento), fnNotNull(rsFatture!NumeroDocumento), 5, fnNotNullN(rsFatture!NumeroRiferimento), fnNotNullN(rsFatture!IDRiferimento)
    rsFatture.MoveNext
    Wend
End If

If Not ((rsOrdini.EOF) And (rsOrdini.BOF)) Then
    rsOrdini.MoveFirst
    While Not rsOrdini.EOF
        SCRIVI_ORD_CLI_RIF_XML fnNotNull(rsOrdini!DataDocumento), fnNotNull(rsOrdini!NumeroDocumento), 1, fnNotNullN(rsOrdini!NumeroRiferimento), fnNotNullN(rsOrdini!IDRiferimento)
    rsOrdini.MoveNext
    Wend
End If

If Not ((rsDDT.EOF) And (rsDDT.BOF)) Then
    rsDDT.MoveFirst
    While Not rsDDT.EOF
        SCRIVI_ORD_CLI_RIF_XML fnNotNull(rsDDT!DataDocumento), fnNotNull(rsDDT!NumeroDocumento), 6, fnNotNullN(rsDDT!NumeroRiferimento), fnNotNullN(rsDDT!IDRiferimento)
    rsDDT.MoveNext
    Wend
End If

rsFatture.Close
Set rsFatture = Nothing

rsOrdini.Close
Set rsOrdini = Nothing

rsDDT.Close
Set rsDDT = Nothing
Exit Sub
ERR_SCRIVI_RIFERIMENTI:
    MsgBox Err.Description, vbCritical, "SCRIVI_RIFERIMENTI"
End Sub
Private Sub SCRIVI_ORD_CLI_RIF_XML(DataOrdine As String, NumeroOrdine As String, IDBlocco As Long, NumeroRiferimentoRiga As Long, IDDocumentoRiga As Long)
On Error GoTo ERR_SCRIVI_ORD_CLI_RIF_XML
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim Avvia As Boolean
Dim rs As DmtOleDbLib.adoResultset

Avvia = False

sSQL = "SELECT IDDatoFatturaPATestataDoc FROM DatoFatturaPATestataDoc "
sSQL = sSQL & " WHERE IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
sSQL = sSQL & " AND IDBloccoXML=" & IDBlocco
If IDBlocco <> 6 Then
    sSQL = sSQL & " AND NumItem=" & fnNormString(NumeroOrdine)
Else
    sSQL = sSQL & " AND NumeroDDT=" & fnNormString(NumeroOrdine)
End If
If NumeroRiferimentoRiga > 0 Then
     sSQL = sSQL & " AND RiferimentoNumeroLinea=" & NumeroRiferimentoRiga
End If

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Avvia = True
End If

rs.CloseResultset
Set rs = Nothing

If Avvia = True Then
    sSQL = "SELECT * FROM DatoFatturaPATestataDoc "
    sSQL = sSQL & "WHERE IDDatoFatturaPATestataDoc=0"
    
    Set rsNew = New ADODB.Recordset
    rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic
    rsNew.AddNew
        rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
        rsNew!IDBloccoXML = IDBlocco
        rsNew!IDOggetto = oDoc.IDOggetto
        rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
        If NumeroRiferimentoRiga > 0 Then
            rsNew!RiferimentoNumeroLinea = NumeroRiferimentoRiga
        End If
        rsNew!IDDocumento = IDDocumentoRiga
        If IDBlocco <> 6 Then ''DIVERSO DA DDT
            If Len(DataOrdine) > 0 Then
                rsNew!Data = DataOrdine
            End If
            rsNew!NumItem = NumeroOrdine
            rsNew!IDDocumento = NumeroOrdine
        Else
            If Len(DataOrdine) > 0 Then
                rsNew!DataDDT = DataOrdine
            End If
            rsNew!NumeroDDT = NumeroOrdine
        End If
    rsNew.Update
    
    rsNew.Close
    Set rsNew = Nothing
End If
Exit Sub
ERR_SCRIVI_ORD_CLI_RIF_XML:
    MsgBox Err.Description, vbCritical, "SCRIVI_ORD_CLI_RIF_XML"

End Sub
Private Sub txtDataOrdineCliente_LostFocus()
On Error Resume Next
    If (Me.txtDataOrdineCliente.Text = fnNotNull(oDoc.Field("Doc_data_vs_ordine_di_rifer", , sTabellaTestata))) Then Exit Sub
    sbImpostaDatiDocumento
End Sub
Private Sub txtNumeroOrdineCliente_LostFocus()
On Error Resume Next
    If (Me.txtNumeroOrdineCliente.Text = fnNotNull(oDoc.Field("Doc_numero_vs_ordine_di_rifer", , sTabellaTestata))) Then Exit Sub
    sbImpostaDatiDocumento
End Sub
Private Sub SCRIVI_CAUSALI_DOC(IDOggetto As Long)
On Error GoTo ERR_SCRIVI_CAUSALI_DOC
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim rsNew As ADODB.Recordset
Dim rsDoc As DmtOleDbLib.adoResultset
Dim TIpo As Long
Dim Ordinamento As Long
Dim TestoVettoreSuccessivo As String
Dim TestoAgenziaTraporto As String

If oDoc.IDTipoOggetto = 8 Then Exit Sub

TIpo = 0

If (oDoc.Field("RV_PODocumentoCRM", , sTabellaTestata) = 1) Then
    TIpo = 1
Else
    If fnNotNullN(oDoc.Field("RV_POIDAnagraficaDestinazione", , sTabellaTestata)) > 0 Then
        TIpo = 2
    End If
End If

sSQL = "SELECT * FROM DatoFatturaPATestataDoc "
sSQL = sSQL & "WHERE IDDatoFatturaPATestataDoc=0"

Set rsNew = New ADODB.Recordset

rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

Ordinamento = ADD_NOTE_TIPO_OGGETTO(oDoc.IDTipoOggetto, oDoc.IDOggetto, TIpo, rsNew)


If Rip_InXMLRifLetteraIntento = 1 Then
    ADD_NOTA_LETTERA_INTENTO oDoc.IDTipoOggetto, oDoc.IDOggetto, rsNew, fnNotNullN(oDoc.Field("Link_nom_lettera_intento", , sTabellaTestata)), Ordinamento
End If

If Rip_InXMLRifNoteIva = 1 Then
    ADD_NOTE_IVA oDoc.IDTipoOggetto, oDoc.IDOggetto, rsNew, Ordinamento
End If

'ANNOTAZIONE 1 DOCUMENTO
If Rip_InXMLRifNota01Doc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni1", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni1", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ANNOTAZIONE 2 DOCUMENTO
If Rip_InXMLRifNota02Doc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni2", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni2", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ANNOTAZIONE 3 DOCUMENTO
If Rip_InXMLRifNota03Doc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni3", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POAnnotazioni3", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If
'ANNOTAZIONE STANDARD DEL DOCUMENTO
If Rip_InXMLRifNotaDoc = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("Doc_annotazioni_variazio", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("Doc_annotazioni_variazio", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'ISTRUZIONI DEL MITTENTE
If Rip_InXMLRifIstrMitt = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POIstruzioniMittente", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = Mid(Trim(fnNotNull(oDoc.Field("RV_POIstruzioniMittente", , sTabellaTestata))), 1, 200)
            rsNew!Ordinamento = Ordinamento
            Ordinamento = Ordinamento + 1
        rsNew.Update
    End If
End If

'VETTORE SUCCESSIVO
If Rip_InXMLRifVettSucc = 1 Then
    If fnNotNullN(oDoc.Field("RV_POIDTrasportatoreSuccessivo", , sTabellaTestata)) > 0 Then
        TestoVettoreSuccessivo = GET_VETTORE_SUCCESSIVO(fnNotNullN(oDoc.Field("RV_POIDTrasportatoreSuccessivo", , sTabellaTestata)))
        If Len(TestoVettoreSuccessivo) > 0 Then
            rsNew.AddNew
                rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                rsNew!IDBloccoXML = 8
                rsNew!IDOggetto = IDOggetto
                rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
                rsNew!Annotazioni = Mid(TestoVettoreSuccessivo, 1, 200)
                rsNew!Ordinamento = Ordinamento
                Ordinamento = Ordinamento + 1
            rsNew.Update
        End If
    End If
End If

'AGENZIA DI TRASPORTO
If Rip_InXMLRifAgenziaTrasp = 1 Then
    If fnNotNullN(oDoc.Field("RV_POIDAgenziaTrasportatore", , sTabellaTestata)) > 0 Then
        TestoAgenziaTraporto = GET_AGENZIA_TRASPORTO(fnNotNullN(oDoc.Field("RV_POIDAgenziaTrasportatore", , sTabellaTestata)))
        If Len(TestoAgenziaTraporto) > 0 Then
            rsNew.AddNew
                rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                rsNew!IDBloccoXML = 8
                rsNew!IDOggetto = IDOggetto
                rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
                rsNew!Annotazioni = Mid(TestoAgenziaTraporto, 1, 200)
                rsNew!Ordinamento = Ordinamento
                Ordinamento = Ordinamento + 1
            rsNew.Update
        End If
    End If
End If
'TARGA AUTOMEZZO
If Rip_InXMLRifTargaAutoMezzo = 1 Then
    If Len(Trim(fnNotNull(oDoc.Field("RV_POTargaAutomezzo", , sTabellaTestata)))) > 0 Then
        rsNew.AddNew
            rsNew!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsNew!IDBloccoXML = 8
            rsNew!IDOggetto = IDOggetto
            rsNew!IDTipoOggetto = oDoc.IDTipoOggetto
            rsNew!Annotazioni = "Targa automezzo: " & Mid(Trim(fnNotNull(oDoc.Field("RV_POTargaAutomezzo", , sTabellaTestata))), 1, 200)
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
Private Function ADD_NOTE_TIPO_OGGETTO(IDTipoOggetto As Long, IDOggetto As Long, TIpo As Long, rsAdd As ADODB.Recordset) As Long
On Error GoTo ERR_ADD_NOTE_TIPO_OGGETTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Ordinamento As Long

sSQL = "SELECT * FROM RV_PONoteDocumentiCoop "
sSQL = sSQL & " WHERE IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)

Ordinamento = 1

If Not rs.EOF Then
    Select Case TIpo
        Case 0
            If Len(Trim(fnNotNull(rs!Annotazioni1))) > 0 Then
                If fnNotNullN(rs!NonRiportareInXMLAnnotazione01) = 0 Then
                    rsAdd.AddNew
                        rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
                        rsAdd!IDBloccoXML = 8
                        rsAdd!IDOggetto = IDOggetto
                        rsAdd!IDTipoOggetto = IDTipoOggetto
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni1)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni3)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni4)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni6)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni7)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni8)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni9)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni10)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni11)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni12)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni13)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni14)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni15)), 1, 200)
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
                        rsAdd!Annotazioni = Mid(Trim(fnNotNull(rs!Annotazioni5)), 1, 200)
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
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto

Set rsNota = Cn.OpenResultset(sSQL)

If Not rsNota.EOF Then
    Nota2 = fnNotNull(rsNota!Annotazioni2)
End If

rsNota.CloseResultset
Set rsNota = Nothing

sSQL = "SELECT IDLetteraIntento, IDTipoLetteraIntento, IDAzienda, Data, Numero, Anno, NumeroCliFor, AnnoCliFor, "
sSQL = sSQL & "IDAnagrafica_CF, IDTipoAnagrafica_CF, IDAzienda_CF, ProgressivoDichiarazione, ProtocolloDichiarazione, DataEmissione "
sSQL = sSQL & "FROM LetteraIntento "
sSQL = sSQL & "WHERE IDLetteraIntento=" & IDLetteraIntento

Set rs = Cn.OpenResultset(sSQL)

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
        rsAdd!Annotazioni = Mid(Trim(Testo), 1, 200)
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

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If Len(Trim(rs!Annotazioni)) > 0 Then
        rsAdd.AddNew
            rsAdd!IDDatoFatturaPATestataDoc = fnGetNewKey("DatoFatturaPATestataDoc", "IDDatoFatturaPATestataDoc")
            rsAdd!IDBloccoXML = 8
            rsAdd!IDOggetto = IDOggetto
            rsAdd!IDTipoOggetto = IDTipoOggetto
            rsAdd!Annotazioni = Mid(Trim(rs!Annotazioni), 1, 200)
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

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_VETTORE_SUCCESSIVO = "Vettore successivo: " & fnNotNull(rs!Vettore)
    If Len(Trim(fnNotNull(rs!PartitaIVA))) > 0 Then
        GET_VETTORE_SUCCESSIVO = GET_VETTORE_SUCCESSIVO + " - Partita I.V.A.: " & fnNotNull(rs!PartitaIVA)
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

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_AGENZIA_TRASPORTO = "Agenzia di trasporto: " & fnNotNull(rs!Anagrafica) + fnNotNull(rs!Nome)
    If Len(Trim(fnNotNull(rs!PartitaIVA))) > 0 Then
        GET_AGENZIA_TRASPORTO = GET_AGENZIA_TRASPORTO + " - Partita I.V.A.: " & fnNotNull(rs!PartitaIVA)
    End If
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_AGENZIA_TRASPORTO:
    MsgBox Err.Description, vbCritical, "GET_AGENZIA_TRASPORTO"
End Function
Private Sub RICALCOLA_DESCRIZIONI_CAUSALI()
On Error GoTo ERR_RICALCOLA_DESCRIZIONI_CAUSALI
Dim Testo As String
Dim sSQL As String

If oDoc.IDOggetto = 0 Then Exit Sub

Testo = "ATTENZIONE!!!" & vbCrLf
Testo = Testo & "Con questo comando verranno eliminate tutte le descrizioni delle causali in XML e rigenerate." & vbCrLf
Testo = Testo & "Continuare con questo comando?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo dati") = vbNo Then Exit Sub

sSQL = "DELETE FROM DatoFatturaPATestataDoc "
sSQL = sSQL & " WHERE IDBloccoXML=8 "
sSQL = sSQL & " AND IDOggetto=" & oDoc.IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & oDoc.IDTipoOggetto
Cn.Execute sSQL


SCRIVI_CAUSALI_DOC oDoc.IDOggetto

MsgBox "Operazione completata", vbInformation, "Ricalcola descrizioni causali XML"
Exit Sub
ERR_RICALCOLA_DESCRIZIONI_CAUSALI:
    MsgBox Err.Description, vbCritical, "RICALCOLA_DESCRIZIONI_CAUSALI"
End Sub
Private Function GET_SEZ_PER_CLIENTE(IDTipoOggetto As Long, IDCliente As Long) As Long
On Error GoTo ERR_GET_SEZ_PER_CLIENTE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_SEZ_PER_CLIENTE = 0

sSQL = "SELECT IDRV_POConfigurazioneCliente, IDSezionalePerDDT, IDSezionalePerFA, IDSezionalePerSNF, "
sSQL = sSQL & "IDSezionalePerNotaDebito, IDSezionalePerNotaCredito "
sSQL = sSQL & " FROM RV_POConfigurazioneCliente "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND IDAnagrafica=" & IDCliente

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Select Case IDTipoOggetto
        Case 2
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerDDT)
        Case 114
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerFA)
        Case 8
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerSNF)
        Case 107
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerNotaDebito)
        Case 11
            GET_SEZ_PER_CLIENTE = fnNotNullN(rs!IDSezionalePerNotaCredito)
    End Select
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_SEZ_PER_CLIENTE:
    MsgBox Err.Description, vbCritical, "GET_SEZ_PER_CLIENTE"

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

sSQL = "SELECT RiportaInXMLRifLetteraIntento, RiportaInXMLRifNoteIva, RiportaInXMLRifNota01Doc, "
sSQL = sSQL & "RiportaInXMLRifNota02Doc, RiportaInXMLRifNota03Doc, RiportaInXMLRifNotaDoc, "
sSQL = sSQL & "RiportaInXMLRifIstrMitt, RiportaInXMLRifVettSucc, RiportaInXMLRifAgenziaTrasp, "
sSQL = sSQL & "RiportaInXMLRifTargaAutoMezzo"
sSQL = sSQL & " FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

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
End If
rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_RECUPERA_CONFIG_CAUS_XML:
    MsgBox Err.Description, vbCritical, "RECUPERA_CONFIG_CAUS_XML"
End Sub
