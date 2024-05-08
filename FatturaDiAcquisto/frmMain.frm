VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{7A1D73E4-F461-11D0-8F01-004033A00AF2}#1.0#0"; "DmtWheel.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{9385BB2E-6637-11D1-850D-002018802E11}#3.1#0"; "Dmtsplit.ocx"
Object = "{E1215E52-40E1-11D3-AF44-00105A2FBE61}#5.1#0"; "DMTLblLinkCtl.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{FCA49525-5F72-11D2-B9EB-00201880103B}#18.1#0"; "DMTPrinterDialog.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   11565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18285
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   11565
   ScaleWidth      =   18285
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   11220
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   18285
      _LayoutVersion  =   2
      _ExtentX        =   32253
      _ExtentY        =   19791
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
      Begin DMTPrinterDialog.DMTDialog DmtPrnDlg 
         Left            =   360
         Top             =   5760
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
      End
      Begin VB.PictureBox PicForm 
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
         Height          =   11145
         Left            =   0
         ScaleHeight     =   11115
         ScaleWidth      =   18165
         TabIndex        =   102
         Top             =   0
         Width           =   18195
         Begin VB.PictureBox PicForm2 
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
            Height          =   10935
            Left            =   120
            ScaleHeight     =   10905
            ScaleWidth      =   17865
            TabIndex        =   103
            Top             =   120
            Width           =   17895
            Begin TabDlg.SSTab SSTab1 
               Height          =   7455
               Left            =   120
               TabIndex        =   116
               Top             =   3400
               Width           =   17655
               _ExtentX        =   31141
               _ExtentY        =   13150
               _Version        =   393216
               TabHeight       =   520
               TabCaption(0)   =   "MERCE CONFERITA/ACQUISTATA"
               TabPicture(0)   =   "frmMain.frx":479EA
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "fraRighe"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "GESTIONE ALTRI IMBALLI"
               TabPicture(1)   =   "frmMain.frx":47A06
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label2(26)"
               Tab(1).Control(1)=   "Label2(34)"
               Tab(1).Control(2)=   "Label2(35)"
               Tab(1).Control(3)=   "Label2(38)"
               Tab(1).Control(4)=   "Label2(41)"
               Tab(1).Control(5)=   "Label2(45)"
               Tab(1).Control(6)=   "txtImpUniAltriImb"
               Tab(1).Control(7)=   "txtIDLottoImballoGest"
               Tab(1).Control(8)=   "txtGiacLottoImballo"
               Tab(1).Control(9)=   "txtDispLottoImballo"
               Tab(1).Control(10)=   "cboUMImbGest"
               Tab(1).Control(11)=   "cboTipoProcessoImballo"
               Tab(1).Control(12)=   "CDImballoGestione"
               Tab(1).Control(13)=   "GrigliaImballi"
               Tab(1).Control(14)=   "cmdEliminaImballo"
               Tab(1).Control(14).Enabled=   0   'False
               Tab(1).Control(15)=   "cmdSalvaImballo"
               Tab(1).Control(16)=   "cmdNuovoImballo"
               Tab(1).Control(17)=   "txtQuantitaImballo"
               Tab(1).Control(18)=   "cmdRiepilogoLottoImballo"
               Tab(1).Control(19)=   "cmdLottoImballo"
               Tab(1).Control(20)=   "txtLottoImballoGest"
               Tab(1).Control(21)=   "chkConfermaDaUtente"
               Tab(1).Control(22)=   "chkTracciaImballoGest"
               Tab(1).Control(23)=   "txtLottoImballoRifEsterno"
               Tab(1).ControlCount=   24
               TabCaption(2)   =   "GESTIONE ALTRI ADDEBITI"
               TabPicture(2)   =   "frmMain.frx":47A22
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Label2(27)"
               Tab(2).Control(1)=   "Label2(28)"
               Tab(2).Control(2)=   "Label6(2)"
               Tab(2).Control(3)=   "Label2(29)"
               Tab(2).Control(4)=   "Label2(30)"
               Tab(2).Control(5)=   "Label2(31)"
               Tab(2).Control(6)=   "Label2(32)"
               Tab(2).Control(7)=   "Label2(33)"
               Tab(2).Control(8)=   "txtQtaAddebiti"
               Tab(2).Control(9)=   "txtImpUniAddebiti"
               Tab(2).Control(10)=   "txtImponibileAddebiti"
               Tab(2).Control(11)=   "txtImpostaAddebiti"
               Tab(2).Control(12)=   "txtTotaleRigaAddebiti"
               Tab(2).Control(13)=   "cboTipoTrattenutaAggiuntiva"
               Tab(2).Control(14)=   "txtAliquotaIvaAddebiti"
               Tab(2).Control(15)=   "cboIvaAddebiti"
               Tab(2).Control(16)=   "CDArticoliAddebiti"
               Tab(2).Control(17)=   "GrigliaAddebiti"
               Tab(2).Control(18)=   "cmdEliminaAddebiti"
               Tab(2).Control(18).Enabled=   0   'False
               Tab(2).Control(19)=   "cmdSalvaAddebiti"
               Tab(2).Control(20)=   "cmdNuovoAddebiti"
               Tab(2).ControlCount=   21
               Begin VB.TextBox txtLottoImballoRifEsterno 
                  Height          =   315
                  Left            =   -69000
                  TabIndex        =   79
                  Top             =   1560
                  Width           =   5415
               End
               Begin VB.CheckBox chkTracciaImballoGest 
                  Caption         =   "Traccia"
                  Height          =   255
                  Left            =   -74640
                  TabIndex        =   216
                  Top             =   2520
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.CheckBox chkConfermaDaUtente 
                  Caption         =   "Conferma da utente"
                  Height          =   255
                  Left            =   -70200
                  TabIndex        =   213
                  Top             =   2520
                  Visible         =   0   'False
                  Width           =   2535
               End
               Begin VB.TextBox txtLottoImballoGest 
                  Height          =   315
                  Left            =   -74880
                  Locked          =   -1  'True
                  TabIndex        =   211
                  Top             =   1560
                  Width           =   5415
               End
               Begin VB.CommandButton cmdLottoImballo 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   -65400
                  Picture         =   "frmMain.frx":47A3E
                  Style           =   1  'Graphical
                  TabIndex        =   210
                  ToolTipText     =   "Seleziona lotti"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.CommandButton cmdRiepilogoLottoImballo 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   -69480
                  Picture         =   "frmMain.frx":47FC8
                  Style           =   1  'Graphical
                  TabIndex        =   209
                  ToolTipText     =   "Riepilogo lotto imballo"
                  Top             =   1560
                  Width           =   375
               End
               Begin DMTEDITNUMLib.dmtNumber txtQuantitaImballo 
                  Height          =   315
                  Left            =   -66360
                  TabIndex        =   77
                  Top             =   960
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
               Begin VB.CommandButton cmdNuovoAddebiti 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -59040
                  TabIndex        =   94
                  Top             =   2520
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaAddebiti 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -59040
                  TabIndex        =   93
                  Top             =   3480
                  Width           =   1455
               End
               Begin VB.CommandButton cmdEliminaAddebiti 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -59040
                  TabIndex        =   95
                  TabStop         =   0   'False
                  Top             =   4440
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNuovoImballo 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -59040
                  TabIndex        =   81
                  Top             =   2400
                  Width           =   1455
               End
               Begin VB.CommandButton cmdSalvaImballo 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -59040
                  TabIndex        =   80
                  Top             =   3360
                  Width           =   1455
               End
               Begin VB.CommandButton cmdEliminaImballo 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -59040
                  TabIndex        =   82
                  TabStop         =   0   'False
                  Top             =   4320
                  Width           =   1455
               End
               Begin VB.Frame fraRighe 
                  Height          =   6855
                  Left            =   120
                  TabIndex        =   117
                  Top             =   480
                  Width           =   17415
                  Begin VB.CheckBox chkLottoImbEsclusivo 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Esclusivo"
                     Height          =   315
                     Left            =   15960
                     TabIndex        =   260
                     TabStop         =   0   'False
                     Top             =   1680
                     Width           =   1335
                  End
                  Begin VB.CommandButton cmdGestioneQualita 
                     Caption         =   "Qualità"
                     Height          =   375
                     Left            =   10200
                     TabIndex        =   256
                     TabStop         =   0   'False
                     Top             =   6360
                     Width           =   2415
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "Elenco qualità"
                     Height          =   375
                     Left            =   12720
                     TabIndex        =   255
                     TabStop         =   0   'False
                     Top             =   6360
                     Width           =   2415
                  End
                  Begin VB.Frame Frame4 
                     Caption         =   "Provenienza"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   975
                     Left            =   9600
                     TabIndex        =   133
                     Top             =   3120
                     Width           =   2655
                     Begin VB.CommandButton cmdApriFrameProvenienza 
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Left            =   2160
                        Picture         =   "frmMain.frx":48552
                        Style           =   1  'Graphical
                        TabIndex        =   134
                        TabStop         =   0   'False
                        ToolTipText     =   "Ricalcola dati cliente"
                        Top             =   0
                        Width           =   375
                     End
                     Begin DMTDataCmb.DMTCombo cboRegione 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   64
                        TabStop         =   0   'False
                        Top             =   480
                        Width           =   2415
                        _ExtentX        =   4260
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
                     Begin DMTDataCmb.DMTCombo cboNazione 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   65
                        TabStop         =   0   'False
                        Top             =   1200
                        Width           =   2415
                        _ExtentX        =   4260
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
                     Begin DMTDataCmb.DMTCombo cboComune 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   66
                        TabStop         =   0   'False
                        Top             =   1800
                        Width           =   2415
                        _ExtentX        =   4260
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
                     Begin DMTDataCmb.DMTCombo cboProvincia 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   135
                        TabStop         =   0   'False
                        Top             =   2400
                        Width           =   2415
                        _ExtentX        =   4260
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
                     Begin VB.Label Label4 
                        Caption         =   "Regione"
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   139
                        Top             =   240
                        Width           =   2415
                     End
                     Begin VB.Label Label4 
                        Caption         =   "Nazione"
                        Height          =   255
                        Index           =   1
                        Left            =   120
                        TabIndex        =   138
                        Top             =   960
                        Width           =   2055
                     End
                     Begin VB.Label Label4 
                        Caption         =   "Comune"
                        Height          =   255
                        Index           =   2
                        Left            =   120
                        TabIndex        =   137
                        Top             =   1560
                        Width           =   1095
                     End
                     Begin VB.Label Label4 
                        Caption         =   "Prov."
                        Height          =   255
                        Index           =   3
                        Left            =   120
                        TabIndex        =   136
                        Top             =   2160
                        Width           =   735
                     End
                  End
                  Begin VB.Frame Frame3 
                     Caption         =   "RIEPILOGO"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   975
                     Left            =   120
                     TabIndex        =   118
                     Top             =   3120
                     Width           =   9375
                     Begin VB.TextBox txtNomeMacchina 
                        Enabled         =   0   'False
                        Height          =   285
                        Left            =   120
                        TabIndex        =   123
                        Top             =   1440
                        Width           =   2655
                     End
                     Begin VB.TextBox txtUtenteMacchina 
                        Enabled         =   0   'False
                        Height          =   285
                        Left            =   2880
                        TabIndex        =   122
                        Top             =   1440
                        Width           =   2535
                     End
                     Begin VB.TextBox txtCodiceUtente 
                        Enabled         =   0   'False
                        Height          =   315
                        Left            =   7680
                        TabIndex        =   120
                        Top             =   1440
                        Width           =   1335
                     End
                     Begin VB.CommandButton cmdApriFrameRiepilogo 
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Left            =   8880
                        Picture         =   "frmMain.frx":48ADC
                        Style           =   1  'Graphical
                        TabIndex        =   119
                        TabStop         =   0   'False
                        ToolTipText     =   "Ricalcola dati cliente"
                        Top             =   0
                        Width           =   375
                     End
                     Begin DMTDataCmb.DMTCombo cboUtente 
                        Height          =   315
                        Left            =   5520
                        TabIndex        =   121
                        Top             =   1440
                        Width           =   2055
                        _ExtentX        =   3625
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
                     Begin DMTEDITNUMLib.dmtNumber txtQtaQuadrata 
                        Height          =   285
                        Left            =   3720
                        TabIndex        =   60
                        Top             =   480
                        Width           =   1695
                        _Version        =   65536
                        _ExtentX        =   2990
                        _ExtentY        =   503
                        _StockProps     =   253
                        Text            =   "0"
                        ForeColor       =   65535
                        BackColor       =   12582912
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
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
                     Begin DMTEDITNUMLib.dmtNumber txtQtaVenduta 
                        Height          =   285
                        Left            =   120
                        TabIndex        =   58
                        Top             =   480
                        Width           =   1695
                        _Version        =   65536
                        _ExtentX        =   2990
                        _ExtentY        =   503
                        _StockProps     =   253
                        Text            =   "0"
                        ForeColor       =   65535
                        BackColor       =   12582912
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
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
                     Begin DMTEDITNUMLib.dmtNumber txtQtaDifferenza 
                        Height          =   285
                        Left            =   5520
                        TabIndex        =   61
                        Top             =   480
                        Width           =   1695
                        _Version        =   65536
                        _ExtentX        =   2990
                        _ExtentY        =   503
                        _StockProps     =   253
                        Text            =   "0"
                        ForeColor       =   65535
                        BackColor       =   12632319
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
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
                     Begin DMTEDITNUMLib.dmtNumber txtQtaAssegnata 
                        Height          =   285
                        Left            =   1920
                        TabIndex        =   59
                        Top             =   480
                        Width           =   1695
                        _Version        =   65536
                        _ExtentX        =   2990
                        _ExtentY        =   503
                        _StockProps     =   253
                        Text            =   "0"
                        ForeColor       =   65535
                        BackColor       =   12582912
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
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
                     Begin DMTEDITNUMLib.dmtNumber txtQtaDifferenzaLavorazione 
                        Height          =   285
                        Left            =   7320
                        TabIndex        =   62
                        Top             =   480
                        Width           =   1695
                        _Version        =   65536
                        _ExtentX        =   2990
                        _ExtentY        =   503
                        _StockProps     =   253
                        Text            =   "0"
                        ForeColor       =   65535
                        BackColor       =   12632319
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Verdana"
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
                     Begin VB.Label Label2 
                        Caption         =   "Diff. di lavorazione"
                        Height          =   255
                        Index           =   14
                        Left            =   7320
                        TabIndex        =   132
                        Top             =   280
                        Width           =   1695
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà Assegnata"
                        Height          =   255
                        Index           =   13
                        Left            =   1920
                        TabIndex        =   131
                        Top             =   280
                        Width           =   1695
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà quadratura"
                        Height          =   255
                        Index           =   6
                        Left            =   3720
                        TabIndex        =   130
                        Top             =   280
                        Width           =   1575
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Q.tà Vend."
                        Height          =   255
                        Index           =   7
                        Left            =   120
                        TabIndex        =   129
                        Top             =   280
                        Width           =   1335
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Diff. di vendita"
                        Height          =   255
                        Index           =   11
                        Left            =   5520
                        TabIndex        =   128
                        Top             =   280
                        Width           =   1335
                     End
                     Begin VB.Label Label5 
                        Caption         =   "Nome macchina"
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   127
                        Top             =   1200
                        Width           =   2655
                     End
                     Begin VB.Label Label5 
                        Caption         =   "Utente macchina"
                        Height          =   255
                        Index           =   1
                        Left            =   2880
                        TabIndex        =   126
                        Top             =   1200
                        Width           =   2655
                     End
                     Begin VB.Label Label5 
                        Caption         =   "Nome utente"
                        Height          =   255
                        Index           =   3
                        Left            =   5520
                        TabIndex        =   125
                        Top             =   1200
                        Width           =   2055
                     End
                     Begin VB.Label Label5 
                        Caption         =   "Codice utente"
                        Height          =   255
                        Index           =   4
                        Left            =   7680
                        TabIndex        =   124
                        Top             =   1200
                        Width           =   1335
                     End
                  End
                  Begin VB.TextBox txtNoteAgg 
                     Height          =   375
                     Left            =   12360
                     TabIndex        =   57
                     Top             =   3720
                     Width           =   4935
                  End
                  Begin VB.CommandButton cmdPesaturaAut 
                     Caption         =   "Pesatura (F8)"
                     Height          =   375
                     Left            =   7680
                     TabIndex        =   224
                     TabStop         =   0   'False
                     Top             =   6360
                     Width           =   2415
                  End
                  Begin VB.CommandButton cmdPesatura 
                     Height          =   315
                     Left            =   120
                     Picture         =   "frmMain.frx":49066
                     Style           =   1  'Graphical
                     TabIndex        =   223
                     ToolTipText     =   "Pesatura pedane conferite"
                     Top             =   2235
                     Width           =   380
                  End
                  Begin VB.CheckBox chkTracciaImballo 
                     Caption         =   "Traccia"
                     Height          =   255
                     Left            =   9120
                     TabIndex        =   207
                     Top             =   6360
                     Visible         =   0   'False
                     Width           =   975
                  End
                  Begin VB.TextBox txtLottoImballo 
                     Height          =   315
                     Left            =   10320
                     Locked          =   -1  'True
                     TabIndex        =   205
                     Top             =   1725
                     Width           =   5175
                  End
                  Begin VB.CommandButton cmdRiepLottoImballoConf 
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   15480
                     Picture         =   "frmMain.frx":495F0
                     Style           =   1  'Graphical
                     TabIndex        =   204
                     ToolTipText     =   "Riepilogo lotto imballo"
                     Top             =   1725
                     Width           =   375
                  End
                  Begin VB.CheckBox chkPrezzoMedio 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Prezzo medio"
                     Height          =   315
                     Left            =   15600
                     TabIndex        =   56
                     Top             =   3120
                     Width           =   1695
                  End
                  Begin DMTLblLinkCtl.LabelLink LinkLavorazione 
                     Height          =   255
                     Left            =   7920
                     TabIndex        =   184
                     Top             =   6360
                     Visible         =   0   'False
                     Width           =   975
                     _ExtentX        =   1720
                     _ExtentY        =   450
                     Caption         =   "Lavorazione"
                     Name            =   "LabelLink"
                  End
                  Begin VB.CommandButton cmdLavorazioneAutomatica 
                     Caption         =   "Lavorazione automatica"
                     Height          =   375
                     Left            =   120
                     TabIndex        =   71
                     TabStop         =   0   'False
                     Top             =   6360
                     Width           =   2415
                  End
                  Begin VB.CommandButton cmdCampionatura 
                     Caption         =   "Campionatura (F5)"
                     Height          =   375
                     Left            =   2640
                     TabIndex        =   72
                     TabStop         =   0   'False
                     Top             =   6360
                     Width           =   2415
                  End
                  Begin VB.TextBox txtLottoDiEntrata 
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   10440
                     Locked          =   -1  'True
                     TabIndex        =   46
                     TabStop         =   0   'False
                     Top             =   2235
                     Width           =   5655
                  End
                  Begin VB.CommandButton cmdRiepilogo 
                     Caption         =   "Riepilogo (F4)"
                     Height          =   375
                     Left            =   5160
                     TabIndex        =   73
                     TabStop         =   0   'False
                     Top             =   6360
                     Width           =   2415
                  End
                  Begin VB.CheckBox chkLottoChiuso 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Lotto chiuso"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   255
                     Left            =   15600
                     TabIndex        =   63
                     TabStop         =   0   'False
                     Top             =   2760
                     Width           =   1695
                  End
                  Begin VB.CommandButton cmdNuovo 
                     Caption         =   "Nuovo"
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
                     Left            =   15840
                     TabIndex        =   68
                     Top             =   4320
                     Width           =   1455
                  End
                  Begin VB.CommandButton cmdSalva 
                     Caption         =   "Salva"
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
                     Left            =   15840
                     TabIndex        =   67
                     Top             =   5040
                     Width           =   1455
                  End
                  Begin VB.CommandButton cmdElimina 
                     Caption         =   "Elimina"
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
                     Left            =   15840
                     TabIndex        =   69
                     TabStop         =   0   'False
                     Top             =   5760
                     Width           =   1455
                  End
                  Begin VB.TextBox txtImballo 
                     Height          =   315
                     Left            =   8880
                     TabIndex        =   31
                     Top             =   1200
                     Width           =   4335
                  End
                  Begin VB.TextBox TxtArticolo 
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   28
                     Top             =   1200
                     Width           =   3615
                  End
                  Begin VB.Frame FraLottoCampagna 
                     Height          =   855
                     Left            =   120
                     TabIndex        =   140
                     Top             =   120
                     Width           =   17175
                     Begin VB.TextBox txtVarietà 
                        Enabled         =   0   'False
                        Height          =   315
                        Left            =   8760
                        TabIndex        =   24
                        TabStop         =   0   'False
                        Top             =   360
                        Width           =   2415
                     End
                     Begin VB.TextBox txtFamiglia 
                        Enabled         =   0   'False
                        Height          =   315
                        Left            =   11280
                        TabIndex        =   25
                        TabStop         =   0   'False
                        Top             =   360
                        Width           =   2055
                     End
                     Begin VB.TextBox txtTipoProduzione 
                        Enabled         =   0   'False
                        Height          =   315
                        Left            =   13440
                        TabIndex        =   26
                        TabStop         =   0   'False
                        Top             =   360
                        Width           =   3615
                     End
                     Begin VB.TextBox txtLottoDiConferimento 
                        Height          =   315
                        Left            =   720
                        TabIndex        =   19
                        Top             =   360
                        Width           =   4125
                     End
                     Begin VB.CommandButton cmdTracciabilitaLottoCampagna 
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Left            =   4800
                        Picture         =   "frmMain.frx":49B7A
                        Style           =   1  'Graphical
                        TabIndex        =   20
                        ToolTipText     =   "Tracciabilità"
                        Top             =   360
                        Width           =   495
                     End
                     Begin VB.CommandButton cmdSelezionaLottoCampagna 
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Left            =   120
                        Picture         =   "frmMain.frx":4A104
                        Style           =   1  'Graphical
                        TabIndex        =   18
                        ToolTipText     =   "Trova lotto di campagna"
                        Top             =   360
                        Width           =   615
                     End
                     Begin VB.CommandButton cmdArticoliDerivati 
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   315
                        Left            =   5280
                        Picture         =   "frmMain.frx":4A68E
                        Style           =   1  'Graphical
                        TabIndex        =   21
                        TabStop         =   0   'False
                        ToolTipText     =   "Articoli derivati"
                        Top             =   360
                        Width           =   495
                     End
                     Begin DMTDATETIMELib.dmtDate txtDataSbloccoLotto 
                        Height          =   315
                        Left            =   7200
                        TabIndex        =   23
                        TabStop         =   0   'False
                        Top             =   360
                        Width           =   1455
                        _Version        =   65536
                        _ExtentX        =   2566
                        _ExtentY        =   556
                        _StockProps     =   253
                        BackColor       =   16777215
                        Enabled         =   0   'False
                        Appearance      =   1
                     End
                     Begin DMTEDITNUMLib.dmtNumber txtIDLottoCampagna 
                        Height          =   315
                        Left            =   5880
                        TabIndex        =   22
                        TabStop         =   0   'False
                        Top             =   360
                        Width           =   1215
                        _Version        =   65536
                        _ExtentX        =   2143
                        _ExtentY        =   556
                        _StockProps     =   253
                        BackColor       =   16777215
                        Enabled         =   0   'False
                        Appearance      =   1
                     End
                     Begin VB.Label Label6 
                        Caption         =   "Varietà"
                        Height          =   255
                        Index           =   0
                        Left            =   8760
                        TabIndex        =   146
                        Top             =   120
                        Width           =   1695
                     End
                     Begin VB.Label Label7 
                        Caption         =   "Famiglia"
                        Height          =   255
                        Index           =   0
                        Left            =   11280
                        TabIndex        =   145
                        Top             =   120
                        Width           =   1815
                     End
                     Begin VB.Label Label6 
                        Caption         =   "Data di sblocco"
                        Height          =   255
                        Index           =   1
                        Left            =   7200
                        TabIndex        =   144
                        Top             =   120
                        Width           =   1455
                     End
                     Begin VB.Label Label7 
                        Caption         =   "Tipo produzione"
                        Height          =   255
                        Index           =   1
                        Left            =   13440
                        TabIndex        =   143
                        Top             =   120
                        Width           =   1695
                     End
                     Begin VB.Label Label7 
                        Caption         =   "Identificativo"
                        Height          =   255
                        Index           =   2
                        Left            =   5880
                        TabIndex        =   142
                        Top             =   120
                        Width           =   1215
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Lotto di produzione"
                        ForeColor       =   &H00000000&
                        Height          =   255
                        Index           =   8
                        Left            =   720
                        TabIndex        =   141
                        Top             =   120
                        Width           =   2775
                     End
                  End
                  Begin VB.TextBox txtOraArrivo 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   16200
                     Locked          =   -1  'True
                     TabIndex        =   47
                     Top             =   2235
                     Width           =   1095
                  End
                  Begin DMTLblLinkCtl.LabelLink LabelLink1 
                     Height          =   255
                     Left            =   12600
                     TabIndex        =   147
                     TabStop         =   0   'False
                     Top             =   2280
                     Visible         =   0   'False
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   450
                     Caption         =   "Gestione scarti "
                     Name            =   "LabelLink"
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTaraUnitaria 
                     Height          =   315
                     Left            =   15120
                     TabIndex        =   33
                     TabStop         =   0   'False
                     Top             =   1200
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
                  Begin DMTDataCmb.DMTCombo cboUM 
                     Height          =   315
                     Left            =   5640
                     TabIndex        =   29
                     TabStop         =   0   'False
                     Top             =   1200
                     Width           =   1455
                     _ExtentX        =   2566
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
                  Begin DmtCodDescCtl.DmtCodDesc CDCodiceImballo 
                     Height          =   615
                     Left            =   7200
                     TabIndex        =   30
                     Top             =   960
                     Width           =   1695
                     _ExtentX        =   2990
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4AC18
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4AC6F
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4ACC6
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
                  Begin DMTEDITNUMLib.dmtNumber txtQta_UM 
                     Height          =   315
                     Left            =   7560
                     TabIndex        =   45
                     TabStop         =   0   'False
                     Top             =   2235
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
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
                     Left            =   6120
                     TabIndex        =   44
                     Top             =   2235
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtTara 
                     Height          =   315
                     Left            =   3120
                     TabIndex        =   42
                     Top             =   2235
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPesoNetto 
                     Height          =   315
                     Left            =   4680
                     TabIndex        =   43
                     Top             =   2235
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtPesoLordo 
                     Height          =   315
                     Left            =   1680
                     TabIndex        =   41
                     Top             =   2235
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtColli 
                     Height          =   315
                     Left            =   480
                     TabIndex        =   40
                     Top             =   2235
                     Width           =   1095
                     _Version        =   65536
                     _ExtentX        =   1931
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDArticolo 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   27
                     Top             =   960
                     Width           =   1815
                     _ExtentX        =   3201
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4AD20
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4AD6F
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4ADC6
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
                  Begin DmtGridCtl.DmtGrid GrigliaCorpo 
                     Height          =   2055
                     Left            =   120
                     TabIndex        =   70
                     TabStop         =   0   'False
                     Top             =   4200
                     Width           =   15615
                     _ExtentX        =   27543
                     _ExtentY        =   3625
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
                  Begin DMTEDITNUMLib.dmtNumber txtNumeroLottoEntrata 
                     Height          =   255
                     Left            =   12720
                     TabIndex        =   148
                     Top             =   1800
                     Visible         =   0   'False
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   450
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
                  Begin DMTDataCmb.DMTCombo cboUMImballo 
                     Height          =   315
                     Left            =   13320
                     TabIndex        =   32
                     TabStop         =   0   'False
                     Top             =   1200
                     Width           =   1695
                     _ExtentX        =   2990
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
                  Begin DmtCodDescCtl.DmtCodDesc CDPedana 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   35
                     Top             =   1485
                     Width           =   5055
                     _ExtentX        =   8916
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4AE20
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4AE76
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4AECD
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
                  Begin DMTEDITNUMLib.dmtNumber txtTaraPedana 
                     Height          =   315
                     Left            =   7320
                     TabIndex        =   38
                     Top             =   1725
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
                  Begin DMTDataCmb.DMTCombo cboUMPedana 
                     Height          =   315
                     Left            =   5160
                     TabIndex        =   36
                     TabStop         =   0   'False
                     Top             =   1725
                     Width           =   1095
                     _ExtentX        =   1931
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
                  Begin DMTEDITNUMLib.dmtNumber txtQuantitaPedana 
                     Height          =   315
                     Left            =   6360
                     TabIndex        =   37
                     Top             =   1725
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
                  Begin DMTEDITNUMLib.dmtCurrency txtTotaleRiga 
                     Height          =   315
                     Left            =   7920
                     TabIndex        =   53
                     TabStop         =   0   'False
                     Top             =   2760
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtImpostaRiga 
                     Height          =   315
                     Left            =   6360
                     TabIndex        =   52
                     TabStop         =   0   'False
                     Top             =   2760
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtImponibileRiga 
                     Height          =   315
                     Left            =   4800
                     TabIndex        =   51
                     TabStop         =   0   'False
                     Top             =   2760
                     Width           =   1455
                     _Version        =   65536
                     _ExtentX        =   2566
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTDataCmb.DMTCombo cboIVA 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   48
                     Top             =   2760
                     Width           =   2175
                     _ExtentX        =   3836
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
                  Begin DMTEDITNUMLib.dmtCurrency txtImportoUnitario 
                     Height          =   315
                     Left            =   3120
                     TabIndex        =   50
                     Top             =   2745
                     Width           =   1575
                     _Version        =   65536
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   " 0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     CurrencyDecimalPlaces=   5
                     CurrencySymbol  =   ""
                     AllowEmpty      =   0   'False
                     DecFinalZeros   =   -1  'True
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtAliquotaIVA 
                     Height          =   315
                     Left            =   2280
                     TabIndex        =   49
                     Top             =   2760
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboTipoLavorazione 
                     Height          =   315
                     Left            =   9480
                     TabIndex        =   54
                     Top             =   2745
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
                  Begin DMTDataCmb.DMTCombo cboLiquidato 
                     Height          =   315
                     Left            =   12840
                     TabIndex        =   55
                     Top             =   2745
                     Width           =   2655
                     _ExtentX        =   4683
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
                  Begin DMTEDITNUMLib.dmtNumber txtIDLottoImballo 
                     Height          =   315
                     Left            =   9120
                     TabIndex        =   208
                     Top             =   6360
                     Width           =   735
                     _Version        =   65536
                     _ExtentX        =   1296
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtCurrency txtImpUniImb 
                     Height          =   315
                     Left            =   16200
                     TabIndex        =   34
                     Top             =   1200
                     Width           =   1095
                     _Version        =   65536
                     _ExtentX        =   1931
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
                  Begin DMTEDITNUMLib.dmtNumber txtTaraAutomezzo 
                     Height          =   315
                     Left            =   8640
                     TabIndex        =   39
                     Top             =   1725
                     Width           =   1575
                     _Version        =   65536
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQtaPresunta 
                     Height          =   315
                     Left            =   9000
                     TabIndex        =   252
                     TabStop         =   0   'False
                     Top             =   2235
                     Width           =   1335
                     _Version        =   65536
                     _ExtentX        =   2355
                     _ExtentY        =   556
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Q.tà presunta"
                     Height          =   270
                     Index           =   42
                     Left            =   9000
                     TabIndex        =   253
                     Top             =   2040
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara automezzo"
                     Height          =   270
                     Index           =   43
                     Left            =   8640
                     TabIndex        =   251
                     Top             =   1515
                     Width           =   1575
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Imp. imb."
                     Height          =   255
                     Index           =   40
                     Left            =   16200
                     TabIndex        =   227
                     ToolTipText     =   "Importo unitario imballo"
                     Top             =   960
                     Width           =   1095
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Annotazioni aggiuntive"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00C00000&
                     Height          =   255
                     Index           =   1
                     Left            =   12360
                     TabIndex        =   226
                     Top             =   3480
                     Width           =   3615
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Lotto imballo"
                     Height          =   270
                     Index           =   39
                     Left            =   10320
                     TabIndex        =   206
                     Top             =   1515
                     Width           =   3375
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tipo lavorazione"
                     Height          =   255
                     Index           =   36
                     Left            =   9480
                     TabIndex        =   190
                     Top             =   2550
                     Width           =   1935
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Stato liquidazione"
                     Height          =   255
                     Index           =   37
                     Left            =   12840
                     TabIndex        =   189
                     Top             =   2550
                     Width           =   1695
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Lotto di entrata"
                     Height          =   270
                     Index           =   12
                     Left            =   10440
                     TabIndex        =   170
                     Top             =   2040
                     Width           =   3375
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara Uni."
                     Height          =   255
                     Index           =   10
                     Left            =   15120
                     TabIndex        =   169
                     Top             =   975
                     Width           =   855
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Unità di misura"
                     Height          =   270
                     Index           =   9
                     Left            =   5640
                     TabIndex        =   168
                     Top             =   980
                     Width           =   1335
                  End
                  Begin VB.Label lblImballo 
                     Caption         =   "Imballo"
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
                     Left            =   8880
                     TabIndex        =   167
                     Top             =   980
                     Width           =   2415
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Q.tà Mov."
                     Height          =   270
                     Index           =   5
                     Left            =   7560
                     TabIndex        =   166
                     Top             =   2040
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Peso netto"
                     Height          =   270
                     Index           =   4
                     Left            =   4680
                     TabIndex        =   165
                     Top             =   2040
                     Width           =   975
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara "
                     Height          =   270
                     Index           =   3
                     Left            =   3120
                     TabIndex        =   164
                     Top             =   2040
                     Width           =   855
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Pezzi"
                     Height          =   270
                     Index           =   2
                     Left            =   6120
                     TabIndex        =   163
                     Top             =   2040
                     Width           =   975
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Peso lordo"
                     Height          =   270
                     Index           =   1
                     Left            =   1680
                     TabIndex        =   162
                     Top             =   2040
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Colli"
                     Height          =   270
                     Index           =   0
                     Left            =   480
                     TabIndex        =   161
                     Top             =   2040
                     Width           =   1095
                  End
                  Begin VB.Label lblArticolo 
                     Caption         =   "Articolo"
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
                     TabIndex        =   160
                     Top             =   980
                     Width           =   3615
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Unità di misura"
                     Height          =   270
                     Index           =   15
                     Left            =   13320
                     TabIndex        =   159
                     Top             =   975
                     Width           =   1335
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Tara"
                     Height          =   270
                     Index           =   16
                     Left            =   7320
                     TabIndex        =   158
                     Top             =   1515
                     Width           =   1095
                  End
                  Begin VB.Label Label2 
                     Caption         =   "U.M."
                     Height          =   270
                     Index           =   17
                     Left            =   5160
                     TabIndex        =   157
                     Top             =   1520
                     Width           =   975
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Numero"
                     Height          =   270
                     Index           =   18
                     Left            =   6360
                     TabIndex        =   156
                     Top             =   1515
                     Width           =   855
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Totale riga"
                     Height          =   255
                     Index           =   19
                     Left            =   7920
                     TabIndex        =   155
                     Top             =   2550
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Imposta riga"
                     Height          =   255
                     Index           =   20
                     Left            =   6360
                     TabIndex        =   154
                     Top             =   2550
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Imponibile riga"
                     Height          =   255
                     Index           =   21
                     Left            =   4800
                     TabIndex        =   153
                     Top             =   2550
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Importo unitario"
                     Height          =   255
                     Index           =   22
                     Left            =   3120
                     TabIndex        =   152
                     Top             =   2550
                     Width           =   1215
                  End
                  Begin VB.Label Label2 
                     Caption         =   "% "
                     Height          =   255
                     Index           =   23
                     Left            =   2280
                     TabIndex        =   151
                     Top             =   2550
                     Width           =   735
                  End
                  Begin VB.Label Label2 
                     Caption         =   "I.V.A."
                     Height          =   255
                     Index           =   24
                     Left            =   120
                     TabIndex        =   150
                     Top             =   2550
                     Width           =   2175
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Ora arrivo"
                     Height          =   270
                     Index           =   25
                     Left            =   16200
                     TabIndex        =   149
                     Top             =   2040
                     Width           =   975
                  End
               End
               Begin DmtGridCtl.DmtGrid GrigliaImballi 
                  Height          =   5295
                  Left            =   -74880
                  TabIndex        =   83
                  Top             =   2040
                  Width           =   15735
                  _ExtentX        =   27755
                  _ExtentY        =   9340
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
               Begin DmtGridCtl.DmtGrid GrigliaAddebiti 
                  Height          =   5175
                  Left            =   -74880
                  TabIndex        =   171
                  Top             =   2040
                  Width           =   15735
                  _ExtentX        =   27755
                  _ExtentY        =   9128
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
               Begin DmtCodDescCtl.DmtCodDesc CDImballoGestione 
                  Height          =   615
                  Left            =   -74880
                  TabIndex        =   75
                  Top             =   720
                  Width           =   5535
                  _ExtentX        =   9763
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4AF27
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4AF76
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4AFD5
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
               Begin DMTDataCmb.DMTCombo cboTipoProcessoImballo 
                  Height          =   315
                  Left            =   -67800
                  TabIndex        =   76
                  Top             =   960
                  Width           =   1335
                  _ExtentX        =   2355
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
               Begin DmtCodDescCtl.DmtCodDesc CDArticoliAddebiti 
                  Height          =   615
                  Left            =   -74880
                  TabIndex        =   84
                  Top             =   720
                  Width           =   5535
                  _ExtentX        =   9763
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4B02F
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4B07E
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4B0DE
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
               Begin DMTDataCmb.DMTCombo cboIvaAddebiti 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   88
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   2175
                  _ExtentX        =   3836
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
               Begin DMTEDITNUMLib.dmtNumber txtAliquotaIvaAddebiti 
                  Height          =   315
                  Left            =   -72720
                  TabIndex        =   89
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboTipoTrattenutaAggiuntiva 
                  Height          =   315
                  Left            =   -66600
                  TabIndex        =   87
                  Top             =   960
                  Width           =   4695
                  _ExtentX        =   8281
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
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleRigaAddebiti 
                  Height          =   315
                  Left            =   -69240
                  TabIndex        =   92
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImpostaAddebiti 
                  Height          =   315
                  Left            =   -70560
                  TabIndex        =   91
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImponibileAddebiti 
                  Height          =   315
                  Left            =   -71880
                  TabIndex        =   90
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImpUniAddebiti 
                  Height          =   315
                  Left            =   -67800
                  TabIndex        =   86
                  Top             =   960
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
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
               Begin DMTEDITNUMLib.dmtNumber txtQtaAddebiti 
                  Height          =   315
                  Left            =   -69360
                  TabIndex        =   85
                  Top             =   960
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboUMImbGest 
                  Height          =   315
                  Left            =   -69360
                  TabIndex        =   182
                  TabStop         =   0   'False
                  Top             =   960
                  Width           =   1455
                  _ExtentX        =   2566
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
               Begin DMTEDITNUMLib.dmtNumber txtDispLottoImballo 
                  Height          =   255
                  Left            =   -71400
                  TabIndex        =   214
                  Top             =   2520
                  Visible         =   0   'False
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtGiacLottoImballo 
                  Height          =   255
                  Left            =   -72360
                  TabIndex        =   215
                  Top             =   2520
                  Visible         =   0   'False
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDLottoImballoGest 
                  Height          =   315
                  Left            =   -73560
                  TabIndex        =   217
                  Top             =   2520
                  Visible         =   0   'False
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtImpUniAltriImb 
                  Height          =   315
                  Left            =   -64920
                  TabIndex        =   78
                  Top             =   960
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin VB.Label Label2 
                  Caption         =   "Riferimento esterno"
                  Height          =   270
                  Index           =   45
                  Left            =   -69000
                  TabIndex        =   261
                  Top             =   1320
                  Width           =   4455
               End
               Begin VB.Label Label2 
                  Caption         =   "Importo"
                  Height          =   270
                  Index           =   41
                  Left            =   -64920
                  TabIndex        =   228
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Lotto imballo"
                  Height          =   270
                  Index           =   38
                  Left            =   -74880
                  TabIndex        =   212
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Unità di misura"
                  Height          =   270
                  Index           =   35
                  Left            =   -69360
                  TabIndex        =   183
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.Label Label2 
                  Caption         =   "Quantità"
                  Height          =   270
                  Index           =   34
                  Left            =   -66360
                  TabIndex        =   181
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Quantità"
                  Height          =   255
                  Index           =   33
                  Left            =   -69360
                  TabIndex        =   180
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Totale riga"
                  Height          =   255
                  Index           =   32
                  Left            =   -69240
                  TabIndex        =   179
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Imposta riga"
                  Height          =   255
                  Index           =   31
                  Left            =   -70560
                  TabIndex        =   178
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Imponibile riga"
                  Height          =   255
                  Index           =   30
                  Left            =   -71880
                  TabIndex        =   177
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Importo unitario"
                  Height          =   255
                  Index           =   29
                  Left            =   -67800
                  TabIndex        =   176
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Caption         =   "Tipo trattenuta"
                  Height          =   255
                  Index           =   2
                  Left            =   -66600
                  TabIndex        =   175
                  Top             =   720
                  Width           =   2535
               End
               Begin VB.Label Label2 
                  Caption         =   "% "
                  Height          =   255
                  Index           =   28
                  Left            =   -72720
                  TabIndex        =   174
                  Top             =   1320
                  Width           =   735
               End
               Begin VB.Label Label2 
                  Caption         =   "I.V.A."
                  Height          =   255
                  Index           =   27
                  Left            =   -74880
                  TabIndex        =   173
                  Top             =   1320
                  Width           =   2175
               End
               Begin VB.Label Label2 
                  Caption         =   "Tipo processo"
                  Height          =   270
                  Index           =   26
                  Left            =   -67800
                  TabIndex        =   172
                  Top             =   720
                  Width           =   1335
               End
            End
            Begin VB.Frame FraOrdineCliente 
               Caption         =   "ORDINE CLIENTE PER LAVORAZIONE AUTOMATICA"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   855
               Left            =   120
               TabIndex        =   230
               Top             =   2520
               Width           =   17655
               Begin VB.CommandButton cmdSelezionaOrdine 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   60
                  Picture         =   "frmMain.frx":4B138
                  Style           =   1  'Graphical
                  TabIndex        =   232
                  TabStop         =   0   'False
                  ToolTipText     =   "Trova ordine"
                  Top             =   480
                  Width           =   375
               End
               Begin VB.CommandButton cmdEliminaRifOrdine 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   480
                  Picture         =   "frmMain.frx":4B6C2
                  Style           =   1  'Graphical
                  TabIndex        =   231
                  ToolTipText     =   "Aggiorna ordine cliente"
                  Top             =   480
                  Width           =   375
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDOrdineCliente 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   233
                  Top             =   480
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
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
               Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdine 
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   234
                  Top             =   480
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DmtCodDescCtl.DmtCodDesc cdCliente 
                  Height          =   615
                  Left            =   2760
                  TabIndex        =   235
                  TabStop         =   0   'False
                  Top             =   225
                  Width           =   4455
                  _ExtentX        =   7858
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4BC4C
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4BC9A
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4BCEC
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
               Begin DMTDATETIMELib.dmtDate txtDataOrdine 
                  Height          =   315
                  Left            =   8400
                  TabIndex        =   236
                  Top             =   480
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataPartenza 
                  Height          =   315
                  Left            =   10680
                  TabIndex        =   237
                  Top             =   480
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTLblLinkCtl.LabelLink lblLinkOrdine 
                  Height          =   255
                  Left            =   4320
                  TabIndex        =   238
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  Caption         =   "Nuovo ordine"
                  Name            =   "LabelLink"
               End
               Begin DMTDataCmb.DMTCombo cboDestinazione 
                  Height          =   315
                  Left            =   12240
                  TabIndex        =   239
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTEDITNUMLib.dmtNumber txtNListaPrelievo 
                  Height          =   315
                  Left            =   9840
                  TabIndex        =   240
                  Top             =   480
                  Width           =   735
                  _Version        =   65536
                  _ExtentX        =   1296
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDOrdinePadre 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   241
                  Top             =   1320
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
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
               Begin VB.Label Label5 
                  Caption         =   "Numero"
                  Height          =   255
                  Index           =   8
                  Left            =   7200
                  TabIndex        =   249
                  Top             =   285
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Caption         =   "Data ordine"
                  Height          =   255
                  Index           =   11
                  Left            =   8400
                  TabIndex        =   248
                  Top             =   285
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Caption         =   "Identificativo ordine"
                  Height          =   255
                  Index           =   10
                  Left            =   840
                  TabIndex        =   247
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label Label4 
                  Caption         =   "Data partenza"
                  Height          =   255
                  Index           =   9
                  Left            =   10680
                  TabIndex        =   246
                  Top             =   285
                  Width           =   1215
               End
               Begin VB.Label Label4 
                  Caption         =   "Destinazione diversa"
                  Height          =   255
                  Index           =   5
                  Left            =   12240
                  TabIndex        =   245
                  Top             =   285
                  Width           =   3615
               End
               Begin VB.Label Label5 
                  Caption         =   "N° lista"
                  Height          =   255
                  Index           =   6
                  Left            =   9840
                  TabIndex        =   244
                  Top             =   285
                  Width           =   735
               End
               Begin VB.Label Label4 
                  Caption         =   "Ordine padre"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   243
                  Top             =   1125
                  Width           =   1815
               End
               Begin VB.Label lblNomeCliente 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   9840
                  TabIndex        =   242
                  Top             =   480
                  Width           =   4335
               End
            End
            Begin VB.Frame fraImpAttesa 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   2415
               Left            =   5880
               TabIndex        =   218
               Top             =   3720
               Visible         =   0   'False
               Width           =   6255
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
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
                  ForeColor       =   &H80000008&
                  Height          =   1815
                  Left            =   0
                  Picture         =   "frmMain.frx":4BD46
                  ScaleHeight     =   1815
                  ScaleWidth      =   6225
                  TabIndex        =   220
                  Top             =   0
                  Width           =   6225
                  Begin VB.Label lblInfo2 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "ATTENDERE"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   255
                     Left            =   120
                     TabIndex        =   221
                     Top             =   1500
                     Width           =   6015
                  End
               End
               Begin MSComctlLib.ProgressBar ProgressBar2 
                  Height          =   135
                  Left            =   120
                  TabIndex        =   219
                  Top             =   2160
                  Width           =   6015
                  _ExtentX        =   10610
                  _ExtentY        =   238
                  _Version        =   393216
                  Appearance      =   0
                  Scrolling       =   1
               End
               Begin VB.Label lblInfo 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   222
                  Top             =   1920
                  Width           =   5895
               End
            End
            Begin VB.Frame Frame2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1935
               Left            =   120
               TabIndex        =   196
               Top             =   600
               Width           =   5535
               Begin VB.TextBox txtProvincia 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  Height          =   285
                  Left            =   4800
                  Locked          =   -1  'True
                  TabIndex        =   200
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox txtComune 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  Height          =   285
                  Left            =   960
                  Locked          =   -1  'True
                  TabIndex        =   199
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   3735
               End
               Begin VB.TextBox txtCAP 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  Height          =   285
                  Left            =   120
                  Locked          =   -1  'True
                  TabIndex        =   198
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   735
               End
               Begin VB.TextBox txtIndirizzo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  Height          =   285
                  Left            =   120
                  Locked          =   -1  'True
                  TabIndex        =   197
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   5295
               End
               Begin DMTDataCmb.DMTCombo cboVettore 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   11
                  Top             =   1080
                  Width           =   3495
                  _ExtentX        =   6165
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
               Begin DMTDataCmb.DMTCombo cboLuogoPresaMerce 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   13
                  Top             =   1560
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
               Begin DMTEDITNUMLib.dmtNumber txtImportoTrasporto 
                  Height          =   315
                  Left            =   3720
                  TabIndex        =   12
                  Top             =   1080
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboTipoOrdConf 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   14
                  Top             =   1560
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
               Begin VB.Label Label9 
                  Caption         =   "Conformità"
                  Height          =   255
                  Index           =   1
                  Left            =   3120
                  TabIndex        =   225
                  Top             =   1360
                  Width           =   2295
               End
               Begin VB.Line Line3 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   120
                  X2              =   5400
                  Y1              =   840
                  Y2              =   840
               End
               Begin VB.Label Label8 
                  Caption         =   "Vettore"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   203
                  Top             =   885
                  Width           =   2535
               End
               Begin VB.Label Label9 
                  Caption         =   "Sede stoccaggio merce"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   202
                  Top             =   1360
                  Width           =   2535
               End
               Begin VB.Label Label8 
                  Caption         =   "Importo trasporto"
                  Height          =   255
                  Index           =   1
                  Left            =   3720
                  TabIndex        =   201
                  Top             =   885
                  Width           =   1695
               End
            End
            Begin VB.Frame Frame1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2535
               Left            =   5880
               TabIndex        =   104
               Top             =   0
               Width           =   11895
               Begin VB.TextBox txtNLetteraIntento 
                  Height          =   315
                  Left            =   7080
                  Locked          =   -1  'True
                  TabIndex        =   259
                  Top             =   360
                  Width           =   495
               End
               Begin VB.CheckBox chkTrattaComeConf 
                  Caption         =   "Tratta come conferimento"
                  Height          =   255
                  Left            =   8880
                  TabIndex        =   258
                  Top             =   360
                  Width           =   2895
               End
               Begin VB.CheckBox chkPreConferimento 
                  Caption         =   "Generato da pre-conferimento"
                  Enabled         =   0   'False
                  Height          =   375
                  Left            =   7680
                  TabIndex        =   257
                  Top             =   840
                  Width           =   3975
               End
               Begin VB.CheckBox chkSelLottoProdAnaFatt 
                  Caption         =   "Seleziona lotto di produzione da anagrafica di fatt."
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   6720
                  TabIndex        =   254
                  Top             =   2040
                  Width           =   5055
               End
               Begin VB.TextBox txtTargaAutomezzo 
                  Height          =   315
                  Left            =   6960
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   1440
                  Width           =   1815
               End
               Begin VB.CommandButton cmdEliminaRifLetInt 
                  Height          =   315
                  Left            =   6360
                  Picture         =   "frmMain.frx":55500
                  Style           =   1  'Graphical
                  TabIndex        =   192
                  ToolTipText     =   "Elimina riferimento lettera intento"
                  Top             =   360
                  Width           =   375
               End
               Begin VB.CommandButton cmdLetteraIntento 
                  Height          =   315
                  Left            =   6720
                  Picture         =   "frmMain.frx":55A8A
                  Style           =   1  'Graphical
                  TabIndex        =   191
                  ToolTipText     =   "Lettere di intento del socio/fornitore"
                  Top             =   360
                  Width           =   375
               End
               Begin VB.TextBox txtNomeSocioFatt 
                  BackColor       =   &H8000000F&
                  Height          =   340
                  Left            =   4440
                  Locked          =   -1  'True
                  TabIndex        =   188
                  TabStop         =   0   'False
                  Top             =   1980
                  Width           =   2175
               End
               Begin VB.TextBox txtNumeroDocumentoAcq 
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   10
                  Top             =   1460
                  Width           =   1815
               End
               Begin VB.CommandButton cmdAltriDatiDocumento 
                  Height          =   315
                  Left            =   8760
                  Picture         =   "frmMain.frx":56014
                  Style           =   1  'Graphical
                  TabIndex        =   115
                  TabStop         =   0   'False
                  ToolTipText     =   "Altri dati del documento"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   375
               End
               Begin VB.TextBox txtPrefissoConferimento 
                  Height          =   315
                  Left            =   6240
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   880
                  Width           =   1335
               End
               Begin DMTEDITNUMLib.dmtNumber txtNumeroDocumento 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   880
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo CboMagazzinoVend 
                  Height          =   315
                  Left            =   7560
                  TabIndex        =   97
                  TabStop         =   0   'False
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   495
                  _ExtentX        =   873
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
               Begin DMTDataCmb.DMTCombo cboMagazzinoConf 
                  Height          =   315
                  Left            =   6960
                  TabIndex        =   96
                  TabStop         =   0   'False
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   495
                  _ExtentX        =   873
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
               Begin DMTDataCmb.DMTCombo cboSezionale 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   2
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   2055
                  _ExtentX        =   3625
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
               Begin DMTDATETIMELib.dmtDate txtDataDocumento 
                  Height          =   315
                  Left            =   3120
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   885
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboEsercizio 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   3
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   3975
                  _ExtentX        =   7011
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
               Begin DMTEDITNUMLib.dmtNumber txtNumeroConferimento 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   885
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataConferimento 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   885
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDATETIMELib.dmtDate txtDataDocumentoAcq 
                  Height          =   315
                  Left            =   5280
                  TabIndex        =   15
                  Top             =   1455
                  Width           =   1575
                  _Version        =   65536
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboTipoDocumentoAcq 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   9
                  Top             =   1455
                  Width           =   3135
                  _ExtentX        =   5530
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
               Begin DmtCodDescCtl.DmtCodDesc CDSocioFatt 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   17
                  Top             =   1740
                  Width           =   5175
                  _ExtentX        =   9128
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":5659E
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":565EC
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":56651
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
               Begin DMTDATETIMELib.dmtDate txtDataLetteraIntento 
                  Height          =   315
                  Left            =   7560
                  TabIndex        =   193
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDLetteraIntento 
                  Height          =   255
                  Left            =   8040
                  TabIndex        =   194
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   615
                  _Version        =   65536
                  _ExtentX        =   1085
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin VB.Label Label1 
                  Caption         =   "Targa automezzo"
                  Height          =   255
                  Index           =   11
                  Left            =   6960
                  TabIndex        =   250
                  Top             =   1200
                  Width           =   1815
               End
               Begin VB.Label lblLetteraIntento 
                  Caption         =   "Lettera d'intento"
                  Height          =   255
                  Left            =   7080
                  TabIndex        =   195
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "Data Doc. Acq."
                  Height          =   255
                  Index           =   4
                  Left            =   5280
                  TabIndex        =   187
                  Top             =   1260
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  Caption         =   "N° Doc. Acq."
                  Height          =   255
                  Index           =   8
                  Left            =   3360
                  TabIndex        =   186
                  Top             =   1260
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Tipo documento"
                  Height          =   255
                  Index           =   9
                  Left            =   120
                  TabIndex        =   185
                  Top             =   1260
                  Width           =   2775
               End
               Begin VB.Label Label1 
                  Caption         =   "Data Liquidazione"
                  Height          =   255
                  Index           =   3
                  Left            =   3120
                  TabIndex        =   114
                  Top             =   675
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "Prefisso conf."
                  Height          =   255
                  Index           =   7
                  Left            =   6360
                  TabIndex        =   113
                  Top             =   675
                  Width           =   1215
               End
               Begin VB.Label Label1 
                  Caption         =   "N° Doc. socio"
                  Height          =   255
                  Index           =   6
                  Left            =   4800
                  TabIndex        =   112
                  Top             =   675
                  Width           =   1335
               End
               Begin VB.Label Label1 
                  Caption         =   "Esercizio"
                  Height          =   255
                  Index           =   5
                  Left            =   2280
                  TabIndex        =   111
                  Top             =   120
                  Width           =   1815
               End
               Begin VB.Label Label1 
                  Caption         =   "N° Doc."
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   107
                  Top             =   680
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Data consegna"
                  Height          =   255
                  Index           =   1
                  Left            =   1320
                  TabIndex        =   106
                  Top             =   680
                  Width           =   1335
               End
               Begin VB.Label Label1 
                  Caption         =   "Sezionale"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   105
                  Top             =   120
                  Width           =   2055
               End
            End
            Begin VB.TextBox txtNomeSocio 
               BackColor       =   &H8000000F&
               Height          =   340
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   1
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin DmtCodDescCtl.DmtCodDesc CDSocio 
               Height          =   615
               Left            =   120
               TabIndex        =   0
               Top             =   0
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   1085
               PropCodice      =   $"frmMain.frx":566AB
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"frmMain.frx":566F9
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"frmMain.frx":56753
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
            Begin VB.TextBox txtAnnotazioni 
               Height          =   375
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   10320
               Width           =   17655
            End
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   135
               Left            =   120
               TabIndex        =   108
               Top             =   10800
               Width           =   17655
               _ExtentX        =   31141
               _ExtentY        =   238
               _Version        =   393216
               Appearance      =   0
            End
            Begin VB.Label Label3 
               Caption         =   "Annotazioni"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   109
               Top             =   10080
               Width           =   5535
            End
         End
         Begin DmtGridCtl.DmtGrid BrwMain 
            Height          =   735
            Left            =   0
            TabIndex        =   229
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1296
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
            ColumnsHeaderHeight=   20
         End
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
         Left            =   720
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   101
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   9075
         Left            =   0
         TabIndex        =   110
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
         Top             =   360
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   0
         MousePointer    =   9  'Size W E
         Top             =   120
         Width           =   60
      End
      Begin VB.Line Line2 
         X1              =   1560
         X2              =   6480
         Y1              =   3360
         Y2              =   3360
      End
   End
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   98
      Top             =   11220
      Width           =   18285
      _ExtentX        =   32253
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
'La collezione dei campi del documento
'collegati ai controlli del form
Private m_FormFields As FormFields
'Il campo con la proprietà TabIndex uguale a 0
Private m_ControlTabIndex0 As Control
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

'cbcx
'Oggetto adibito alla gestione del processo On_Extend
'Private m_ExtendApplication As DmtExtendAppLib.ExtendApplication


'rif1
'L'oggetto per la gestione dei sottodocumenti
'RATEIZZAZIONE
Private WithEvents m_DocumentsLink As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink.VB_VarHelpID = -1
'PRODOTTI
Private WithEvents m_DocumentsLink1 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink1.VB_VarHelpID = -1
'ADEGUAMENTI
Private WithEvents m_DocumentsLink2 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink2.VB_VarHelpID = -1

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

'///////////////////////////////////////////////////////////////////////////////////
' ATTENZIONE:
' Occorre impostare questa costante!
' (ed eventualmente personalizzare il codice della funzione Caption2Display
'///////////////////////////////////////////////////////////////////////////////////
' Costante che identifica il campo più significativo del documento, il cui valore
' verrà visualizzato nella Caption del Form ed in quei messaggi in cui è mostrato
' il contenuto del campo principale del documento attivo.
' La costante può essere una stringa tipo "NomeCampo" o un intero che funge da indice
' nella collection m_Document.Fields().
'(Se l'applicazione può essere chiamata da un link occorre impostare anche la variabile
'sMessage1 presente nel metodo FormUnload.)
Private Const CAMPO_PER_CAPTION = "Anagrafica"


'Versione del controllo ActiveBar
Private Const BARMENUVERSION = "3.0"
'Variabile per la gestione degli shortcut del Menu
Private aryShortCut(1) As New ActiveBar3LibraryCtl.ShortCut



Private oReport As dmtReportLib.dmtReport

'****************************VARIABILI CONTRATTO**********************************
'VARIBILE ARRAY IN CUI VENGONO INSERITE L'ID DELLE RIGHE DA ELIMINARE DEFINITIVAMENTE
Private ArrayDelete(150) As Long
'VARIABILE CONTATATORE CHE SERVE PER IL CONTEGGIO DELLE LINEE DA ELIMINARE
Private ContDelete As Long

'VARIBILE ARRAY IN CUI VENGONO INSERITE L'ID DELLE RIGHE DA ELIMINARE DEFINITIVAMENTE
Private ArrayDeleteConferimento(150) As Long
'VARIABILE CONTATATORE CHE SERVE PER IL CONTEGGIO DELLE LINEE DA ELIMINARE
Private ContDeleteConferimento As Long



'VARIABILE CHE SERVE PER VEDERE SE IL DOCUMENTO E' IN AGGIORNAMENTO IN MODO CHE ALL'EVENTO
'ON REPOSITION DEL DOTTODOCUMENTO NON VENGANO EFFETTUATI DI NUOVO TUTTI GLI AGGIORNAMENTI
Private AggiornamentoDocumento As Integer

'VARIABILE CHE IMPOSTA UNA NUOVA RIGA
'0 = Nuova riga
'1 = Riga in modifica
Private Nuova_Riga As Integer

Private bVariazioneDettaglio As Boolean

'******************************************************************************

Private Mov As DmtMovim.cMovimentazione

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
            
            Caption2Display = m_App.Caption & ": " & fnNotNull(m_Document.Fields(CAMPO_PER_CAPTION).Value) & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
        Else
            
            Caption2Display = m_App.Caption & ": " & fnNotNull(BrwMain.AllColumns(CAMPO_PER_CAPTION).Value) & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
        End If
    Else
        Caption2Display = m_App.FunctionName & " [Versione " & App.Major & "." & App.Minor & "." & App.Revision & "]"
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
    
    '///////////////////////////////////////////////////////////////////
    'Inserire qui il codice di controllo sulla validità e consistenza
    'dei dati da salvare.
    '///////////////////////////////////////////////////////////////////
    
If Link_Esercizio <= 0 Then
    MsgBox "L'esercizio per la data del conferimento è inesistente", vbCritical, "Validazione dati"
    Me.txtDataDocumento.SetFocus
    PermissionToSave = False
    Exit Function
End If
If Link_PeriodoIVA <= 0 Then
    MsgBox "Il periodo I.V.A. inerente alla data di conferimento è inesistente", vbCritical, "Validazione dati"
    Me.txtDataDocumento.SetFocus
    PermissionToSave = False
    Exit Function
End If

If fnNotNullN(Me.CDSocio.KeyFieldID) = 0 Then
    MsgBox "Manca il nome del socio", vbCritical, "Validazione dati"
    Me.CDSocio.SetFocus
    PermissionToSave = False
    Exit Function
End If
If Me.cboMagazzinoConf.CurrentID = 0 Then
    MsgBox "Manca il magazzino di conferimento", vbCritical, "Validazione dati"
    Me.cboMagazzinoConf.SetFocus
    PermissionToSave = False
    Exit Function
End If
If Me.CboMagazzinoVend.CurrentID = 0 Then
    MsgBox "Manca il magazzino di vendita", vbCritical, "Validazione dati"
    Me.CboMagazzinoVend.SetFocus
    PermissionToSave = False
    Exit Function
End If
    
If Me.cboSezionale.CurrentID = 0 Then
    MsgBox "Manca il sezionale", vbCritical, "Validazione dati"
    Me.cboSezionale.SetFocus
    PermissionToSave = False
    Exit Function
End If
    
If Me.txtNumeroDocumento.Value = 0 Then
    MsgBox "Manca il numero del documento", vbCritical, "Validazione dati"
    Me.txtNumeroDocumento.SetFocus
    PermissionToSave = False
    Exit Function
End If
    
If Me.txtDataDocumento.Value = 0 Then
    MsgBox "Manca la data di liquidazione", vbCritical, "Validazione dati"
    Me.txtDataDocumento.SetFocus
    PermissionToSave = False
    Exit Function
End If
    
If Me.txtDataConferimento.Value = 0 Then
    MsgBox "Manca la data di consegna merce", vbCritical, "Validazione dati"
    Me.txtDataConferimento.SetFocus
    PermissionToSave = False
    Exit Function
End If
    
If Me.chkPreConferimento.Value = vbUnchecked Then
    If GET_ESISTENZA_NUMERO_DOCUMENTO = True Then
        'MsgBox "Il numero documento è già stato associato ad un precedente documento", vbCritical, "Impossibile salvare"
        Me.txtNumeroDocumento.Enabled = True
        Me.txtNumeroDocumento.Value = fnGetNumeroDocumento
        MsgBox "Al documento verrà associato il numero " & Me.txtNumeroDocumento.Value, vbInformation, "Validazione dati"
        PermissionToSave = False
        Exit Function
    End If
End If

PermissionToSave = True

End Function

Private Function GET_ESISTENZA_NUMERO_DOCUMENTO() As Boolean
Dim sSQL  As String
Dim rs As DmtOleDbLib.adoResultset

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
    sSQL = "SELECT IDRV_POCaricoMerceTesta FROM RV_POCaricoMerceTesta "
    sSQL = sSQL & "WHERE IDSezionale=" & Me.cboSezionale.CurrentID
    sSQL = sSQL & " AND IDAzienda=" & m_App.IDFirm
    sSQL = sSQL & " AND NumeroDocumento=" & Me.txtNumeroDocumento.Value
    sSQL = sSQL & " AND IDTipoDocumentoCoop=" & 2
    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate("01/01/" & Year(Me.txtDataDocumento.Text))
    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate("31/12/" & Year(Me.txtDataDocumento.Text))
Else
    sSQL = "SELECT IDRV_POCaricoMerceTesta FROM RV_POCaricoMerceTesta "
    sSQL = sSQL & "WHERE IDSezionale=" & Me.cboSezionale.CurrentID
    sSQL = sSQL & " AND IDAzienda=" & m_App.IDFirm
    sSQL = sSQL & " AND NumeroDocumento=" & Me.txtNumeroDocumento.Value
    sSQL = sSQL & " AND IDTipoDocumentoCoop=" & 2
    sSQL = sSQL & " AND DataDocumento>=" & fnNormDate("01/01/" & Year(Me.txtDataDocumento.Text))
    sSQL = sSQL & " AND DataDocumento<=" & fnNormDate("31/12/" & Year(Me.txtDataDocumento.Text))
    sSQL = sSQL & " AND IDRV_POCaricoMerceTesta<>" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
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

Private Function ControlloNumeroDocumento() As Boolean
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
    ControlloNumeroDocumento = True
    
    sSQL = "SELECT * FROM Oggetto "
    sSQL = sSQL & "WHERE ((Numero=" & fnNormString(Me.txtNumeroDocumento.Value) & ") "
    sSQL = sSQL & "AND (IDSezionale=" & Me.cboSezionale.CurrentID & ") "
    sSQL = sSQL & "AND (DataEmissione >=" & fnNormDate(DataInizio_Esercizio) & ") "
    sSQL = sSQL & "AND (DataEmissione <=" & fnNormDate(DataFine_Esercizio) & "))"
    
    
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        ControlloNumeroDocumento = False
    End If
    
    rs.CloseResultset
    Set rs = Nothing
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
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            m_Document.MovePrevious
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
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
        
        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            m_Document.MoveNext
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
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
    
    
    BrwMain.Conditions.ClearValues

    'Annulla una eventuale operazione precedente.
    If m_Document.TableNew Then
        m_Document.AbortNew
    End If

    'Creazione buffers vuoti
    m_Document.NewDoc
    
    GENERA_FILTRO_PER_TIPO_OGGETTO
    
    'Refresh delle variabili di stato
    m_Search = False
    m_Changed = False
    m_Saved = False
    
    'Refresh della toolbar in modalità inserimento
    SetStatus4Modality Insert
    
    'Ripristina la vista del Form
    BrwMain.Visible = False
    
    CreateBrowserConditions
    
    'Il primo campo del Form riceve l'input focus
    SetFocusTabIndex0


   

    
    
End Sub
Private Function GetConteggioRecord(NomeTabella As String, CampoTabella As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT COUNT(" & CampoTabella & ") AS Conteggio "
sSQL = sSQL & "FROM  " & NomeTabella
sSQL = sSQL & " WHERE " & CampoTabella & "=1"

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GetConteggioRecord = 0
Else
    GetConteggioRecord = fnNotNullN(rs!Conteggio)
End If

rs.CloseResultset
Set rs = Nothing
End Function


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
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDAnagrafica = 0
        ctrControl.Description = ""
        ctrControl.SecondDescription = ""
        
    ElseIf sType = "dmtCurrency" Or sType = "dmtNumber" Then
        ctrControl.Value = 0
        ctrControl.Text = ""
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDNode = 0
    ElseIf sType = "DmtCodDesc" Then
        ctrControl.Load 0

    ElseIf sType = "DmtFirmGerarchy" Then
        ctrControl.LoadActivity 0
    ElseIf sType = "DMTProgControl" Then
        'Queste istruzioni forzano il refresh
        'e il reset del componente
        ctrControl.IDArticolo = 0
        ctrControl.Show
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
    Dim cField As FormField
    
    For Each cField In m_FormFields
        'Viene ripulito il campo di immissione.
        ClearControl cField.Control
    Next
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
        Case "ControlloCollegamenti"
             cmdControlloCollegamenti_Click
        Case "Tour"
             cmdTour_Click
        Case "CreaDocumento"
             cmdCreaDocumento_Click
        Case "CalcoloMedi"
             cmdCalcoloMedi_Click
        Case "Annotazioni"
             cmdAnnotazioni_Click
        Case "ChiudiLotto"
             cmdChiudiLotto_Click
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

    'rif3 start
    
    Dim Fields As DmtDocManLib.Fields
    Dim Control As Control
    Dim Field As FormField
    
    On Error Resume Next
    
    'In questi casi non si deve far nulla
    If Not (m_Document.EOF = True Or m_Document.BOF = True) Then

       'Passa alla collezione Fields dell'oggetto
        'Document i valori da salvare
        For Each Field In m_FormFields
    
            Select Case TypeName(Field.Control)
                Case "TextBox"
                    Field.Control.Text = fnNotNull(m_Document.Fields(Field.Name).Value)
                Case "DMTCombo"
                    Field.Control.WriteOn fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "Town"
                    If Field.Name = "IDComune" Then
                        Field.Control.TownID = fnNotNullN(m_Document.Fields(Field.Name).Value)
                    ElseIf Field.Name = "Cap" Then
                        Field.Control.Zip = fnNotNull(m_Document.Fields(Field.Name).Value)
                    End If
                Case "dmtDate"
                    Field.Control.Text = fnNotNull(m_Document.Fields(Field.Name).Value)
                Case "dmtNumber"
                    Field.Control.Text = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtCurrency"
                    Field.Control.Text = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtTime"
                    Field.Control.Text = fnNotNull(m_Document.Fields(Field.Name).Value)
                Case "DmtSearchACS"
                        Field.Control.Description = fnNotNull(m_Document.Fields("Anagrafica").Value)
                        Field.Control.SecondDescription = fnNotNull(m_Document.Fields("Nome").Value)
                        Field.Control.IDAnagrafica = m_Document.Fields(Field.Name).Value
                Case "CheckBox"
                    Field.Control.Value = Abs(fnNotNullN(m_Document.Fields(Field.Name).Value))
                Case "DmtCodDesc"
                    Field.Control.Load fnNotNullN(m_Document.Fields(Field.Name).Value)
            End Select
           
        Next

  
    End If

    'rif3 end
    
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
'Nome: CreateFormFields
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Crea la collezione FormFields che associa i campi del
'documento con i controlli di input del Form. Vengono
'anche creati i controlli del Form necessari e calcolato
'il layout del Form.
'**/
Private Sub CreateFormFields()
    Dim Field As FormField
        
        
    'rif2 start
    
    'Se non esiste il documento aperto non si può creare la collezione
    If m_Document Is Nothing Then Exit Sub
    
    'Se la collezione è già stata creata esce
    If Not m_FormFields Is Nothing Then Exit Sub
    
    'Istanzia la collezione.  Il codice sottostante viene eseguito soltanto la prima volta
    Set m_FormFields = New FormFields
    
    'rif2   End
    
    
    'IDAnagraficaSocio
    Set Field = New FormField
    Set Field.Control = Me.CDSocio
    Field.Name = "IDAnagrafica"
    Field.Visible = True
    Me.CDSocio.Tag = Field.Name
    m_FormFields.Add Field
    
   
    'IDSezionale
    Set Field = New FormField
    Set Field.Control = Me.cboSezionale
    Field.Name = "IDSezionale"
    Field.Visible = True
    Me.cboSezionale.Tag = Field.Name
    m_FormFields.Add Field

    'IDEsercizio
    Set Field = New FormField
    Set Field.Control = Me.cboEsercizio
    Field.Name = "IDEsercizio"
    Field.Visible = True
    Me.cboEsercizio.Tag = Field.Name
    m_FormFields.Add Field

    'IDMagazzinoDiConferimento
    Set Field = New FormField
    Set Field.Control = Me.cboMagazzinoConf
    Field.Name = "IDMagazzinoConferimento"
    Field.Visible = True
    Me.cboMagazzinoConf.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDMagazzinoDiVendita
    Set Field = New FormField
    Set Field.Control = Me.CboMagazzinoVend
    Field.Name = "IDMagazzinoVendita"
    Field.Visible = True
    Me.CboMagazzinoVend.Tag = Field.Name
    m_FormFields.Add Field
    
    'Numero documento
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroDocumento
    Field.Name = "NumeroDocumento"
    Field.Visible = True
    Me.txtNumeroDocumento.Tag = Field.Name
    m_FormFields.Add Field
        
    'Data documento
    Set Field = New FormField
    Set Field.Control = Me.txtDataDocumento
    Field.Name = "DataDocumento"
    Field.Visible = True
    Me.txtDataDocumento.Tag = Field.Name
    m_FormFields.Add Field
        
    'Annotazioni
    Set Field = New FormField
    Set Field.Control = Me.txtAnnotazioni
    Field.Name = "Annotazioni"
    Field.Visible = True
    Me.txtAnnotazioni.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indirizzo
    Set Field = New FormField
    Set Field.Control = Me.txtIndirizzo
    Field.Name = "Indirizzo"
    Field.Visible = True
    Me.txtIndirizzo.Tag = Field.Name
    m_FormFields.Add Field

    'Cap
    Set Field = New FormField
    Set Field.Control = Me.txtCAP
    Field.Name = "Cap"
    Field.Visible = True
    Me.txtCAP.Tag = Field.Name
    m_FormFields.Add Field
    
    'Comune
    Set Field = New FormField
    Set Field.Control = Me.txtComune
    Field.Name = "Comune"
    Field.Visible = True
    Me.txtComune.Tag = Field.Name
    m_FormFields.Add Field

    'Provincia
    Set Field = New FormField
    Set Field.Control = Me.txtProvincia
    Field.Name = "Provincia"
    Field.Visible = True
    Me.txtProvincia.Tag = Field.Name
    m_FormFields.Add Field

    'Numero documento socio
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroConferimento
    Field.Name = "NumeroDocumentoSocio"
    Field.Visible = True
    Me.txtNumeroConferimento.Tag = Field.Name
    m_FormFields.Add Field

    'Prefisso documento del socio
    Set Field = New FormField
    Set Field.Control = Me.txtPrefissoConferimento
    Field.Name = "PrefissoNumeroConferimento"
    Field.Visible = True
    Me.txtPrefissoConferimento.Tag = Field.Name
    m_FormFields.Add Field

    'Data di consegna merce
    Set Field = New FormField
    Set Field.Control = Me.txtDataConferimento
    Field.Name = "DataConferimento"
    Field.Visible = True
    Me.txtDataConferimento.Tag = Field.Name
    m_FormFields.Add Field


    'Tipo documento di acquisto
    Set Field = New FormField
    Set Field.Control = Me.cboTipoDocumentoAcq
    Field.Name = "IDTipoDocumentoAcq"
    Field.Visible = True
    Me.cboTipoDocumentoAcq.Tag = Field.Name
    m_FormFields.Add Field

    'Numero documento di Acquisto
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroDocumentoAcq
    Field.Name = "NumeroDocumentoAcq"
    Field.Visible = True
    Me.txtNumeroDocumentoAcq.Tag = Field.Name
    m_FormFields.Add Field

    'data documento di Acquisto
    Set Field = New FormField
    Set Field.Control = Me.txtDataDocumentoAcq
    Field.Name = "DataDocumentoAcq"
    Field.Visible = True
    Me.txtDataDocumentoAcq.Tag = Field.Name
    m_FormFields.Add Field

    'IDAnagraficaFatturazione
    Set Field = New FormField
    Set Field.Control = Me.CDSocioFatt
    Field.Name = "IDAnagraficaFatturazione"
    Field.Visible = True
    Me.CDSocioFatt.Tag = Field.Name
    m_FormFields.Add Field

    'IDVettore
    Set Field = New FormField
    Set Field.Control = Me.cboVettore
    Field.Name = "IDVettore"
    Field.Visible = True
    Me.cboVettore.Tag = Field.Name
    m_FormFields.Add Field

    'IDAltraSedeAzienda
    Set Field = New FormField
    Set Field.Control = Me.cboLuogoPresaMerce
    Field.Name = "IDLuogoPresaMerce"
    Field.Visible = True
    Me.cboLuogoPresaMerce.Tag = Field.Name
    m_FormFields.Add Field

    'IDLetteraIntento
    Set Field = New FormField
    Set Field.Control = Me.txtIDLetteraIntento
    Field.Name = "IDLetteraIntento"
    Field.Visible = True
    Me.txtIDLetteraIntento.Tag = Field.Name
    m_FormFields.Add Field

    'ImportoTrasporto
    Set Field = New FormField
    Set Field.Control = Me.txtImportoTrasporto
    Field.Name = "ImportoTrasporto"
    Field.Visible = True
    Me.txtImportoTrasporto.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDRV_POTipoOrdineConforme
    Set Field = New FormField
    Set Field.Control = Me.cboTipoOrdConf
    Field.Name = "IDRV_POTipoOrdineConforme"
    Field.Visible = True
    Me.cboTipoOrdConf.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDOggettoOrdine
    Set Field = New FormField
    Set Field.Control = Me.txtIDOrdineCliente
    Field.Name = "IDOggettoOrdine"
    Field.Visible = True
    Me.txtIDOrdineCliente.Tag = Field.Name
    m_FormFields.Add Field
    
    'Targa automezzo
    Set Field = New FormField
    Set Field.Control = Me.txtTargaAutomezzo
    Field.Name = "TargaAutomezzo"
    Field.Visible = True
    Me.txtTargaAutomezzo.Tag = Field.Name
    m_FormFields.Add Field
    
    'Attiva selezione lotto di produzione da socio di fatturazione
    Set Field = New FormField
    Set Field.Control = Me.chkSelLottoProdAnaFatt
    Field.Name = "AttivaSelLottoProdAnaFatt"
    Field.Visible = True
    Me.chkSelLottoProdAnaFatt.Tag = Field.Name
    m_FormFields.Add Field
    
    'Tratta come conferimento
    Set Field = New FormField
    Set Field.Control = Me.chkTrattaComeConf
    Field.Name = "TrattaComeConferimento"
    Field.Visible = True
    Me.chkTrattaComeConf.Tag = Field.Name
    m_FormFields.Add Field
    
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
Dim rs As DmtOleDbLib.adoResultset
    'Inserire qui le
    'inizializzazioni da effettuare prima dell'apertura del documento.
    
    
    'rif6 begin
    
    Dim NewLink As DmtDocManLib.Link
    
'**************************CORPO DOCUMENTO*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POCaricoMerceRighe"
    
    Set m_DocumentsLink = m_Document.AddDocumentsLink("RV_POCaricoMerceRighe")
          
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink.PrimaryKey = "IDRV_POCaricoMerceRighe" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sulla tabella Articolo
    Set NewLink = m_DocumentsLink.AddLink("IDUnitaDiMisuraDiamante", "UnitaDiMisura", ltLeft, "IDUnitaDiMisura")
    NewLink.AddLinkColumn "UnitaDiMisura"
    
    'Crea un Link LEFT JOIN sulla tabella RV_POTipoLavorazione
    Set NewLink = m_DocumentsLink.AddLink("IDRV_POTipoLavorazione", "RV_POTipoLavorazione", ltLeft, "IDRV_POTipoLavorazione")
    NewLink.AddLinkColumn "RV_POTipoLavorazione.TipoLavorazione"
    
    'Crea un Link LEFT JOIN sulla tabella RV_POTipoLavorazione
    Set NewLink = m_DocumentsLink.AddLink("IDRV_POTipoConfLiquidazione", "RV_POTipoConfLiquidazione", ltLeft, "IDRV_POTipoConfLiquidazione")
    NewLink.AddLinkColumn "RV_POTipoConfLiquidazione.TipoConfLiquidazione"
    
    m_DocumentsLink.AddOrderedColumn "Link_Ordinamento", ocAscending
'********************************************************************************************************************
    
'**************************CORPO DOCUMENTO***************************************************************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POCaricoMerceRighe"
    
    Set m_DocumentsLink1 = m_Document.AddDocumentsLink("RV_POCaricoMerceImballi")
          
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink1.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink1.PrimaryKey = "IDRV_POCaricoMerceImballi" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sulla tabella Articolo
    Set NewLink = m_DocumentsLink1.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "Articolo.CodiceArticolo"
    NewLink.AddLinkColumn "Articolo.Articolo"
    
    'Crea un Link LEFT JOIN sulla tabella TipoGestioneImballo
    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POTipoProcessoCoop", "RV_POTipoProcessoCoop", ltLeft, "IDRV_POTipoProcessoCoop")
    NewLink.AddLinkColumn "RV_POTipoProcessoCoop.TipoProcessoCoop"

'************************************************************************************

'**************************CORPO DOCUMENTO***************************************************************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POCaricoMerceAddebiti"
    
    Set m_DocumentsLink2 = m_Document.AddDocumentsLink("RV_POCaricoMerceAddebiti")
          
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink2.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink2.PrimaryKey = "IDRV_POCaricoMerceAddebiti" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sulla tabella Articolo
    Set NewLink = m_DocumentsLink2.AddLink("IDArticolo", "Articolo", ltLeft, "IDArticolo")
    NewLink.AddLinkColumn "Articolo.CodiceArticolo"
    NewLink.AddLinkColumn "Articolo.Articolo"
    
    'Crea un Link LEFT JOIN sulla tabella RV_POTipoTrattenutaAggiuntiva
    Set NewLink = m_DocumentsLink2.AddLink("IDRV_POTipoTrattenutaAggiuntiva", "RV_POTipoTrattenutaAggiuntiva", ltLeft, "IDRV_POTipoTrattenutaAggiuntiva")
    NewLink.AddLinkColumn "RV_POTipoTrattenutaAggiuntiva.TipoTrattenuta"

    'Crea un Link LEFT JOIN sulla tabella Iva
    Set NewLink = m_DocumentsLink2.AddLink("IDIva", "Iva", ltLeft, "IDIva")
    NewLink.AddLinkColumn "Iva.Iva"
    NewLink.AddLinkColumn "Iva.AliquotaIva"
    
'************************************************************************************


End Sub


'**+
'Autore: Carlo B. Collovà
'Data creazione: 20/11/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: InitExtensions
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Inizializza la componente adibita alla gestione dell'evento On_Extend
'
'**/
'Private Sub InitExtensions()
    'cbcx
    
    'Istanzia l'oggetto
    'Set m_ExtendApplication = New DmtExtendApp.ExtendApplication
    'Set m_ExtendApplication = New DmtExtendAppLib.ExtendApplication
    
    'Assegna un riferimento all'oggetto Application.
    'In questo modo la maggior parte dei parametri di inizializzazione vengono
    'letti da quest'ultimo
    'Set m_ExtendApplication.Application = m_App
    
    'Assegna un riferimento al controllo ActiveBar affinchè la classe
    'che gestisce i dati aggiuntivi possa interagire con la user interface
    'della manutenzione.
    'Set m_ExtendApplication.MenuBar = BarMenu
        
    'Se la funzione correntemente in esecuzione prevede l'evento On_Extend
    'vengono effettuate tutte le inizializzazioni del caso (come l'aggiunta di bottoni
    ' e menu alla BarMenu, ecc.) altrimenti la classe ExtendApplication non effettua
    'alcuna operazione.
    'm_ExtendApplication.Initialize
    
    'NOTA:
    '-----------------------------------------------------------------------------------------------------
    'Tutte le proprietà di m_ExtendApplication presenti anche nell'interfaccia IExtendApplication ed impostate
    'dopo la chiamata al metodo Initialize saranno settate anche in cContactPlus
    '-----------------------------------------------------------------------------------------------------
    
    'Assegna un riferimento del documento corrente
    'Set m_ExtendApplication.CurrentDocument = m_Document

'End Sub


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
    
    BrwMain.Conditions.Clear

    BrwMain.Conditions.WidthConditions = 350
    BrwMain.Conditions.WidthFields = 250
    BrwMain.Conditions.WidthIntervals = 100
    
    BrwMain.Title.BackColor = vb3DFace
    BrwMain.Title.ForeColor = vbBlack
    BrwMain.Title.Font.Bold = True
        
    Set Cond = BrwMain.Conditions.Add("Anagrafica", "Anagrafica", m_DocType.TableName, True, False, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("CodiceSocio", "Codice socio", m_DocType.TableName, True, False, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("NumeroDocumento", "Numero documento", m_DocType.TableName, False, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("DataConferimento", "Data consegna merce", m_DocType.TableName, False, True, , dgCondTypeDate)
        Cond.RangeChecked = True
        Cond.FromValue = DateAdd("m", -3, Date)
        Cond.ToValue = Date
    Set Cond = BrwMain.Conditions.Add("DataDocumento", "Data comp. Liq.", m_DocType.TableName, False, True, , dgCondTypeDate)
    Set Cond = BrwMain.Conditions.Add("Indirizzo", "Indirizzo", m_DocType.TableName, False, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("Comune", "Comune", m_DocType.TableName, True, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("Cap", "C.A.P.", m_DocType.TableName, False, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("Provincia", "Provincia", m_DocType.TableName, True, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("NumeroDocumentoSocio", "Numero conf. socio", m_DocType.TableName, False, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("AnagraficaPerFatturazione", "Anagrafica di fatturazione", m_DocType.TableName, True, False, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("IDVettore", "Vettore", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.Indentation = 20
        Cond.RecordSource = "SELECT * FROM Vettore"
        Cond.DisplayField = "Vettore"
        Cond.KeyField = "IDVettore"
    Set Cond = BrwMain.Conditions.Add("TargaAutomezzo", "Targa automezzo", m_DocType.TableName, False, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("IDLuogoPresaMerce", "Sede stoccaggio merce", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.Indentation = 20
        Cond.RecordSource = "SELECT * FROM SitoPerAnagrafica WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        Cond.DisplayField = "SitoPerAnagrafica"
        Cond.KeyField = "IDSitoPerAnagrafica"
    Set Cond = BrwMain.Conditions.Add("IDTipoDocumentoAcq", "Tipo documento Acq.", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.Indentation = 20
        Cond.RecordSource = "SELECT * FROM RV_POTipoDocumentoAcq"
        Cond.DisplayField = "TipoDocumentoAcq"
        Cond.KeyField = "IDRV_POTipoDocumentoAcq"
    Set Cond = BrwMain.Conditions.Add("NumeroDocumentoAcq", "Numero documento Acquisto", m_DocType.TableName, False, True, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("DataDocumentoAcq", "Data documento acquisto", m_DocType.TableName, False, True, , dgCondTypeDate)
    Set Cond = BrwMain.Conditions.Add("IDRV_POTipoOrdineConforme", "Tipo ordine conforme", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.Indentation = 20
        Cond.RecordSource = "SELECT * FROM RV_POTipoOrdineConforme ORDER BY TipoOrdineConforme"
        Cond.DisplayField = "TipoOrdineConforme"
        Cond.KeyField = "IDRV_POTipoOrdineConforme"
    Set Cond = BrwMain.Conditions.Add("GeneratoDaPreConferimento", "Generato da pre-conferimento", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
    Set Cond = BrwMain.Conditions.Add("TrattaComeConferimento", "Tratta come conferimento", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
    
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
    'm_Report.Copies = m_Report.Copies
    'm_Report.Orientation = m_Report.Orientation
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
    
    
'    If Len(m_App.Caller) > 0 Then
'
'        If m_App.Caller = "RV_POControlloLavorazione" Then Exit Function
'        If m_App.Caller = "RV_POControlloVendite" Then Exit Function
'        If m_App.Caller = "RV_POMenuGreenTop" Then Exit Function
'
'        'Il programma è stato chiamato da un link.
'
'        'Se non verrà correttamente selezionato un elemento sarà restituito il valore -1 all'applicazione client.
'        lIDField = -1
'
'        'Se il documento è vuoto non si deve far nulla.
'        'Se la browse è in modalità Filter Definition non formula la domanda di riporto dei dati nel programma chiamante.
'        If (Not (m_Document.EOF And m_Document.BOF)) And (BrwMain.GuiMode <> dgFilterDefinition) Then
'
'            'ATTENZIONE: La stringa sMessage1 deve essere personalizzata a seconda dei casi!!!
'            sMessage1 = " il " & m_DocType.Name
'            sMessage = sMessage1 & " """ & m_Document.Fields(CAMPO_PER_CAPTION).Value & """"
'
'            gResource.CustomStrings.Clear
'            gResource.CustomStrings.Add sMessage, 1
'
'            'Viene chiesto se si intende riportare il record corrente al programma chiamante.
'            If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYPASTE), m_App.FunctionName) = vbYes Then
'                'Legge l'ID del record corrente affinchè venga riportato all'applicazione chiamante.
'                lIDField = m_Document.Fields("ID" & m_App.TableName).Value
'            End If
'
'        End If
'
'        'Scrive sul registry l'ID da passare all'aplicazione chiamante.
'        SaveSetting REGISTRY_KEY, m_App.Caller, "IDField", lIDField
'
'    End If
'
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

        If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
            'Il record è bloccato - si va in modalità tabellare
            
            BrwMain.Visible = True

            'Input Focus al browser
            'BrwMain.SetFocus

            'Refresh dello stato dei bottoni della ToolBar standard e dei menu
            SetStatus4Modality Browse

            Exit Sub
        Else
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
        End If

    End If
    
    
    
    'Se si era in fase di immissione di un nuovo record viene annullata
    m_Document.AbortNew
    
    If BrwMain.Visible Then 'Modalità tabellare
        
        'Input Focus al browser
        'BrwMain.SetFocus
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality Browse
        
    Else 'Modalità form
        
        'Refresh dello stato dei bottoni della ToolBar standard e dei menu
        SetStatus4Modality Modify
        
        'Input Focus al primo campo del form
        SetFocusTabIndex0
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
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "ControlloCollegamenti"
    BarMenu.Bands("StandardPO").Tools("ControlloCollegamenti").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("ControlloCollegamenti").SetPicture 0, gResource.GetBitmap(IDB_KIT16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("ControlloCollegamenti").ToolTipText = "Controllo collegamenti" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("ControlloCollegamenti").Description = "Controllo collegamenti"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("ControlloCollegamenti").Caption = "Controllo collegamenti"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep11"
    BarMenu.Bands("StandardPO").Tools("Sep11").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Tour"
    BarMenu.Bands("StandardPO").Tools("Tour").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("Tour").SetPicture 0, gResource.GetBitmap(IDB_COLLCOMPLETI16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("Tour").ToolTipText = "Tour" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("Tour").Description = "Tour"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("Tour").Caption = "Tour"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep12"
    BarMenu.Bands("StandardPO").Tools("Sep12").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "CreaDocumento"
    BarMenu.Bands("StandardPO").Tools("CreaDocumento").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("CreaDocumento").SetPicture 0, gResource.GetBitmap(IDB_CONFIG_REFRESH16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("CreaDocumento").ToolTipText = "Crea documento" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("CreaDocumento").Description = "Crea documento"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("CreaDocumento").Caption = "Crea documento"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep13"
    BarMenu.Bands("StandardPO").Tools("Sep13").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "CalcoloMedi"
    BarMenu.Bands("StandardPO").Tools("CalcoloMedi").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("CalcoloMedi").SetPicture 0, gResource.GetBitmap(IDB_LIFOFIFO16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("CalcoloMedi").ToolTipText = "Prezzi medi" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("CalcoloMedi").Description = "Prezzi  medi"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("CalcoloMedi").Caption = "Prezzi medi"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep14"
    BarMenu.Bands("StandardPO").Tools("Sep14").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Annotazioni"
    BarMenu.Bands("StandardPO").Tools("Annotazioni").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("Annotazioni").SetPicture 0, gResource.GetBitmap(IDB_EXPORT_16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("Annotazioni").ToolTipText = "Annotazioni" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("Annotazioni").Description = "Annotazioni"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("Annotazioni").Caption = "Annotazioni"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep15"
    BarMenu.Bands("StandardPO").Tools("Sep15").ControlType = ddTTSeparator
    
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "ChiudiLotto"
    BarMenu.Bands("StandardPO").Tools("ChiudiLotto").Style = ddSIconText
    BarMenu.Bands("StandardPO").Tools("ChiudiLotto").SetPicture 0, gResource.GetBitmap(IDB_BOLT16), &HC0C0C0
    BarMenu.Bands("StandardPO").Tools("ChiudiLotto").ToolTipText = "Chiudi lotto" '"GetCaption4MenuBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("ChiudiLotto").Description = "Chiudi lotto"  'GetDescription4StatusBar("Mnu_FormView")
    BarMenu.Bands("StandardPO").Tools("ChiudiLotto").Caption = "Chiudi lotto"  'GetDescription4StatusBar("Mnu_FormView")
    ToolID = ToolID + 1
    BarMenu.Bands("StandardPO").Tools.Add ToolID, "Sep16"
    BarMenu.Bands("StandardPO").Tools("Sep16").ControlType = ddTTSeparator
    
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
        gResource.CustomStrings.Add Chr(34) & m_App.FunctionName & Chr(34), 1

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

    sRecord = IIf(m_Document.Fields(CAMPO_PER_CAPTION).Value <> Empty, m_Document.Fields(CAMPO_PER_CAPTION).Value, TheApp.FunctionName)
  
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
    SetStatus4Modality Find
    
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
    
    'rif13
    
    'Set BrwMain.Recordset = m_Document.Dataset.Recordset
    Set BrwMain.Recordset = m_Document.Data
    
    
    
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
    
    m_DocType.Fields("IDFiliale").Value = m_App.Branch
    m_DocType.Fields("IDAzienda").Value = m_App.IDFirm
    m_DocType.Fields("IDUtente").Value = TheApp.IDUser
    m_DocType.Fields("IDTipoDocumentoCoop").Value = 2
    m_DocType.Fields("PreConferimento").Value = 0
    
    'Comunica all'oggetto DocType i valori da usare per la ricerca
    For Each Cond In BrwMain.Conditions
        
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
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
            Case dgCondTypeNumber
                If Cond.RangeChecked = True Then
                    m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
            
            Case dgCondTypeDate
                If Cond.RangeChecked = True Then
                    m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
            
            Case dgCondTypeTime
                If Cond.RangeChecked = True Then
                    m_DocType.Fields(Cond.FieldName).Value = Array(Cond.FromValue, Cond.ToValue)
                Else
                    If Len(Cond.FromValue) > 0 Then
                        m_DocType.Fields(Cond.FieldName).Value = BrwMain.Conditions(Cond.FieldName).FromValue
                    End If
                End If
          
            'Altre condizioni
            Case Else
                m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
  
        End Select
       
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
'Nome: SetFocusTabIndex0
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità:
'Da l'input focus al campo con TabIndex uguale a 0.
'**/
Private Sub SetFocusTabIndex0()
    On Error GoTo SetFocusTabIndex0_Error
    
    Dim ControlObject As Control
    Dim iIndex As Long
    Dim bError As Boolean
    
    If m_ControlTabIndex0 Is Nothing Then
        For Each ControlObject In frmMain.Controls
            iIndex = ControlObject.TabIndex
            If bError Then
                '**+ Controllo corrente non ha proprietà TabIndex,
                '    quindi va saltato.
                bError = False
            Else
                If ControlObject.TabIndex = 0 Then
                    Set m_ControlTabIndex0 = ControlObject
                    Exit For
                End If
            End If
        Next
    End If
    m_ControlTabIndex0.SetFocus

    Exit Sub
SetFocusTabIndex0_Error:
    bError = True
    Resume Next
End Sub

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
    
    'rif5 begin
    
    IsFieldInput = IsFieldInput Or TypeName(Control) = "DMTCombo"
    IsFieldInput = IsFieldInput Or TypeName(Control) = "Town"
    
    'rif5 end
    
    
End Function

'**+
'Nome: FieldPresent
'
'Parametri:
'Name - Nome del campo
'
'Valori di ritorno:
'Se il campo specificato nel parametro è presente nella
'collezione FormFields torna vero altrimenti torna falso
'
'Funzionalità:
'Controlla la presenza di un campo nella collezione FormFields
'**/
Private Function FieldPresent(ByVal Name As String) As Boolean
    Dim Field As FormField

    For Each Field In m_FormFields
        FieldPresent = (Name = Field.Name)
        If FieldPresent Then Exit For
    Next
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
Dim sSQL As String
  
'SETTARE LE GRIGLIE DEI SOTTODOCUMENTI
    Dim cl As dmtgridctl.dgColumnHeader

    'Inizializzazione della griglia adibita alla visualizzazione tabellare dei sotto-documenti
    '-------------------------------------------------------------------------------
       
    
   
    If Me.GrigliaCorpo.ColumnsHeader.Count = 0 Then
        With Me.GrigliaCorpo.ColumnsHeader
            .Add "CodiceArticolo", "Cod. Articolo", dgchar, True, 2000, 0, True, True, False
            .Add "Articolo", "Articolo", dgchar, True, 2000, 0, True, True, False
            .Add "UnitaDiMisura", "U.M", dgchar, True, 1000, 0, True, True, False
            .Add "PrezzoMedio", "Prezzo medio", dgBoolean, False, 1000, dgAligncenter
            .Add "TipoLavorazione", "Tipo lavorazione", dgchar, False, 1800, dgAlignleft
            
            
            Set cl = .Add("Colli", "Colli", dgDouble, True, 1000, 0, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("PesoLordo", "Peso Lordo", dgDouble, True, 1000, 0, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("PesoNetto", "Peso netto", dgDouble, True, 1000, 0, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("Tara", "Tara", dgDouble, True, 1000, 0, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("Pezzi", "Pezzi", dgDouble, True, 1000, 0, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("Qta_UM", "Q.tà", dgDouble, True, 1000, 0, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "CodiceImballo", "Cod. Imballo", dgchar, True, 2000, 0, True, True, False
            .Add "DescrizioneImballo", "Imballo", dgchar, True, 2000, 0, True, True, False
            .Add "TipoConfLiquidazione", "Stato Liquidazione", dgchar, False, 1800, dgAlignleft
                    
        End With
    End If
    Me.GrigliaCorpo.EnableMove = True
    


    If Me.GrigliaImballi.ColumnsHeader.Count = 0 Then
        With Me.GrigliaImballi.ColumnsHeader
            .Add "IDRV_POCaricoMerceImballi", "IDRV_POCaricoMerceImballi", dgInteger, False, 500, dgAlignRight
            .Add "IDRV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", dgInteger, False, 500, dgAlignRight
            .Add "CodiceArticolo", "Cod. Articolo", dgchar, True, 2000, 0, True, True, False
            .Add "Articolo", "Articolo", dgchar, True, 2000, 0, True, True, False
            .Add "IDRV_POTipoProcessoCoop", "IDRV_POTipoProcessoCoop", dgInteger, False, 500, dgAlignRight
            .Add "TipoProcessoCoop", "Tipo processo", dgchar, True, 1500, 0, True, True, False
            Set cl = .Add("Quantita", "Quantita", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "RiferimentoEsterno", "Rif. esterno", dgchar, True, 2000, 0, True, True, False
        End With
    End If
    Me.GrigliaImballi.EnableMove = True


    If Me.GrigliaAddebiti.ColumnsHeader.Count = 0 Then
        With Me.GrigliaAddebiti.ColumnsHeader
            .Add "IDRV_POCaricoMerceAddebiti", "IDRV_POCaricoMerceAddebiti", dgInteger, False, 500, dgAlignRight
            .Add "IDRV_POCaricoMerceTesta", "IDRV_POCaricoMerceTesta", dgInteger, False, 500, dgAlignRight
            .Add "IDArticolo", "IDArticolo", dgInteger, False, 500, dgAlignRight
            .Add "CodiceArticolo", "Cod. Articolo", dgchar, True, 2000, 0, True, True, False
            .Add "Articolo", "Articolo", dgchar, True, 2000, 0, True, True, False
            Set cl = .Add("Quantita", "Quantita", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("ImportoUnitario", "Importo unitario", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "IDRV_POTipoTrattenutaAggiuntiva", "IDRV_POTipoTrattenutaAggiuntiva", dgInteger, False, 500, dgAlignRight
            .Add "TipoTrattenuta", "Tipo trattenuta aggiuntiva", dgchar, True, 1500, 0, True, True, False
            .Add "IDIva", "IDIva", dgInteger, False, 500, dgAlignRight
            .Add "Iva", "Iva", dgchar, True, 2000, 0, True, True, False
            Set cl = .Add("AliquotaIva", "Aliquota", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("TotaleRigaNettoIva", "Imponibile riga", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("ImpostaRiga", "Imposta riga", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("TotaleRigaLordoIva", "Totale riga", dgDouble, True, 1000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
        End With
    End If
    Me.GrigliaAddebiti.EnableMove = True

'''''''''''''''''''''''''CONTROLLI STANDARD''''''''''''''''''''''''''''''''''''
    'Anagrafica socio
    With Me.CDSocio
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepFornitore"
        .Filter = "IDAzienda = " & m_App.IDFirm
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Anagrafica"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = 29 'Anagrafica
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    
    'Magazzino di conferimento
    With Me.cboMagazzinoConf
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .SQL = "SELECT * FROM Magazzino WHERE IDFiliale=" & m_App.Branch
        .Fill
    End With
    
    'Magazzino di vendita
    With Me.CboMagazzinoVend
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .SQL = "SELECT * FROM Magazzino WHERE IDFiliale=" & m_App.Branch
        .Fill
    End With
    
    'Sezionale
    With Me.cboSezionale
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT RV_POSezionalePerDocumento.IDSezionale, Sezionale.Sezionale, RV_POSchemaCoop.IDFiliale, RV_POSchemaCoop.IDUtente, "
        .SQL = .SQL & "RV_POSezionalePerDocumento.IDDocumentoCoop "
        .SQL = .SQL & "FROM RV_POSchemaCoop INNER JOIN "
        .SQL = .SQL & "RV_POSezionalePerDocumento ON "
        .SQL = .SQL & "RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POSezionalePerDocumento.IDRV_POSchemaCoop INNER JOIN "
        .SQL = .SQL & "Sezionale ON RV_POSezionalePerDocumento.IDSezionale = dbo.Sezionale.IDSezionale "
        .SQL = .SQL & "WHERE (RV_POSchemaCoop.IDFiliale = " & m_App.Branch & ") And (dbo.RV_POSchemaCoop.IDUtente = 0) And (dbo.RV_POSezionalePerDocumento.IDDocumentoCoop = " & IDDocumento & ")"
        .Fill
    End With

    'Esercizio
    With Me.cboEsercizio
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDEsercizio"
        .DisplayField = "Esercizio"
        .SQL = "SELECT * FROM Esercizio WHERE IDAzienda=" & m_App.IDFirm
        .Fill
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
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
    
    'Imballo
    With Me.CDCodiceImballo
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
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
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
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice Articolo"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione Articolo"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With

    'Unita di misura
    With Me.cboUM
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With

    'Unita di misura imballo
    With Me.cboUMImballo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With

    'Unita di misura Pedana
    With Me.cboUMPedana
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With
    'Regione
    With Me.cboRegione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRegione"
        .DisplayField = "Regione"
        .SQL = "SELECT * FROM Regione ORDER BY Regione"
        .Fill
    End With

    'Nazione
    With Me.cboNazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDNazione"
        .DisplayField = "Nazione"
        .SQL = "SELECT * FROM Nazione ORDER BY Nazione"
        .Fill
    End With

    'Comune
    With Me.cboComune
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDComune"
        .DisplayField = "Comune"
        .SQL = "SELECT * FROM Comune ORDER BY Comune"
        .Fill
    End With

    'Provincia
    With Me.cboProvincia
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDProvincia"
        .DisplayField = "Provincia"
        .SQL = "SELECT * FROM Provincia ORDER BY Provincia"
        .Fill
    End With

    'Utente
    With Me.cboUtente
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente ORDER BY Utente"
        .Fill
    End With

    With Me.cboIVA
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "IVa"
        .SQL = "SELECT * FROM Iva"
        .Fill
    End With


    'Identificativo del tipo documento di Acquisto
    With Me.cboTipoDocumentoAcq
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoDocumentoAcq"
        .DisplayField = "TipoDocumentoAcq"
        .SQL = "SELECT * FROM RV_POTipoDocumentoAcq"
        .Fill
    End With
    
    'Inizializzazione della LabelLink
    '----------------------------
    Set LabelLink1.Application = TheApp    'L’oggetto Application
    LabelLink1.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    LabelLink1.IDFunction = GET_FUNZIONE("RV_POLavorazioneL")
    
    'Viene disabilitata la voce "Ricerca" del menu popup
    LabelLink1.PopMenuItems("Mnu_SearchObject").Enabled = False



''''''GESTIONE ALTRI IMBALLI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Imballo
    With Me.CDImballoGestione
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

    'Unita di misura
    With Me.cboUMImbGest
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDUnitaDiMisura"
        .DisplayField = "UnitaDiMisura"
        .SQL = "SELECT * FROM UnitaDiMisura"
        .Fill
    End With

    'Identificativo del tipo di gestione dell'imballo
    With Me.cboTipoProcessoImballo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoProcessoCoop"
        .DisplayField = "TipoProcessoCoop"
        .SQL = "SELECT * FROM RV_POTipoProcessoCoop"
        .Fill
    End With

'''''''''GESTIONE ALTRI ADDEBITI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Articolo per altri addebiti
    With Me.CDArticoliAddebiti
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

    With Me.cboIvaAddebiti
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "IVa"
        .SQL = "SELECT * FROM Iva"
        .Fill
    End With

    With Me.cboTipoTrattenutaAggiuntiva
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoTrattenutaAggiuntiva"
        .DisplayField = "TipoTrattenuta"
        .SQL = "SELECT * FROM RV_POTipoTrattenutaAggiuntiva"
        .Fill
    End With

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

    With Me.cboVettore
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDVettore"
        .DisplayField = "Vettore"
        .SQL = "SELECT IDVettore, Vettore FROM Vettore"
        .SQL = .SQL & " ORDER BY Vettore"
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

    'Inizializzazione della LabelLink
    '----------------------------
    Set Me.LinkLavorazione.Application = TheApp    'L’oggetto Application
    Me.LinkLavorazione.WindowHandleClient = Me.hwnd   'L’Handle del Form che contiene la LabelLink
    Me.LinkLavorazione.IDFunction = GET_FUNZIONE("RV_POAssegnazioneMerce")
    
    'Viene disabilitata la voce "Ricerca" del menu popup
    Me.LinkLavorazione.PopMenuItems("Mnu_SearchObject").Enabled = False


    With Me.cboTipoLavorazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .SQL = "SELECT * FROM RV_POTipoLavorazione ORDER BY TipoLavorazione"
        .Fill
    End With

    With Me.cboLiquidato
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoConfLiquidazione"
        .DisplayField = "TipoConfLiquidazione"
        .SQL = "SELECT * FROM RV_POTipoConfLiquidazione"
        .Fill
    End With
    With Me.cboTipoOrdConf
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoOrdineConforme"
        .DisplayField = "TipoOrdineConforme"
        .SQL = "SELECT IDRV_POTipoOrdineConforme, TipoOrdineConforme FROM RV_POTipoOrdineConforme"
        .SQL = .SQL & " ORDER BY TipoOrdineConforme"
    End With
    
    If LINK_TIPO_LIQ_CONF <= 1 Then
        Me.cboLiquidato.Visible = False
        Me.Label2(37).Visible = False
        Me.cboTipoLavorazione.Width = 4815
    Else
        Me.cboLiquidato.Visible = True
        Me.Label2(37).Visible = True
        Me.cboTipoLavorazione.Width = 3225
    End If
    
    'Cliente per ordine
     With Me.cdCliente
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "RV_POIETipoAnagraficaCliente"
        .Filter = "IDAzienda = " & m_App.IDFirm
        '.MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Clienti"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice"
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Anagrafica"
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("Anagrafica") 'Articoli
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With
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
Dim Field As DmtDocManLib.Field
Dim DocLink As DmtDocManLib.DocumentsLink
Dim sSQL As String
Dim NumeroInserimenti As Long
Dim OLD_CURSOR As Long
Dim Testo As String
Dim AGGIORNA_DATA_CONFERIMENTO As Boolean
Dim NuovoDocumento As Boolean
Dim X As Long
Dim ErroreCoda As Boolean

SALVA_DOC_OK = 0
If CHECK_ABILITAZIONE_DIAMANTE = False Then Exit Sub

If MODULO_ATTIVATO = 0 Then
    If Len(MODULO_DESCRIZIONE) > 0 Then
        MsgBox "Il modulo " & MODULO_DESCRIZIONE & " non è stato abilitato", vbInformation, TheApp.FunctionName
    Else
        MsgBox "Questa funzionalità non può essere avviata senza abilitazione", vbInformation, TheApp.FunctionName
    End If
Exit Sub
End If

    AGGIORNA_DATA_CONFERIMENTO = False
    
    AggiornamentoDocumento = 1
 
    If Not PermissionToSave Then
        Exit Sub
    End If
        
    If m_Document(m_Document.PrimaryKey).Value > 0 Then
        If fnNotNull(m_Document("DataDocumento").Value) <> Me.txtDataDocumento.Text Then
            If GET_CONTROLLO_COLLEGAMENTI_CONFERIMENTO(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = True Then
                Testo = "ATTENZIONE!!!" & vbCrLf
                Testo = Testo & "Alcune righe di conferimento risultano aperte in altre postazioni" & vbCrLf
                Testo = Testo & "Vuoi visualizzare il dettaglio?"
                If MsgBox(Testo, vbQuestion + vbYesNo, "Impossibile salvare") = vbYes Then
                    LINK_TESTA_DOCUMENTO = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
                    frmCollegamentiConferimento.Show vbModal
                End If
                
                Exit Sub
            Else
                AGGIORNA_DATA_CONFERIMENTO = True
            End If
        End If
    End If

    
    
    SCRIVI_CODA fnNotNullN(m_Document(m_Document.PrimaryKey))
    
    APERTURA_FORM_CODA = False
    
    For Each Field In m_Document.Fields
        'Sul campo chiave primaria non si deve far nulla
        If Not Field.PrimaryKey Then
            If FieldPresent(Field.Name) Then
               
                'rif4 begin

                Select Case TypeName(m_FormFields(Field.Name).Control)
                    Case "TextBox"
                        Field.Value = m_FormFields(Field.Name).Control.Text
                    Case "DmtCodDesc"
                        Field.Value = m_FormFields(Field.Name).Control.KeyFieldID
                    Case "DMTCombo"
                        Field.Value = m_FormFields(Field.Name).Control.CurrentID
                    Case "Town"
                        If Field.Name = "IDComune" Then
                            Field.Value = m_FormFields(Field.Name).Control.CityID
                        ElseIf Field.Name = "Cap" Then
                            Field.Value = m_FormFields(Field.Name).Control.Zip
                        End If
                    Case "dmtDate"
                        If (m_FormFields(Field.Name).Control.Text = "") Or (IsNull(m_FormFields(Field.Name).Control.Value)) Then
                            Field.Value = Null
                        Else
                            Field.Value = m_FormFields(Field.Name).Control.Value
                        End If
                    Case "dmtNumber"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    Case "dmtCurrency"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    
                    Case "dmtTime"
                        Field.Value = m_FormFields(Field.Name).Control.Value
                    Case "DmtSearchACS"
                        Field.Value = m_FormFields(Field.Name).Control.IDAnagrafica
                    Case "CheckBox"
                        Field.Value = fnNormBoolean(m_FormFields(Field.Name).Control.Value)
                    Case "Label"
                        Field.Value = m_FormFields(Field.Name).Control.Caption
                
                End Select
                
                'rif4 end
                
            Else
                If Field.Name = "IDAzienda" Then
                    Field.Value = m_App.IDFirm
                End If
                
                If Field.Name = "IDFiliale" Then
                    Field.Value = m_App.Branch
                End If
                'If Field.Name = "IDEsercizio" Then
                '    Field.Value = Link_Esercizio
                'End If
                If Field.Name = "IDTipoDocumentoCoop" Then
                    Field.Value = 2
                End If
                If Field.Name = "CodiceSocio" Then
                    Field.Value = Me.CDSocio.Code
                End If
                If Field.Name = "Anagrafica" Then
                    Field.Value = Me.CDSocio.Description
                End If
                If Field.Name = "Nome" Then
                    Field.Value = Me.txtNomeSocio.Text
                End If

                If Field.Name = "IDTipoOggetto" Then
                    Field.Value = fnGetTipoOggetto
                End If
                If Field.Name = "IDOggetto" Then
                    If fnNotNullN(m_Document("IDOggetto").Value) = 0 Then
                        Field.Value = GET_LINK_OGGETTO
                    End If
                End If
                If Field.Name = "PreConferimento" Then
                    Field.Value = 0
                End If

                'Se il processo in corso è "Manutenzione da Shell"
                'la variabile m_LinkedField contiene il nome del
                'campo collegato alla applicazione chiamante
                'quindi il campo relativo deve essere valorizzato
                'con il valore ricevuto dalla applicazione chiamante
                'nCBC+
                If Field.Name = m_LinkedField Then
                    Field.Value = m_App.CallerFieldValue
                End If
            End If
        End If
    Next
 
If m_Document(m_Document.PrimaryKey).Value < 0 Then
    NuovoDocumento = True
End If

''''''''''''''''''''''''''''''CONTROLLA LA CODA DEI SALVATAGGI'''''''''''''''''''''''''''''
X = 0
ErroreCoda = False
Do
    X = GET_NUMERO_DOCUMENTO(NuovoDocumento)
    If X = -1 Then
        X = 1
        ErroreCoda = True
    End If
Loop Until X = 1

If ErroreCoda = True Then
    X = -1
End If

If X = -1 Then
    Me.Enabled = True
    Me.SetFocus
    Me.Caption = Caption2Display
    Screen.MousePointer = 0
    ''''''''ELIMINAZIONE RIFERIMENTO CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Me.Enabled = True
Me.SetFocus
Me.Caption = Caption2Display

OLD_CURSOR = Cn.CursorLocation
Cn.CursorLocation = adUseClient


frmAttesa.Show
Me.Enabled = False

DoEvents

Me.Caption = "SALVATAGGIO IN CORSO..................."
DoEvents

frmAttesa.lblInfo = Me.Caption
DoEvents

Cn.BeginTrans

m_Document.SaveDocument

Cn.CommitTrans

Me.Caption = "TRACCIABILITA' IMBALLI IN CORSO..........."
DoEvents

frmAttesa.lblInfo = Me.Caption
DoEvents

CREA_RIGHE_TRACCIABILITA_IMBALLO fnNotNullN(m_Document(m_Document.PrimaryKey).Value), m_DocType.ID, Me.CDSocio.KeyFieldID, Me.txtDataConferimento.Text, Me.txtNumeroDocumento.Text, Me.CDSocio.Code
DELETE_LOTTO_IMBALLO

Me.Caption = "MOVIMENTAZIONE IN CORSO..................."
DoEvents

frmAttesa.lblInfo = Me.Caption
DoEvents

If SalvataggioRighe(m_Document(m_Document.PrimaryKey).Value) = False Then
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Sono stati riscontrati errori nella procedura di movimentazione delle righe del documento." & vbCrLf
    Testo = Testo & "Vuoi tentare di salvare di nuovo il documento?"
    
    If MsgBox(Testo, vbQuestion + vbYesNo, "Errore di salvataggio delle righe") = vbYes Then
        ''''''''ELIMINAZIONE RIFERIMENTO CODA'''''''''''''''''''''''''''''''
        sSQL = "DELETE FROM RV_POTMP "
        sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
        'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
        Cn.Execute sSQL
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Cn.CursorLocation = OLD_CURSOR
        
        Me.Caption = Caption2Display(False)
        
        AggiornamentoDocumento = 0
            
        OnSave
    End If
End If
Unload frmAttesa
Me.Enabled = True
Me.SetFocus

sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
Cn.Execute sSQL

AGGIORNA_DATA_DI_LIQUIDAZIONE fnNotNullN(m_Document(m_Document.PrimaryKey).Value), Me.txtDataDocumento.Text

Cn.CursorLocation = OLD_CURSOR
   
m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions

'Refresh delle variabili di stato
m_Changed = False
m_Search = False
m_Saved = True

'Refresh dello stato della ToolBar standard in modalità variazione
SetStatus4Modality Modify

AggiornamentoDocumento = 0

sSQL = "DELETE FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & m_Document(m_Document.PrimaryKey).Value
sSQL = sSQL & " AND IDArticolo IS NULL"

Cn.Execute sSQL


m_DocumentsLink.Refresh

DELETE_PESATURE

ADD_PESATURE_CONFERIMENTO fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

CONTROLLO_MOVIMENTAZIONE_CONFERIMENTO fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

SALVA_DOC_OK = 1

If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
    m_DocumentsLink.Move 0
End If
    
Exit Sub
ERR_OnSave:
    Unload frmAttesa
    Me.Enabled = True
    Me.SetFocus
    
    MsgBox Err.Description, vbCritical, "OnSave"

    Cn.RollbackTrans
    ''''''''''''''''''''ELIMINAZIONE RIGA DI CODA'''''''''''''''''''''''''''''''
    sSQL = "DELETE FROM RV_POTMP "
    sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
    'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID
    Cn.Execute sSQL
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Cn.CursorLocation = OLD_CURSOR
    
    Me.Caption = Caption2Display(False)
    
    AggiornamentoDocumento = 0

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
Dim Link_Oggetto As Long
Dim Testo As String
Dim sSQL As String
Dim AvvioFormAttesa As Boolean

    AvvioFormAttesa = False
    'Se si è in modalità tabellare potrebbe essere necessario sincronizzare
    'il documento con il record evidenziato nella browse
    
    If BrwMain.Visible = True Then
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            m_Document.Move BrwMain.ListIndex - 1
        End If
    End If
    
    
    If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemDeleteAction) Then
        Exit Sub
    End If
    
    'Se in fase di inserimento di un nuovo
    'record non c'è niente da fare
    If m_Document.TableNew Then
        Exit Sub
    End If
    
    'Conferma della cancellazione
    gResource.CustomStrings.Clear
    sToRemove = m_Document.Fields(CAMPO_PER_CAPTION).Value
    gResource.CustomStrings.Add Chr(34) & sToRemove & Chr(34), 1
    If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYREMOVE), m_App.FunctionName) = vbYes Then
        
        If ControlloRigheMovimentate = False Then
            MsgBox "Impossibile eliminare il documento poichè alcuni lotti del documento risultano movimentati da altri documenti", vbCritical, "Impossibile eliminare"
            Exit Sub
        End If
        
        If ControlloRigheMovimentateCampionatura = False Then
            MsgBox "Impossibile eliminare il documento poichè alcuni lotti del documento risultano movimentati nella campionatura", vbCritical, "Impossibile eliminare"
            Exit Sub
        End If
        
        If ControlloRigheMovimentateLottiImballo = True Then
            MsgBox "Impossibile eliminare il documento poichè alcuni lotti imballi del documento risultano movimentati nella campionatura", vbCritical, "Impossibile eliminare"
            Exit Sub
        End If
        
        
        Screen.MousePointer = 11
        frmAttesa.Show
        Me.Enabled = False
        frmAttesa.lblInfo.Caption = "ELIMINAZIONE DEL DOCUMENTO......"
        DoEvents
        AvvioFormAttesa = True
        
        Link_Oggetto = fnNotNullN(m_Document("IDOggetto").Value)
        
        frmAttesa.lblInfo.Caption = "ELIMINAZIONE MOVIMENTI MERCE......"
        DoEvents
        
        EliminaRigheDaDocumento
        
        frmAttesa.lblInfo.Caption = "ELIMINAZIONE MOVIMENTI IMBALLI......"
        DoEvents
        EliminaRigheDaDocumentoImballi fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        
        frmAttesa.lblInfo.Caption = "RECUPERO LOTTI IMBALLI....."
        DoEvents
        LOTTI_IMBALLI_PER_ELIMINAZIONE fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        frmAttesa.lblInfo.Caption = "ELIMINAZIONE LOTTI IMBALLI....."
        DoEvents
        DELETE_LOTTO_IMBALLO
        
        If Not (m_Document.EOF Or m_Document.BOF) Then
            'Cancella l'eventuale blocco sul record da cancellare.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        End If

        frmAttesa.lblInfo.Caption = "ELIMINAZIONE DOCUMENTO......"
        DoEvents
        
        m_Document.DeleteDocument
        
        'COLLEGAMENTO TOUR''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        frmAttesa.lblInfo.Caption = "ELIMINAZIONE COLLEGAMENTO TOUR......"
        DoEvents
        
        GET_DATI_TOUR Link_Oggetto
        
        sSQL = "DELETE FROM RV_POTourRighe "
        sSQL = sSQL & "WHERE IDOggettoOrdine=" & Link_Oggetto
        Cn.Execute sSQL
        
        If LINK_TOUR > 0 Then
            GET_CAMBIO_POSIZIONE LINK_TOUR, POSIZIONE_TOUR
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                         

        
        frmAttesa.lblInfo.Caption = "ELIMINAZIONE OGGETTO......"
        DoEvents
        sSQL = "DELETE FROM Oggetto "
        sSQL = sSQL & "WHERE IDOggetto=" & Link_Oggetto
        Cn.Execute sSQL
        
        
        Unload frmAttesa
        Me.Enabled = True
        AvvioFormAttesa = False
        Screen.MousePointer = 0
            
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
                If Not m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
                'Il record è bloccato.
                'Va in modalità tabellare
                BrwMain.Visible = True
                SetStatus4Modality Browse
            Else
            'Il record non è bloccato.
            
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
            m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
            
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

End If
Exit Sub
ERR_OnDelete:
     Screen.MousePointer = 0
    If AvvioFormAttesa = True Then
        Unload frmAttesa
    End If


    MsgBox Err.Description, vbCritical, TheApp.FunctionName

    
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
        SetFocusTabIndex0
        
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
    
    GENERA_FILTRO_PER_TIPO_OGGETTO
    
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
Private Sub OnMoveCurrentRecord(ByVal Tipo As Integer, ByVal sToolName As String)
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
       Select Case Tipo
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
Private Sub OnPrint(ByVal ToolName As String)
    Dim lFlags As Long
    Dim OLDCursor As Integer
    Dim sStr As String
    Dim Field As DmtDocManLib.Field
    
    
    OLDCursor = Screen.MousePointer
    
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
    
    If Not BrwMain.Visible Then
        'Modalità Form - deve stampare solo il record corrente
        
        'SE IL FLAG NEI PARAMETRI E' ATTIVO BISOGNA ESEGUIRE LA PROCEDURA
        If Abs(CALCOLA_RIEP_IMBALLI_STAMPA) = 1 Then
            AVVIA_PROCEDURA fraImpAttesa, Cn, Me.CDSocio.KeyFieldID, fnGetEsercizio(Date), Me.ProgressBar2, Me.lblInfo2
        End If
        
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next
        
        'Viene inserita la condizione di ricerca basata sull'ID del record corrente.
        m_DocType.Fields("IDFiliale").Value = m_App.Branch
        m_DocType.Fields("IDAzienda").Value = m_App.IDFirm
        m_DocType.Fields("IDUtente").Value = TheApp.IDUser
        m_DocType.Fields("IDTipoDocumentoCoop").Value = 2
        m_DocType.Fields("ID" & m_App.TableName).Value = m_Document.Fields("ID" & m_App.TableName).Value
        
        'Viene creato un filtro temporaneo per il Crystals Reports.
        m_DocType.RemoveFilter "Form"
        Set m_Report.Filter = m_DocType.AddFilterWithConditions("Form")
    Else
        'Modalità vista tabellare
        
        'Viene passato il filtro corrente al Crystals Reports.
        Set m_Report.Filter = m_ActiveFilter
    End If
            
        
    
    Select Case ToolName
    
        Case "PrePrint", "Mnu_PrePrint"
            On Error GoTo ErrorHandler
            
            Screen.MousePointer = vbHourglass
            
            m_TabMode = BrwMain.Visible
            PicForm.Visible = False
            BrwMain.Visible = False
            ActivityBox.Visible = False
            
            SetStatus4Modality Preview, OpenPrw
            Refresh
            
            m_PreviewWindowHandle = m_Document.Preview(m_Report, "", hwnd, CInt(BarMenu.ClientAreaLeft / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaTop / Screen.TwipsPerPixelY), CInt(BarMenu.ClientAreaWidth / Screen.TwipsPerPixelX), CInt(BarMenu.ClientAreaHeight / Screen.TwipsPerPixelY), False)
            lFlags = SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOMOVE
            SetWindowPos m_PreviewWindowHandle, HWND_TOP, 0, 0, 0, 0, lFlags
            
        Case "Print", "Mnu_Print"
            PrintDocument ToolName
            
        Case "ExportWord", "Mnu_ExportWord"
            ExportDocument ecWord
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Word, TheApp.Name

        Case "ExportExcel", "Mnu_ExportExcel"
            ExportDocument ecExcel
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), Excel, TheApp.Name
            
        Case "ExportHtml", "Mnu_ExportHtml"
            ExportDocument ecHtml
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), HTML, TheApp.Name
        
        Case "ExportPDF", "Mnu_ExportPDF"
            ExportDocument ecPdf
            fnExtractNameFromTag BarMenu.Bands("Standard").Tools("Export"), PDF, TheApp.Name
        
        Case "MailWord"
            SendDocument ecWord
            
        Case "MailExcel"
            SendDocument ecExcel
            
        Case "MailHtml"
            SendDocument ecHtml
        
        Case "MailPDF"
            SendDocument ecPdf
    End Select
    
   
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




Private Function fncListinoDefault() As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "Select IDListinoDiBase From ConfigurazioneVendite Where IDAzienda=" & m_App.IDFirm
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fncListinoDefault = rs!IDListinoDiBase
    
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function





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
    oSearch.AddDisplayField "Codice", "Codice", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Fornitore.Codice "
    sSQL = sSQL & "FROM Anagrafica INNER JOIN "
    sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica "
    sSQL = sSQL & "WHERE Fornitore.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Anagrafica"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Anagrafica")
                
    End If
End If
If Name = "Codice fornitore" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Codice socio"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Codice", "Codice", 1   'STRINGTYPE
    oSearch.AddDisplayField "Socio\Fornitore", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1 'STRINGTYPE
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Anagrafica.IDAnagrafica, Fornitore.Codice, Anagrafica.Anagrafica, Anagrafica.Nome  "
    sSQL = sSQL & "FROM Anagrafica INNER JOIN "
    sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica "
    sSQL = sSQL & "WHERE Fornitore.IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " ORDER BY Codice"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Codice")
                
    End If
End If




If Name = "Comune" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Comune"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Comune" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Comune", "Comune", 1 'STRINGTYPE
    oSearch.AddDisplayField "Provincia", "Provincia", 1   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Comune.IDComune, Comune.Comune, Provincia.Provincia "
    sSQL = sSQL & "FROM Comune LEFT OUTER JOIN "
    sSQL = sSQL & "Provincia ON Comune.IDProvincia = Provincia.IDProvincia "
    sSQL = sSQL & "ORDER BY Comune"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Comune")
                
    End If
End If
If Name = "Provincia" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Provincia"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Comune" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    
    oSearch.AddDisplayField "Provincia", "Provincia", 1   'STRINGTYPE
    oSearch.AddDisplayField "Nome provincia", "NomeProvincia", 1   'STRINGTYPE
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Provincia.IDProvincia, Provincia.Provincia, Provincia.NomeProvincia "
    sSQL = sSQL & "FROM Provincia "
    sSQL = sSQL & "ORDER BY Provincia"


    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("Provincia")
                
    End If
End If
    Set oRes = Nothing
    Set oSearch = Nothing

End Sub

Private Sub cboComune_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDProvincia FROM Comune "
sSQL = sSQL & "WHERE IDComune=" & Me.cboComune.CurrentID


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboProvincia.WriteOn 0
Else
    Me.cboProvincia.WriteOn fnNotNullN(rs!IDProvincia)
End If


rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub cboEsercizio_Click()
On Error Resume Next
    Link_Esercizio = fnGetEsercizio(Me.txtDataConferimento.Text)
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        Me.txtNumeroConferimento.Value = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "NumeroConferimento")
        Me.txtPrefissoConferimento.Text = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "PrefissoNumeroConferimento")
    End If
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboIVA_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & Me.cboIVA.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Me.txtAliquotaIVA.Value = fnNotNullN(rs!AliquotaIva)
Else
    Me.txtAliquotaIVA.Value = 0
End If

rs.CloseResultset
Set rs = Nothing

fnTotaleRiga
End Sub

Private Sub fnTotaleRiga()
On Error GoTo ERR_fnTotaleRiga
Dim imponibileImballo As Double
    imponibileImballo = Me.txtColli.Value * Me.txtImpUniImb.Value
    
    Me.txtImponibileRiga.Value = (Me.txtQta_UM.Value * Me.txtImportoUnitario.Value) + imponibileImballo
    
    Me.txtImpostaRiga.Value = (Me.txtImponibileRiga.Value / 100) * Me.txtAliquotaIVA.Value
    
    Me.txtTotaleRiga.Value = Me.txtImponibileRiga.Value + Me.txtImpostaRiga.Value
Exit Sub
ERR_fnTotaleRiga:
    MsgBox Err.Description, vbCritical, "fnTotaleRiga"
End Sub

Private Sub cboIvaAddebiti_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & Me.cboIvaAddebiti.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Me.txtAliquotaIvaAddebiti.Value = fnNotNullN(rs!AliquotaIva)
Else
    Me.txtAliquotaIvaAddebiti.Value = 0
End If

rs.CloseResultset
Set rs = Nothing


GET_TOTALI_RIGA_ALTRI_ADDEBITI

End Sub



Private Sub cboLuogoPresaMerce_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboMagazzinoConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CboMagazzinoVend_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboSezionale_Click()
    Link_Sezionale = Me.cboSezionale.CurrentID
    Me.txtNumeroDocumento.Value = fnGetNumeroDocumento
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoDocumentoAcq_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoOrdConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoProcessoImballo_Click()
Me.txtGiacLottoImballo.Value = 0
Me.txtDispLottoImballo.Value = 0
Me.cmdLottoImballo.Enabled = False
Me.cmdRiepilogoLottoImballo.Enabled = False

If Me.chkTracciaImballoGest.Value = vbChecked Then
    If Me.cboTipoProcessoImballo.CurrentID = 2 Then
        Me.txtGiacLottoImballo.Value = GET_GIACENZA_LOTTO_IMBALLO(Me.CDImballoGestione.KeyFieldID)
        
        Me.txtDispLottoImballo.Value = Me.txtGiacLottoImballo.Value + GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(0, fnGetTipoOggetto("RV_POGestImbConf"), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
    
        If Me.txtIDLottoImballoGest.Value > 0 Then
            Me.txtDispLottoImballo.Value = Me.txtDispLottoImballo.Value - GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(Me.txtIDLottoImballoGest.Value, fnGetTipoOggetto("RV_POGestImbConf"), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
        Else
            Me.cmdLottoImballo.Enabled = True
            CREA_RECORDSET_LOTTI_IMBALLI Me.CDImballoGestione.KeyFieldID, fnGetTipoOggetto("RV_POGestImbConf"), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)
        End If
    End If
    If Me.cboTipoProcessoImballo.CurrentID = 1 Then
        If Me.txtIDLottoImballoGest.Value > 0 Then
            Me.cmdRiepilogoLottoImballo.Enabled = True
        End If
    End If
End If
End Sub

Private Sub cboUM_Click()
    Link_UnitaDiMisura_Coop = fnGetUMCoop(Me.cboUM.CurrentID)
End Sub

Private Sub cboVettore_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDArticoliAddebiti_ChangeElement()
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If fnNotNullN(m_DocumentsLink2(m_DocumentsLink2.PrimaryKey).Value) > 0 Then Exit Sub

''''TROVA IVA VENDITA ARTICOLO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDIvaVendita "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & Me.CDArticoliAddebiti.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboIvaAddebiti.WriteOn 0
Else
    Me.cboIvaAddebiti.WriteOn fnNotNullN(rs!IDIvaVendita)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Me.txtQtaAddebiti.Value = 1
Me.txtImpUniAddebiti.Value = GET_PREZZO_IMBALLO(Me.CDArticoliAddebiti.KeyFieldID)

GET_TOTALI_RIGA_ALTRI_ADDEBITI

End Sub

Private Sub CDArticolo_ChangeElement()
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Link_Lotto As Long
Dim IDIva As Long
Dim LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL As Long

     'If IsNull(m_DocumentsLink("IDArticolo").Value) Then
        If CDArticolo.KeyFieldID > 0 Then
            If ((IsNull(m_DocumentsLink("IDArticolo").Value)) Or (fnNotNullN(m_DocumentsLink("IDArticolo").Value) <> Me.CDArticolo.KeyFieldID)) Then
                If GET_ARTICOLO_ANNULLATO(Me.CDArticolo.KeyFieldID) = True Then
                    MsgBox "ATTENZIONE!!!" & vbCrLf & "L'articolo selezionato è stato bloccato e pertanto non può essere utilizzato", vbInformation, "Inserimento dati"
                    Me.TxtArticolo.Text = ""
                    Me.CDArticolo.Load 0
                Exit Sub
                End If
            End If
            
            Me.TxtArticolo.Text = Me.CDArticolo.Description
            
            sSQL = "SELECT IDUnitaDiMisuraAcquisto, IDUnitaDiMisuraVendita, RV_POIDImballoVendita, RV_POIDCalibro, "
            sSQL = sSQL & "IDTipoProdotto, RV_POIDTipoCategoria, RV_POIDUnitaDiMisuraLiq, "
            sSQL = sSQL & "RV_POQuantitaPerCollo, RV_POMoltiplicatore, PesoNetto, RV_POIDTipoPesoArticolo, "
            sSQL = sSQL & "IDIvaAcquisto, RV_POIDTipoLavorazione, RV_POIDImballoConferimento "
            sSQL = sSQL & "FROM Articolo WHERE IDArticolo=" & Me.CDArticolo.KeyFieldID
            
'            sSQL = "SELECT IDUnitaDiMisuraAcquisto, RV_POIDImballoConferimento, IDTipoProdotto, "
'            sSQL = sSQL & "IDIvaAcquisto, RV_POIDTipoLavorazione "
'            sSQL = sSQL & "FROM Articolo "
'            sSQL = sSQL & "WHERE IDArticolo=" & Me.CDArticolo.KeyFieldID
            
            Set rs = Cn.OpenResultset(sSQL)
            If rs.EOF = False Then
                If ((IsNull(m_DocumentsLink("IDArticolo").Value)) Or (fnNotNullN(m_DocumentsLink("IDArticolo").Value) <> Me.CDArticolo.KeyFieldID)) Then

                    Link_UnitaDiMisura_Acquisto = IIf(IsNull(rs!IDUnitaDiMisuraAcquisto), 0, rs!IDUnitaDiMisuraAcquisto)
                    Me.cboUM.WriteOn Link_UnitaDiMisura_Acquisto
                    If fnNotNullN(rs!RV_POIDImballoConferimento) > 0 Then
                        Me.CDCodiceImballo.Load fnNotNullN(rs!RV_POIDImballoConferimento)
                    End If
                    
                    IDIva = 0
                    
                    IDIva = GET_LINK_IVA_FORNITORE(Me.CDSocio.KeyFieldID)
                    IDIva = GET_LINK_IVA_LETTERA_INTENTO(Me.txtIDLetteraIntento.Value, IDIva)
                    If IDIva = 0 Then
                        Me.cboIVA.WriteOn fnNotNullN(rs!IDIvaAcquisto)
                    Else
                        Me.cboIVA.WriteOn IDIva
                    End If
                
                    If fnNotNullN(rs!RV_POIDTipoLavorazione) > 0 Then
                        Me.cboTipoLavorazione.WriteOn fnNotNullN(rs!RV_POIDTipoLavorazione)
                    End If
                    If Flag_GestioneArticoli = True Then
                        If Link_TipoGrezzo > 0 Then
                            If fnNotNullN(rs!IDTipoProdotto) <> Link_TipoGrezzo Then
                                MsgBox "ATTENZIONE!!" & vbCrLf & "L'articolo selezionato non è un prodotto grezzo", vbInformation, "Gestione articoli"
                            End If
                        End If
                    End If
                
                End If
                
                QUANTITA_PER_COLLO = fnNotNullN(rs!RV_POQuantitaPerCollo)
                MOLTIPLICATORE = fnNotNullN(rs!RV_POMoltiplicatore)
                PESO_LORDO_ARTICOLO = fnNotNullN(rs!PesoNetto)
                LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL = fnNotNullN(rs!RV_POIDTipoPesoArticolo)
                ParametroPesoArticolo
                If LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL = 0 Then
                    ParametroPesoArticolo
                    'TIPO_PESO_ARTICOLO = 0
                Else
                    TIPO_PESO_ARTICOLO = LINK_TIPO_PESO_LORDO_ARTICOLO_LOCAL
                End If

            End If
            rs.CloseResultset
            Set rs = Nothing
        End If
    'End If
    
    If (Me.CDArticolo.KeyFieldID > 0) Then
        If (m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value <= 0) Then
            If Me.txtNumeroLottoEntrata.Value = 0 Then
                Link_Lotto = GET_NUMERO_LOTTO
                Me.txtNumeroLottoEntrata.Value = Link_Lotto
            Else
                Link_Lotto = Me.txtNumeroLottoEntrata.Value
            End If
            Me.txtLottoDiEntrata.Text = fnGetCodiceLotto(1, 1, Me.txtLottoDiEntrata.Text, Link_Lotto)
        End If
    End If

End Sub

Private Sub cdCliente_ChangeElement()
On Error GoTo ERR_cdCliente_ChangeElement
    With Me.cboDestinazione
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT * FROM SitoPerAnagrafica WHERE IDAnagrafica=" & Me.cdCliente.KeyFieldID
        .Fill
    End With
Exit Sub
ERR_cdCliente_ChangeElement:
    MsgBox Err.Description, vbCritical, "cdCliente_ChangeElement"
End Sub

Private Sub CDCodiceImballo_ChangeElement()
On Error Resume Next
Dim Link_Lotto As Long

    If Me.CDCodiceImballo.KeyFieldID > 0 Then
        If ((IsNull(m_DocumentsLink("IDImballo").Value)) Or (fnNotNullN(m_DocumentsLink("IDImballo").Value) <> Me.CDCodiceImballo.KeyFieldID)) Then
            If GET_ARTICOLO_ANNULLATO(Me.CDCodiceImballo.KeyFieldID) = True Then
                MsgBox "ATTENZIONE!!!" & vbCrLf & "L'articolo selezionato è stato bloccato e pertanto non può essere utilizzato", vbInformation, "Inserimento dati"
                Me.txtImballo.Text = ""
                Me.CDCodiceImballo.Load 0
            Exit Sub
            End If
        End If
        
        Me.txtImballo.Text = Me.CDCodiceImballo.Description
        
        Me.txtTaraUnitaria.Value = fnGetTaraImballo(Me.CDCodiceImballo.KeyFieldID)
        
        
        If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) <= 0 Then
            Me.chkTracciaImballo.Value = GET_TRACCIA_IMBALLO(Me.CDCodiceImballo.KeyFieldID)
        End If
    End If
    
    Me.cboUMImballo.WriteOn GET_UM_ARTICOLO(Me.CDCodiceImballo.KeyFieldID)
    
    If (Me.CDArticolo.KeyFieldID > 0) Then
        If (m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value <= 0) Then
            If Me.txtNumeroLottoEntrata.Value = 0 Then
                Link_Lotto = GET_NUMERO_LOTTO
                Me.txtNumeroLottoEntrata.Value = Link_Lotto
            Else
                Link_Lotto = Me.txtNumeroLottoEntrata.Value
            End If
            Me.txtLottoDiEntrata.Text = fnGetCodiceLotto(1, 1, Me.txtLottoDiEntrata.Text, Link_Lotto)
        End If
    End If
    
    If Me.CDCodiceImballo.KeyFieldID = 0 Then
        Me.CDCodiceImballo.SetFocus
    Else
        Me.CDPedana.SetFocus
    End If
    
End Sub

Private Sub CDImballoGestione_AfterRunServerApplication(ByVal lKeyFieldID As Long)
    CDImballoGestione_ChangeElement
End Sub

Private Sub CDImballoGestione_ChangeElement()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraAcquisto, IDUnitaDiMisuraVendita, RV_POTracciabilitaImballo "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & Me.CDImballoGestione.KeyFieldID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.cboUMImbGest.WriteOn 0
    Me.chkTracciaImballoGest.Value = Unchecked
    
Else
    Me.cboUMImbGest.WriteOn fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
    Me.chkTracciaImballoGest.Value = Abs(fnNotNullN(rs!RV_POTracciabilitaImballo))
End If

rs.CloseResultset
Set rs = Nothing

End Sub

Private Sub CDPedana_ChangeElement()
On Error Resume Next
    
    If Me.CDPedana.KeyFieldID > 0 Then
        
        If ((IsNull(m_DocumentsLink("IDArticoloPedana").Value)) Or (fnNotNullN(m_DocumentsLink("IDArticoloPedana").Value) <> Me.CDPedana.KeyFieldID)) Then
            If GET_ARTICOLO_ANNULLATO(Me.CDPedana.KeyFieldID) = True Then
                MsgBox "ATTENZIONE!!!" & vbCrLf & "L'articolo selezionato è stato bloccato e pertanto non può essere utilizzato", vbInformation, "Inserimento dati"
                Me.CDPedana.Load 0
            Exit Sub
            End If
        End If
        
        Me.txtTaraPedana.Value = fnGetTaraImballo(Me.CDPedana.KeyFieldID)
        
        
        Me.cboUMPedana.WriteOn GET_UM_ARTICOLO(Me.CDPedana.KeyFieldID)
    End If
    
End Sub

Private Sub CDSocio_ChangeElement()
On Error Resume Next
Dim Testo As String

Dim rs As ADODB.Recordset
Dim sSQL As String

    

sSQL = "SELECT Anagrafica.IDAnagrafica, Anagrafica.Anagrafica, Anagrafica.Nome, Anagrafica.Cap, "
sSQL = sSQL & "Comune.Comune, Provincia.Provincia, Anagrafica.Indirizzo, "
sSQL = sSQL & "Fornitore.IDAzienda , Fornitore.IDCategoriaAnagrafica, Fornitore.Codice, "
sSQL = sSQL & "Anagrafica.IDComune, Anagrafica.IDNazione, Comune.IDProvincia, Provincia.IDRegione, Anagrafica.IDCategoriaAnagrafica AS IDCategoria "
sSQL = sSQL & "FROM Provincia RIGHT OUTER JOIN "
sSQL = sSQL & "Comune ON Provincia.IDProvincia = Comune.IDProvincia RIGHT OUTER JOIN "
sSQL = sSQL & "Anagrafica INNER JOIN "
sSQL = sSQL & "Fornitore ON Anagrafica.IDAnagrafica = Fornitore.IDAnagrafica ON dbo.Comune.IDComune = dbo.Anagrafica.IDComune "
sSQL = sSQL & "WHERE IDAzienda=" & m_App.IDFirm
sSQL = sSQL & "AND Anagrafica.IDAnagrafica=" & Me.CDSocio.KeyFieldID

    
Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection
    
If Not rs.EOF Then

    If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then
        If fnNotNullN(rs!IDCategoria) = Link_TipoSocio Then
            Testo = "ATTENZIONE!!!" & vbCrLf
            Testo = Testo & "E' stato selezionato un socio dentro un carico merce di acquisto" & vbCrLf
            Testo = Testo & "Vuoi continuare?"
            
            If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo inserimento") = vbNo Then
                Me.CDSocio.Load 0
                If Me.CDSocio.Enabled = True Then Me.CDSocio.SetFocus
            End If
        End If
    End If


    If Me.CDSocio.KeyFieldID > 0 Then
        Me.txtIndirizzo.Text = fnNotNull(rs!Indirizzo)
        Me.txtCAP.Text = fnNotNull(rs!Cap)
        Me.txtComune.Text = fnNotNull(rs!Comune)
        Me.txtProvincia.Text = fnNotNull(rs!Provincia)
        LINK_REGIONE_SOCIO = fnNotNullN(rs!IDRegione)
        LINK_NAZIONE_SOCIO = fnNotNullN(rs!IDNazione)
        LINK_COMUNE_SOCIO = fnNotNullN(rs!IDComune)
        LINK_PROVINCIA_SOCIO = fnNotNullN(rs!IDProvincia)
        Me.txtNomeSocio.Text = fnNotNull(rs!Nome)
        ATTIVA_LOTTO_PROD_ANA_FATT = GET_ATTIVA_LOTTO_PROD_ANA_FATT(Me.CDSocio.KeyFieldID)
    End If
Else
    Me.txtIndirizzo.Text = ""
    Me.txtCAP.Text = ""
    Me.txtComune.Text = ""
    Me.txtProvincia.Text = ""
    LINK_REGIONE_SOCIO = 0
    LINK_NAZIONE_SOCIO = 0
    LINK_COMUNE_SOCIO = 0
    LINK_PROVINCIA_SOCIO = 0
    Me.txtNomeSocio.Text = ""
    ATTIVA_LOTTO_PROD_ANA_FATT = 0
End If

If m_Document(m_Document.PrimaryKey).Value <= 0 Then
    Me.CDSocioFatt.Load GET_LINK_ANA_FATT(Me.CDSocio.KeyFieldID)
    Me.chkSelLottoProdAnaFatt.Value = GET_ATTIVA_LOTTO_PROD_ANA_FATT(Me.CDSocio.KeyFieldID)
    cboEsercizio_Click
    If Me.CDSocio.KeyFieldID > 0 Then
        Me.txtIDLetteraIntento.Value = GET_LINK_LETTERA_INTENTO_PRED(Me.CDSocio.KeyFieldID, 3, Me.txtDataConferimento.Text, TheApp.IDFirm)
    End If
    
    cmdNuovo_Click
End If

If Not (BrwMain.Visible) Then Change

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

If Not (BrwMain.Visible) Then Change
End Sub
Private Sub chkLottoChiuso_Click()
    If Me.chkLottoChiuso.Value = vbChecked Then
        Me.cboLiquidato.WriteOn fnNotNullN(LINK_STATO_LIQ_CHIUSO)
    End If
End Sub

Private Sub chkSelLottoProdAnaFatt_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdAnnotazioni_Click()
    frmAnnotazioni.Show vbModal
End Sub

Private Sub chkTrattaComeConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdApriFrameProvenienza_Click()
    If Me.Frame4.Height = 975 Then
        Me.Frame4.Height = 3135
    Else
        Me.Frame4.Height = 975
    End If
End Sub

Private Sub cmdApriFrameRiepilogo_Click()
    If Me.Frame3.Height = 975 Then
        Me.Frame3.Height = 3135
    Else
        Me.Frame3.Height = 975
    End If
End Sub

Private Sub cmdArticoliDerivati_Click()
If LINK_TIPO_ARTICOLO_CONFERITO > 1 Then
    frmArticoliDerivati.Show vbModal
    If Me.CDArticolo.KeyFieldID > 0 Then
        If Me.CDCodiceImballo.KeyFieldID > 0 Then
            Me.txtColli.SetFocus
        Else
            Me.CDCodiceImballo.SetFocus
        End If
    Else
        Me.CDArticolo.SetFocus
    End If
Else
    Me.CDArticolo.SetFocus
End If
End Sub

Private Sub cmdCalcoloMedi_Click()
    frmPesiMedi.Show
    
End Sub

Private Sub cmdCampionatura_Click()
On Error GoTo ERR_cmdCampionatura_Click
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) <= 0 Then Exit Sub
    
    Link_RigaConferimento = fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)
    LINK_ARTICOLO_CONFERITO = fnNotNullN(m_DocumentsLink("IDArticolo").Value)
    
    frmCampionatura.Show vbModal

Exit Sub
ERR_cmdCampionatura_Click:
    MsgBox Err.Description, vbCritical, "cmdCampionatura_Click"
End Sub

Private Sub cmdChiudiLotto_Click()
Dim sSQL As String
Dim Testo As String

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub


Testo = "ATTENZIONE!!!" & vbCrLf
Testo = Testo & "Questa procedura chiuderà tutti i lotti conferiti di questo documento" & vbCrLf & vbCrLf
Testo = Testo & "Continuare con questo comando?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Chiusura lotti") = vbNo Then Exit Sub

sSQL = "UPDATE RV_POCaricoMerceRighe SET "
sSQL = sSQL & " Chiuso=" & fnNormBoolean(1)
sSQL = sSQL & " WHERE IDRV_POCaricoMerceTesta=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
Cn.Execute sSQL

m_DocumentsLink.Refresh
End Sub

Private Sub cmdControlloCollegamenti_Click()
On Error GoTo ERR_cmdControlloCollegamenti_Click
    LINK_TESTA_DOCUMENTO = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    GET_CONTROLLO_COLLEGAMENTI_CONFERIMENTO LINK_TESTA_DOCUMENTO
    
    frmCollegamentiConferimento.Show vbModal

Exit Sub
ERR_cmdControlloCollegamenti_Click:
    MsgBox Err.Description, vbCritical, "cmdControlloCollegamenti_Click"
    
End Sub

Private Sub cmdCreaDocumento_Click()
Dim Testo As String
    
    If LINK_OGGETTO_ACQ > 0 Then
        Testo = "Il documento è stato passato già come documento di acquisto"
        MsgBox Testo, vbInformation, TheApp.FunctionName
        Exit Sub
    End If
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If Me.cboTipoDocumentoAcq.CurrentID = 0 Then Exit Sub
    
    If Me.txtDataDocumentoAcq.Value = 0 Then Exit Sub
    
    If Len(Trim(Me.txtNumeroDocumentoAcq.Text)) = 0 Then Exit Sub

    If m_Changed = True Then
        Testo = "Salvare il documento prima di procedere alla creazione del documento di acquisto"
        MsgBox Testo, vbCritical, TheApp.FunctionName
        Exit Sub
    End If

    If Me.cboTipoDocumentoAcq.CurrentID = 1 Then
        If LINK_FUNZIONE_FA = 0 Then
            Testo = "Per eseguire questo comando bisogna configurare il tipo di documento di acquisto"
            MsgBox Testo, vbCritical, TheApp.FunctionName
            
            Exit Sub
        End If
    End If
    If Me.cboTipoDocumentoAcq.CurrentID = 2 Then
        If LINK_FUNZIONE_DDT = 0 Then
            Testo = "Per eseguire questo comando bisogna configurare il tipo di documento di acquisto"
            MsgBox Testo, vbCritical, TheApp.FunctionName
            Exit Sub
        End If
    End If
    
    LINK_TESTA_DOCUMENTO = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    frmCreaFattura.Show vbModal
    
    LINK_OGGETTO_ACQ = GET_LINK_OGGETTO_ACQ(Me.cboTipoDocumentoAcq.CurrentID, fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
    
    If LINK_OGGETTO_ACQ > 0 Then
        Me.cboTipoDocumentoAcq.Enabled = False
        Me.txtNumeroDocumentoAcq.Enabled = False
        Me.txtDataDocumentoAcq.Enabled = False
        Me.cboVettore.Enabled = False
        Me.CDSocioFatt.Enabled = False
    Else
        Me.cboTipoDocumentoAcq.Enabled = True
        Me.txtNumeroDocumentoAcq.Enabled = True
        Me.txtDataDocumentoAcq.Enabled = True
        Me.cboVettore.Enabled = True
        Me.CDSocioFatt.Enabled = True
    End If
End Sub

Private Sub cmdElimina_Click()
On Error GoTo ERR_cmdElimina_Click
Dim TestoMsg As String
Dim Avvia As Boolean
Dim IDLottoImballo As Long


If Me.chkPreConferimento.Value = vbChecked Then Exit Sub

 Avvia = False

 If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) > 0 Then
     If GET_CONTROLLA_ESISTENZA_LIQUIDAZIONE(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)) = True Then
         MsgBox "Impossibile eliminare poichè il conferimento risulta legato ad una liquidazione", vbInformation, "Eliminazione dato"
     Exit Sub
     End If
 End If
 
 If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) > 0 Then
     If GET_CONTROLLA_ESISTENZA_COLLEGAMENTO(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value, fnGetTipoOggetto) = True Then
         MsgBox "Impossibile eliminare poichè il conferimento risulta legato al flusso della tracciabilità", vbInformation, "Eliminazione dato"
     Exit Sub
     End If
 End If
 If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) > 0 Then
     If Me.txtIDLottoImballo.Value > 0 Then
         If GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO(Me.txtIDLottoImballo.Value) = True Then
             MsgBox "Impossibile eliminare poichè il lotto imballo risulta movimentato", vbInformation, "Eliminazione dato"
             Exit Sub
         End If
     End If
 End If
 

 If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) > 0 Then
     TestoMsg = "Sei sicuro di voler elimare la riga di conferimento?" & vbCrLf

     If MsgBox(TestoMsg, vbInformation + vbYesNo, "Eliminazione riga") = vbYes Then
 
     If ATTIVAZIONE_NUOVO_CALCOLO = False Then
 
         If m_DocumentsLink("IDMovimento_Carico").Value > 0 Then
             ArrayDelete(ContDelete) = fnNotNullN(m_DocumentsLink("IDMovimento_Carico").Value)
             ContDelete = ContDelete + 1
         End If
         If m_DocumentsLink("IDMovimento_Vendita").Value > 0 Then
             ArrayDelete(ContDelete) = fnNotNullN(m_DocumentsLink("IDMovimento_Vendita").Value)
             ContDelete = ContDelete + 1
         End If
         If m_DocumentsLink("IDMovimentoImballo").Value > 0 Then
             ArrayDelete(ContDelete) = fnNotNullN(m_DocumentsLink("IDMovimentoImballo").Value)
             ContDelete = ContDelete + 1
         End If
 
         ArrayDeleteConferimento(ContDeleteConferimento) = m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value
         ContDeleteConferimento = ContDeleteConferimento + 1
 
     End If
 
 
         If Me.txtIDLottoImballo.Value > 0 Then
             DELETE_TMP_LOTTO_IMBALLO Me.txtIDLottoImballo.Value
         End If
         m_DocumentsLink.DeleteRowFromBuffer
         Avvia = True
         If Not (BrwMain.Visible) Then Change
     End If
 Else
     If MsgBox("Eliminare la riga selezionata?", vbInformation + vbYesNo, "Eliminazione riga") = vbYes Then
         If ATTIVAZIONE_NUOVO_CALCOLO = False Then
     
             If m_DocumentsLink("IDMovimento_Carico").Value > 0 Then
                 ArrayDelete(ContDelete) = fnNotNullN(m_DocumentsLink("IDMovimento_Carico").Value)
                 ContDelete = ContDelete + 1
             End If
             If m_DocumentsLink("IDMovimento_Vendita").Value > 0 Then
                 ArrayDelete(ContDelete) = fnNotNullN(m_DocumentsLink("IDMovimento_Vendita").Value)
                 ContDelete = ContDelete + 1
             End If
             If m_DocumentsLink("IDMovimentoImballo").Value > 0 Then
                 ArrayDelete(ContDelete) = fnNotNullN(m_DocumentsLink("IDMovimentoImballo").Value)
                 ContDelete = ContDelete + 1
             End If
     
             ArrayDeleteConferimento(ContDeleteConferimento) = m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value
             ContDeleteConferimento = ContDeleteConferimento + 1
     
         End If
         
         If Me.txtIDLottoImballo.Value > 0 Then
             DELETE_TMP_LOTTO_IMBALLO Me.txtIDLottoImballo.Value
         End If
        
         m_DocumentsLink.DeleteRowFromBuffer
         Avvia = True
         If Not (BrwMain.Visible) Then Change
     End If
 End If


 If Avvia = True Then
 End If
    
Exit Sub
ERR_cmdElimina_Click:
    MsgBox Err.Description, vbCritical, "cmdElimina_Click"
End Sub

Private Sub cmdEliminaAddebiti_Click()
On Error GoTo ERR_cmdEliminaImballo_Click
Dim Testo As String
Dim NumeroRecord As Long


Testo = "Sei sicuro di voler eliminare la riga?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Gestione altri addebiti") = vbNo Then Exit Sub



NumeroRecord = Me.GrigliaAddebiti.ListIndex - 1

m_DocumentsLink2.Delete

If Not (m_DocumentsLink2.BOF And m_DocumentsLink2.EOF) Then
    Me.GrigliaAddebiti.Recordset.Move NumeroRecord
End If

Exit Sub
ERR_cmdEliminaImballo_Click:
    MsgBox Err.Description, vbCritical, "Elimnazione imballi del conferimento"
End Sub

Private Sub cmdEliminaImballo_Click()
On Error GoTo ERR_cmdEliminaImballo_Click
Dim Testo As String
Dim NumeroRecord As Long
Dim IDOggetto As Long
Dim IDTipoOggetto As Long
Dim IDTipoProcesso As Long
Dim IDLottoImballo As Long

Testo = "Sei sicuro di voler eliminare la riga?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Gestione imballi da conferimento") = vbNo Then Exit Sub
If Me.cboTipoProcessoImballo.CurrentID = 1 Then
    
    If GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO(Me.txtIDLottoImballoGest.Value) = True Then
        Testo = "Impossibile eliminare poichè il lotto imballo risulta essere movimentato"
        
        MsgBox Testo, vbInformation, "Eliminazione carico imballo"
        
        Exit Sub
    End If
    
End If

IDOggetto = fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)
IDTipoOggetto = fnGetTipoOggetto("RV_POGestImbConf")
IDTipoProcesso = Me.cboTipoProcessoImballo.CurrentID
IDLottoImballo = Me.txtIDLottoImballoGest.Value

NumeroRecord = Me.GrigliaImballi.ListIndex - 1

m_DocumentsLink1.Delete
If IDTipoProcesso = 2 Then
    RIPRISTINO_LOTTI_DA_MOVIMENTO IDOggetto, IDTipoOggetto, IDOggetto
Else
    ELIMINAZIONE_LOTTO_IMBALLO IDLottoImballo
End If

If Not (m_DocumentsLink1.BOF And m_DocumentsLink1.EOF) Then
    Me.GrigliaImballi.Recordset.Move NumeroRecord
End If

'''''ELIMINAZIONE MOVIMENTI''''''''''''''''''''''''''''''''''''
Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection

Mov.IDTipoOggetto = IDTipoOggetto
Mov.IDOggetto = IDOggetto
Mov.Delete

Set Mov = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Sub
ERR_cmdEliminaImballo_Click:
    MsgBox Err.Description, vbCritical, "Elimnazione imballi del conferimento"

End Sub

Private Sub cmdEliminaRifLetInt_Click()
On Error GoTo ERR_cmdEliminaRifLetInt_Click
Dim Testo As String

If Me.txtIDLetteraIntento.Value = 0 Then Exit Sub
Testo = "Sei sicuro di voler eliminare il riferimento alla lettera d'intento?" & vbCrLf
If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento lettera d'intento") = vbNo Then Exit Sub

Me.txtIDLetteraIntento.Value = 0

Exit Sub

ERR_cmdEliminaRifLetInt_Click:
MsgBox Err.Description, vbCritical, "ERR_cmdEliminaRifLetInt_Click"
End Sub

Private Sub cmdGestioneQualita_Click()
On Error GoTo ERR_cmdGestioneQualita_Click

    If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) <= 0 Then Exit Sub

    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoApertura", "1"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoQualitaGestione", "5"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDAnagrafica", Me.CDSocio.KeyFieldID
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDRiferimento", fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)

    Shell MenuOptions.ProgramsPath & "\RV_POQualitaGestioneNew.exe"
    
Exit Sub
ERR_cmdGestioneQualita_Click:
    MsgBox Err.Description, vbCritical, "cmdGestioneQualita_Click"
End Sub

Private Sub cmdLavorazioneAutomatica_Click()
Dim Testo As String
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey)) <= 0 Then Exit Sub
    
    If GET_CONTROLLO_LAVORAZIONE(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey))) = True Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "Nella riga di conferimento corrispondono una o più lavorazioni, pertanto per utilizzare questo comando eliminare tutte le righe di lavorazione associate"
        MsgBox Testo, vbInformation, "Lavorazione automatica"
        Exit Sub
    End If
    
    Me.LinkLavorazione.IDReturn = fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey))
    Me.LinkLavorazione.RunApplication
    
End Sub

Private Sub cmdLetteraIntento_Click()
 If Me.CDSocio.KeyFieldID = 0 Then Exit Sub
    frmLetteraIntento.Show vbModal
End Sub

Private Sub cmdLottoImballo_Click()
    CONFERMA_LOTTO_IMBALLO_DA_UTENTE = Me.chkConfermaDaUtente.Value
    
    
    frmLottoImballo.Show vbModal
    
    Me.chkConfermaDaUtente.Value = CONFERMA_LOTTO_IMBALLO_DA_UTENTE
    

End Sub

Private Sub cmdNuovo_Click()
If Me.CDSocio.KeyFieldID = 0 Then Exit Sub
    
    If Me.chkPreConferimento.Value = vbChecked Then Exit Sub
    
    If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
        AggiornamentoDocumento = 1
        
        m_DocumentsLink.MoveLast
    
        If IsNull(m_DocumentsLink("IDArticolo").Value) Then
            m_DocumentsLink.DeleteRowFromBuffer
        End If
        
        AggiornamentoDocumento = 0
    End If
    
    
    If m_DocumentsLink.TableNew Then
        m_DocumentsLink.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    
    m_DocumentsLink.NewRow
    
    
    Me.cmdSalva.Enabled = True
    Me.cmdElimina.Enabled = True
    Me.CDArticolo.Load 0
    Me.CDCodiceImballo.Load 0
    Link_LottoArticolo = 0
    Me.cboRegione.WriteOn LINK_REGIONE_SOCIO
    Me.cboNazione.WriteOn LINK_NAZIONE_SOCIO
    Me.cboComune.WriteOn LINK_COMUNE_SOCIO
    Me.cboProvincia.WriteOn LINK_PROVINCIA_SOCIO
    Me.txtNomeMacchina.Text = GET_NOMECOMPUTER
    Me.txtUtenteMacchina.Text = GET_NOMEUTENTE
    Me.cboUtente.WriteOn TheApp.IDUser
    Me.txtOraArrivo.Text = GET_ORARIO(Now)
    Me.cboLiquidato.WriteOn LINK_STATO_LIQ_NUOVO
    Me.chkPrezzoMedio.Value = Abs(PREZZO_MEDIO_AUT)
    
    
    Me.cmdSelezionaLottoCampagna.SetFocus
    
    bVariazioneDettaglio = False
    
    
End Sub

Private Sub cmdNuovoAddebiti_Click()
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

    
    If m_DocumentsLink2.TableNew Then
        m_DocumentsLink2.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    
    m_DocumentsLink2.NewRow
    
    
        
    Me.CDArticoliAddebiti.SetFocus
    
End Sub

Private Sub cmdNuovoImballo_Click()
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    If Me.chkPreConferimento.Value = vbChecked Then Exit Sub
    
    If m_DocumentsLink1.TableNew Then
        m_DocumentsLink1.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    
    m_DocumentsLink1.NewRow
    
    
        
    Me.CDImballoGestione.SetFocus
    
End Sub

Private Sub cmdPesatura_Click()
On Error GoTo ERR_cmdPesatura_Click
Dim NumeroColliPesatura As Double
Dim TotalePesoLordoPesatura As Double
Dim TotalePezziPesatura As Double

Dim Testo As String

If Me.CDSocio.KeyFieldID = 0 Then Exit Sub
If Me.CDArticolo.KeyFieldID = 0 Then Exit Sub
If Me.chkPreConferimento.Value = vbChecked Then Exit Sub

If fnNotNullN(m_DocumentsLink("Link_Ordinamento").Value) <= 0 Then
    Link_Ordinamento_riga_conf = LINK_ORDINAMENTO
Else
    Link_Ordinamento_riga_conf = fnNotNullN(m_DocumentsLink("Link_Ordinamento").Value)
End If

NumeroColliPesatura = GET_SOMMA_COLLI_PES
TotalePesoLordoPesatura = GET_SOMMA_PESO_PES
TotalePezziPesatura = GET_SOMMA_PEZZI_PES

If ((NumeroColliPesatura <> Me.txtColli.Value) Or (TotalePesoLordoPesatura <> Me.txtPesoLordo.Value) Or (TotalePezziPesatura <> Me.txtPezzi.Value)) Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Il numero dei colli o dei pesi o dei pezzi del conferimento non sono uguali a quelli che risultano dalle pesature, "
    Testo = Testo & "pertanto continuando con questa funzionalità i dati inseriti potrebbero andare persi." & vbCrLf
    Testo = Testo & "Continuare con questo comando?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Validazione dati") = vbNo Then Exit Sub
End If


frmPesaturaPedane.Show vbModal

CalcoloPesoNetto 1

If CONFERMA_SALVA_PESATURA = 1 Then
    cmdSalva_Click
    If (SALVA_RIGA_OK = 1) Then
        OnSave
        If (SALVA_DOC_OK = 1) Then
            SetRigaConferimento Link_Ordinamento_riga_conf
        End If
    End If
End If
    
Exit Sub
ERR_cmdPesatura_Click:
    MsgBox Err.Description, vbCritical, "cmdPesatura_Click"
    
End Sub

Private Sub cmdPesaturaAut_Click()
On Error GoTo ERR_cmdPesaturaAut_Click
Dim NumeroColliPesatura As Double
Dim Testo As String

Dim TotalePesoLordoPesatura As Double
Dim TotalePezziPesatura As Double

If Me.CDSocio.KeyFieldID = 0 Then Exit Sub
If Me.CDArticolo.KeyFieldID = 0 Then Exit Sub
If Me.chkPreConferimento.Value = vbChecked Then Exit Sub

If fnNotNullN(m_DocumentsLink("Link_Ordinamento").Value) <= 0 Then
    Link_Ordinamento_riga_conf = LINK_ORDINAMENTO
Else
    Link_Ordinamento_riga_conf = fnNotNullN(m_DocumentsLink("Link_Ordinamento").Value)
End If

NumeroColliPesatura = GET_SOMMA_COLLI_PES
TotalePesoLordoPesatura = GET_SOMMA_PESO_PES
TotalePezziPesatura = GET_SOMMA_PEZZI_PES

If ((NumeroColliPesatura <> Me.txtColli.Value) Or (TotalePesoLordoPesatura <> Me.txtPesoLordo.Value) Or (TotalePezziPesatura <> Me.txtPezzi.Value)) Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Il numero dei colli o dei pesi o dei pezzi del conferimento non sono uguali a quelli che risultano dalle pesature, "
    Testo = Testo & "pertanto continuando con questa funzionalità i dati inseriti potrebbero andare persi." & vbCrLf
    Testo = Testo & "Continuare con questo comando?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Validazione dati") = vbNo Then Exit Sub
End If

frmPesaturaPedane.Show vbModal

CalcoloPesoNetto 1

If CONFERMA_SALVA_PESATURA = 1 Then
    cmdSalva_Click
    If (SALVA_RIGA_OK = 1) Then
        OnSave
        If (SALVA_DOC_OK = 1) Then
            SetRigaConferimento Link_Ordinamento_riga_conf
        End If
    End If
End If


Exit Sub
ERR_cmdPesaturaAut_Click:
    MsgBox Err.Description, vbCritical, "cmdPesaturaAut_Click"
End Sub

Private Sub cmdRiepilogo_Click()
    If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey)) > 0 Then
        Link_RigaConferimento = Me.GrigliaCorpo("IDRV_POCaricoMerceRighe").Value
        frmRiepilogo.Show vbModal
    End If
End Sub

Private Sub cmdRiepilogoLottoImballo_Click()
    If Me.txtIDLottoImballoGest.Value = 0 Then Exit Sub
    
    LINK_LOTTO_IMBALLO = Me.txtIDLottoImballoGest.Value
    
    frmRiepilogoLottoImb.Show vbModal
End Sub

Private Sub cmdRiepLottoImballoConf_Click()
    If Me.txtIDLottoImballo.Value = 0 Then Exit Sub
    
    LINK_LOTTO_IMBALLO = Me.txtIDLottoImballo.Value
    
    frmRiepilogoLottoImb.Show vbModal
End Sub

Private Sub cmdSalva_Click()
Dim Variazione As Boolean
Dim Link_Lotto As Long
Dim Testo As String

If Me.chkPreConferimento.Value = vbChecked Then Exit Sub

Variazione = bVariazioneDettaglio
SALVA_RIGA_OK = 0
CalcoloPesoNetto 1

If LINK_OGGETTO_ACQ > 0 Then
    Testo = "ATTENZIONE!!!!!" & vbCrLf
    Testo = Testo & "Il documento è collegato ad un documento di acquisto pertanto è impossibile continuare"
    MsgBox Testo, vbCritical, "Salvataggio documento"
Exit Sub
End If

If PermessoSalvataggio = True Then

    If Me.chkPreConferimento.Value = vbChecked Then
        m_DocumentsLink("IDRV_POTipoConfLiquidazione").Value = Me.cboLiquidato.CurrentID
        m_DocumentsLink("PrezzoMedio").Value = Me.chkPrezzoMedio.Value
        m_DocumentsLink("Chiuso").Value = Me.chkLottoChiuso.Value
        m_DocumentsLink("IDRV_POTipoLavorazione").Value = Me.cboTipoLavorazione.CurrentID
    Else

        m_DocumentsLink("IDLottoDiCampagna").Value = Me.txtIDLottoCampagna.Value
        m_DocumentsLink("LottoDiConferimento").Value = Me.txtLottoDiConferimento.Text
        m_DocumentsLink("IDArticolo").Value = Me.CDArticolo.KeyFieldID
        m_DocumentsLink("CodiceArticolo").Value = Me.CDArticolo.Code
        m_DocumentsLink("Articolo").Value = Me.TxtArticolo.Text
        m_DocumentsLink("IDUnitaDiMisura").Value = Link_UnitaDiMisura_Coop
        m_DocumentsLink("IDUnitaDiMisuraDiamante").Value = Me.cboUM.CurrentID
        m_DocumentsLink("Colli").Value = Me.txtColli.Value
        m_DocumentsLink("PesoLordo").Value = Me.txtPesoLordo.Value
        m_DocumentsLink("PesoNetto").Value = Me.txtPesoNetto.Value
        m_DocumentsLink("TaraUnitaria").Value = Me.txtTaraUnitaria.Value
        m_DocumentsLink("Tara").Value = Me.txtTara.Value
        m_DocumentsLink("Pezzi").Value = Me.txtPezzi.Value
        m_DocumentsLink("Qta_UM").Value = Me.txtQta_UM.Value
        m_DocumentsLink("IDImballo").Value = Me.CDCodiceImballo.KeyFieldID
        m_DocumentsLink("CodiceImballo").Value = Me.CDCodiceImballo.Code
        m_DocumentsLink("DescrizioneImballo").Value = Me.txtImballo.Text
        m_DocumentsLink("Chiuso").Value = Me.chkLottoChiuso.Value
        m_DocumentsLink("CodiceLotto").Value = Me.txtLottoDiEntrata.Text
        m_DocumentsLink("IDCodiceLotto").Value = Me.txtNumeroLottoEntrata.Value
        m_DocumentsLink("DataSbloccoLottoCampagna").Value = Me.txtDataSbloccoLotto.Value
        m_DocumentsLink("IDRegione").Value = Me.cboRegione.CurrentID
        m_DocumentsLink("IDNazione").Value = Me.cboNazione.CurrentID
        m_DocumentsLink("IDComune").Value = Me.cboComune.CurrentID
        m_DocumentsLink("IDProvincia").Value = Me.cboProvincia.CurrentID
        
        m_DocumentsLink("NomePC").Value = Me.txtNomeMacchina.Text
        m_DocumentsLink("UtentePC").Value = Me.txtUtenteMacchina.Text
        m_DocumentsLink("IDUtente").Value = Me.cboUtente.CurrentID
        m_DocumentsLink("CodiceUtente").Value = Me.txtCodiceUtente.Text
        
        m_DocumentsLink("IDArticoloPedana").Value = Me.CDPedana.KeyFieldID
        m_DocumentsLink("TaraPedana").Value = Me.txtTaraPedana.Value
        m_DocumentsLink("CodiceArticoloPedana").Value = Me.CDPedana.Code
        m_DocumentsLink("ArticoloPedana").Value = Me.CDPedana.Description
        m_DocumentsLink("QuantitaPedana").Value = Me.txtQuantitaPedana.Value
        
        m_DocumentsLink("IDIvaAcquisto").Value = Me.cboIVA.CurrentID
        m_DocumentsLink("AliquotaIVA").Value = Me.txtAliquotaIVA.Value
        m_DocumentsLink("ImportoUnitario").Value = Me.txtImportoUnitario.Value
        m_DocumentsLink("TotaleImponibileRiga").Value = Me.txtImponibileRiga.Value
        m_DocumentsLink("TotaleImpostaRiga").Value = Me.txtImpostaRiga.Value
        m_DocumentsLink("TotaleLordoRiga").Value = Me.txtTotaleRiga.Value
        m_DocumentsLink("OraArrivoMerce").Value = txtOraArrivo.Text
    
        m_DocumentsLink("IDRV_POTipoLavorazione").Value = Me.cboTipoLavorazione.CurrentID
        m_DocumentsLink("IDRV_POTipoConfLiquidazione").Value = Me.cboLiquidato.CurrentID
        m_DocumentsLink("PrezzoMedio").Value = Me.chkPrezzoMedio.Value
        m_DocumentsLink("TaraMezzo").Value = Me.txtTaraAutomezzo.Value
           
        m_DocumentsLink("IDRV_POLottoImballo").Value = Me.txtIDLottoImballo.Value
        m_DocumentsLink("TracciaImballo").Value = Me.chkTracciaImballo.Value
           
        m_DocumentsLink("ImportoUnitarioImballo").Value = Me.txtImpUniImb.Value
        m_DocumentsLink("QuantitaPresunta").Value = Me.txtQtaPresunta.Value
        m_DocumentsLink("AnnotazioniAggiuntive").Value = Me.txtNoteAgg.Text
        m_DocumentsLink("LottoImballoUtilizzoEsclusivo").Value = Me.chkLottoImbEsclusivo.Value
        If fnNotNullN(m_DocumentsLink("Link_Ordinamento").Value) = 0 Then
            m_DocumentsLink("Link_Ordinamento").Value = LINK_ORDINAMENTO
            LINK_ORDINAMENTO = LINK_ORDINAMENTO + 1
        End If
    End If
    
    m_DocumentsLink.SaveRowToBuffer
    
    SALVA_RIGA_OK = 1
    
    m_DocumentsLink.Move Me.GrigliaCorpo.ListIndex - 1
    
    If Not (BrwMain.Visible) Then Change
    
    If Variazione = False Then
        cmdNuovo_Click
    End If
End If

End Sub



Private Sub cmdSalvaAddebiti_Click()
On Error GoTo ERR_cmdSalvaAddebiti_Click
Dim OLD_CURSOR As Long
Dim Testo As String

    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If Me.CDArticoliAddebiti.KeyFieldID = 0 Then
        MsgBox "Inserire l'articolo", vbCritical, "Gestione altri addebiti"
        Me.CDArticoliAddebiti.SetFocus
        Exit Sub
    End If
        
    If Me.txtQtaAddebiti.Value = 0 Then
        MsgBox "Inserire la quantità", vbCritical, "Gestione altri addebiti"
        Me.txtQtaAddebiti.SetFocus
        Exit Sub
    End If

    If Me.cboTipoTrattenutaAggiuntiva.CurrentID = 0 Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Il tipo di trattenuta aggiuntiva non è stato inserito" & vbCrLf
        Testo = Testo & "Vuoi continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Salvattaggio addebititi aggiuntivi") = vbNo Then Exit Sub
    End If
    
    m_DocumentsLink2("IDArticolo").Value = Me.CDArticoliAddebiti.KeyFieldID
    m_DocumentsLink2("IDRV_POTipoTrattenutaAggiuntiva").Value = Me.cboTipoTrattenutaAggiuntiva.CurrentID
    m_DocumentsLink2("Quantita").Value = Me.txtQtaAddebiti.Value
    m_DocumentsLink2("ImportoUnitario").Value = Me.txtImpUniAddebiti.Value
    m_DocumentsLink2("IDIva").Value = Me.cboIvaAddebiti.CurrentID
    m_DocumentsLink2("TotaleRigaNettoIva").Value = Me.txtImponibileAddebiti.Value
    m_DocumentsLink2("ImpostaRiga").Value = Me.txtImpostaAddebiti.Value
    m_DocumentsLink2("TotaleRigaLordoIva").Value = Me.txtTotaleRigaAddebiti.Value
    
    OLD_CURSOR = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    m_DocumentsLink2.Save
    
    
    Cn.CursorLocation = OLD_CURSOR
    
    m_DocumentsLink2.Move Me.GrigliaAddebiti.ListIndex - 1
    
    
    
Exit Sub
ERR_cmdSalvaAddebiti_Click:
    MsgBox Err.Description, vbCritical, "Errore di salvataggio"
End Sub

Private Sub cmdSalvaImballo_Click()
On Error GoTo ERR_cmdSalvaImballo_Click
Dim OLD_CURSOR As Long
Dim Testo As String
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    If Me.chkPreConferimento.Value = vbChecked Then Exit Sub
    
    If Me.CDImballoGestione.KeyFieldID = 0 Then
        MsgBox "Inserire l'imballo da movimentare", vbCritical, "Gestione altri imballi"
        Exit Sub
    End If
        
    If Me.cboUMImbGest.CurrentID = 0 Then
        MsgBox "Inserire l'unita di misura dell'articolo imballo", vbCritical, "Gestione altri imballi"
        Exit Sub
    End If
    
    If Me.cboTipoProcessoImballo.CurrentID = 0 Then
        MsgBox "Inserire il processo di movimentazione per l'imballo", vbCritical, "Gestione altri imballi"
        Exit Sub
    End If
    
    If Me.txtQuantitaImballo.Value = 0 Then
        MsgBox "Inserire una quantita maggiore di zero", vbCritical, "Gestione altri imballi"
        Exit Sub
    End If
  
    If Me.cboTipoProcessoImballo.CurrentID = 1 Then
        If (Me.chkTracciaImballoGest.Value = vbChecked) Then
            Me.txtIDLottoImballoGest.Value = CREA_MODIFICA_LOTTO_IMBALLO(Me.txtIDLottoImballoGest.Value, 0, 0, fnNotNullN(m_Document(m_Document.PrimaryKey).Value), Me.txtQuantitaImballo.Value, fnGetTipoOggetto("RV_POGestImbConf"), Me.CDSocio.KeyFieldID, Me.CDImballoGestione.KeyFieldID, Me.txtDataConferimento.Text, Me.txtNumeroConferimento.Text, Me.CDSocio.Code, 0, Me.txtLottoImballoRifEsterno.Text)
        End If
    Else
        If Me.txtQuantitaImballo.Value > Me.txtDispLottoImballo.Value Then
            If Me.chkTracciaImballoGest.Value > 0 Then
                Testo = "ATTENZIONE!!" & vbCrLf
                Testo = Testo & "L'imballo selezionato è configurato per la tracciabilità, tuttavia la disponibilità dei lotti non è sufficiente " & vbCrLf
                'Testo = Testo & "la disponibilità dei lotti non è abbastanza per evadere la richiesta " & vbCrLf
                Testo = Testo & "Vuoi continuare?"
                If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo tracciabilità lotti imballo") = vbNo Then Exit Sub
            End If
        End If
    End If
    
    m_DocumentsLink1("IDArticolo").Value = Me.CDImballoGestione.KeyFieldID
    m_DocumentsLink1("IDRV_POTipoProcessoCoop").Value = Me.cboTipoProcessoImballo.CurrentID
    m_DocumentsLink1("Quantita").Value = Me.txtQuantitaImballo.Value
    m_DocumentsLink1("TracciaImballo").Value = Me.chkTracciaImballoGest.Value
    m_DocumentsLink1("IDRV_POLottoImballo").Value = Me.txtIDLottoImballoGest.Value
    m_DocumentsLink1("ConfermaDaUtente").Value = Me.chkConfermaDaUtente.Value
    m_DocumentsLink1("Importo").Value = Me.txtImpUniAltriImb.Value
    m_DocumentsLink1("RiferimentoEsterno").Value = Me.txtLottoImballoRifEsterno.Text

    OLD_CURSOR = Cn.CursorLocation
    Cn.CursorLocation = adUseClient
    
    m_DocumentsLink1.Save
    
    ''''MOVIMENTAZIONE IMBALLI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MOVIMENTAZIONE_GESTIONE_IMBALLI fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), Me.cboTipoProcessoImballo.CurrentID, 20
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Cn.CursorLocation = OLD_CURSOR
    
    
    CREA_RECORDSET_LOTTI_IMBALLI Me.CDImballoGestione.KeyFieldID, fnGetTipoOggetto("RV_POGestImbConf"), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)
    
    
    
    m_DocumentsLink1.Move Me.GrigliaImballi.ListIndex - 1
    
    
Exit Sub
ERR_cmdSalvaImballo_Click:
    MsgBox Err.Description, vbCritical, "Errore di salvataggio"
    
End Sub

Private Sub cmdSelezionaLottoCampagna_Click()
Dim Link_Lotto As Long
On Error GoTo ERR_cmdSelezionaLottoCampagna_Click
'If GET_ESISTENZA_BIO = True Then
    If Me.chkPreConferimento.Value = vbChecked Then Exit Sub

    frmSelezionaLottoDiCampagna.Show vbModal
    If Me.txtIDLottoCampagna.Value > 0 Then
        GET_DATI_LOTTO_CAMPAGNA Me.txtIDLottoCampagna.Value
    End If
    If Len(Trim(Me.txtLottoDiConferimento.Text)) > 0 Then
        If (Me.CDArticolo.KeyFieldID > 0) Then
            If (m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value <= 0) Then
                If Me.txtNumeroLottoEntrata.Value = 0 Then
                    Link_Lotto = GET_NUMERO_LOTTO
                    Me.txtNumeroLottoEntrata.Value = Link_Lotto
                Else
                    Link_Lotto = Me.txtNumeroLottoEntrata.Value
                End If
                Me.txtLottoDiEntrata.Text = fnGetCodiceLotto(1, 1, Me.txtLottoDiEntrata.Text, Link_Lotto)
            End If
        End If
        If Me.CDArticolo.KeyFieldID > 0 Then
            Me.txtColli.SetFocus
        Else
            If LINK_TIPO_ARTICOLO_CONFERITO > 1 Then
                frmArticoliDerivati.Show vbModal
                If Me.CDArticolo.KeyFieldID > 0 Then
                    If Me.CDCodiceImballo.KeyFieldID > 0 Then
                        Me.txtColli.SetFocus
                    Else
                        Me.CDCodiceImballo.SetFocus
                    End If
                Else
                    Me.CDArticolo.SetFocus
                End If
            Else
                Me.CDArticolo.SetFocus
            End If
        End If
    Else
        Me.txtLottoDiConferimento.SetFocus
    End If



    Exit Sub
'End If
ERR_cmdSelezionaLottoCampagna_Click:
    MsgBox Err.Description, vbCritical, "cmdSelezionaLottoCampagna_Click"
End Sub

Private Sub cmdSelezionaOrdine_Click()
Dim Testo As String
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim IDCliente As Long
Dim NumeroDocumento As Long
Dim DataDocumento As String

IDCliente = Me.cdCliente.KeyFieldID
NumeroDocumento = Me.txtNumeroOrdine.Value
DataDocumento = Me.txtDataOrdine.Text

    Me.cdCliente.Load IDCliente
    Me.txtNumeroOrdine.Value = NumeroDocumento
    Me.txtDataOrdine.Text = DataDocumento

    If GET_NUMERO_ORDINI_DA_RICERCA(Me.cdCliente.KeyFieldID, Me.txtNumeroOrdine.Value, Me.txtDataOrdine.Text, Me.txtDataPartenza.Text) > 1 Then
        frmTrovaOrdine.Show vbModal
    End If


End Sub
Private Function GET_NUMERO_ORDINI_DA_RICERCA(IDCliente As Long, NumeroOrdine As Long, DataOrdine As String, DataPartenza As String) As Long
Dim sSQL As String
Dim sSQL_WHERE As String
Dim rs As DmtOleDbLib.adoResultset
Dim I As Long
Dim IDOrdineReturn As Long

sSQL_WHERE = ""

sSQL = "SELECT * FROM RV_POIERicercaOrdine "
sSQL = sSQL & "WHERE Doc_ordine_chiuso = 0 "
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDFiliale=" & TheApp.Branch


If Me.cdCliente.KeyFieldID > 0 Then
    sSQL = sSQL & " AND Link_nom_anagrafica=" & IDCliente
End If

If Me.txtNumeroOrdine.Value > 0 Then
    sSQL = sSQL & " AND Doc_numero=" & NumeroOrdine
End If

If Me.txtDataOrdine.Value > 0 Then
    sSQL = sSQL & " AND Doc_data=" & fnNormDate(DataOrdine)
End If
    
If Me.txtDataPartenza.Value > 0 Then
    sSQL = sSQL & " AND Doc_data_prevista_evasione=" & fnNormDate(DataPartenza)
End If

I = 0

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    I = I + 1
    IDOrdineReturn = fnNotNullN(rs!IDOggetto)
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

If I = 1 Then
    GET_NUMERO_ORDINI_DA_RICERCA = I
    Me.txtIDOrdineCliente.Value = IDOrdineReturn
Else
    GET_NUMERO_ORDINI_DA_RICERCA = I
End If

End Function

Private Sub cmdTour_Click()
If fnNotNullN(m_Document("IDOggetto").Value) > 0 Then
    Link_Oggetto = fnNotNullN(m_Document("IDOggetto").Value)
    frmTour.Show vbModal
End If
End Sub

Private Sub cmdTracciabilitaLottoCampagna_Click()
On Error GoTo ERR_cmdTracciabilitaLottoCampagna_Click
'If GET_ESISTENZA_BIO = True Then
    frmTracciabilitaBio.Show vbModal
    Exit Sub
'End If
ERR_cmdTracciabilitaLottoCampagna_Click:
    MsgBox Err.Description, vbCritical, "cmdTracciabilitaLottoCampagna_Click"
End Sub



Private Sub Command1_Click()
On Error GoTo ERR_cmdGestioneQualita_Click

    If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) <= 0 Then Exit Sub

    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoApertura", "2"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDTipoQualitaGestione", "5"
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDAnagrafica", Me.CDSocio.KeyFieldID
    SaveSetting REGISTRY_KEY, "RV_POQualitaGestioneNew", "IDRiferimento", fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)

    Shell MenuOptions.ProgramsPath & "\RV_POQualitaGestioneNew.exe"
    
Exit Sub
ERR_cmdGestioneQualita_Click:
    MsgBox Err.Description, vbCritical, "Command1_Click"
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

        Me.txtPesoLordo.DecimalPlaces = Numero_Decimali_Pesi
        Me.txtPesoNetto.DecimalPlaces = Numero_Decimali_Pesi
        Me.txtTara.DecimalPlaces = Numero_Decimali_Pesi
        
        m_bOnFirstTime = False

        'Se il filtro di default restituisce dei record si va in modalità variazione
        'ma solo se il primo record non è bloccato altrimenti si va in modalità tabellare
        If Not (m_Document.EOF = True And m_Document.BOF = True) Then
            'Il filtro ha restituito almeno un record
             
            'Controlla se il primo record su cui si dovrebbe andare in variazione è bloccato.
            If m_Semaphore.IsActionAvailable(m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions) Then
                'Il primo record NON è bloccato
                'allora si effettua il blocco e si va in modalità Variazione
                
                m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
                    
                'La vista alla partenza deve essere quella del Form
                BrwMain.Visible = False

                'Imposta la modalità variazione
                SetStatus4Modality Modify
                
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If FORM_PESI_MEDI_SHOW = True Then
        
        Unload frmPesiMedi
        
        
    End If
    
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
    
    
    Select Case KeyCode
        Case vbKeyPageDown
            DMTSplitBar1.ScrollDown
        Case vbKeyPageUp
            DMTSplitBar1.ScrollUp
    End Select

    If KeyCode = vbKeyReturn Then
        cmdSalva_Click
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
    If KeyCode = vbKeyF4 Then
        If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey)) > 0 Then
            cmdRiepilogo_Click
        End If
    End If
    
    If KeyCode = vbKeyF8 Then
        cmdPesaturaAut_Click
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
    
    'chiude e distrugge il riferimento alle connessioni
        'CloseConnection
    'Distrugge il riferimento al recordset
    Set BrwMain.Recordset = Nothing
    
    Cancel = FormUnload
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
            SetStatus4Modality Find
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
                SetStatus4Modality Browse
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

Private Sub LabelLink1_AfterRunServerApplication(ByVal lIDResultKey As Long)
Dim NumeroRecord As Long
NumeroRecord = Me.GrigliaCorpo.ListIndex - 1

m_DocumentsLink.Refresh

m_DocumentsLink.Move NumeroRecord

End Sub

Private Sub LabelLink1_BeforeRunServerApplication()
    Me.LabelLink1.IDReturn = fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)
End Sub


Private Sub lblArticolo_Click()
    Dim oSearch As dmtFind.Find
    Dim sSQL As String
    Dim oRes As DmtOleDbLib.adoResultset
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
    
    oSearch.Filters.Add "Articolo", Me.TxtArticolo.Text
    
    'Con la query SQL montata sotto l'impostazione di questa proprietà
    'non è necessaria
    'oSearch.IDName = "Comune.IDComune"
 
    'Quando si apre la finestra viene effettuta una ricerca preliminare con il filtro
    'SELECT ... FROM ... WHERE Comune LIKE ‘XXX%’  (essendo TextBox(0).Text = "XXX")
    oSearch.Start = Me.TxtArticolo.Text

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
            Me.CDArticolo.Load oRes!IDArticolo
            Me.CDArticolo.Code = oRes!CodiceArticolo
            If Me.CDArticolo.KeyFieldID > 0 Then
                Me.TxtArticolo.Text = oRes!Articolo
                Link_UnitaDiMisura_Coop = fnGetUMCoop(oRes!IDUnitaDiMisuraAcquisto)
            End If
    End If
            
End If
    
    Set oRes = Nothing
    Set oSearch = Nothing
End Sub

Private Sub lblImballo_Click()
Dim oSearch As dmtFind.Find
Dim sSQL As String
Dim oRes As DmtOleDbLib.adoResultset
If Me.CDCodiceImballo.Code = "" Then
    'Crea un'istanza dell'oggetto Find
    Set oSearch = New dmtFind.Find
    
    'Assegna la connessione aperta
    oSearch.Database = Cn
    
    'La Caption della finestra di ricerca
    oSearch.Caption = "Imballi"
    
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
    
    oSearch.Filters.Add "Articolo", Me.txtImballo.Text
    
    'Con la query SQL montata sotto l'impostazione di questa proprietà
    'non è necessaria
    'oSearch.IDName = "Comune.IDComune"
 
    'Quando si apre la finestra viene effettuta una ricerca preliminare con il filtro
    'SELECT ... FROM ... WHERE Comune LIKE ‘XXX%’  (essendo TextBox(0).Text = "XXX")
    oSearch.Start = Me.txtImballo.Text

    'Query SQL con cui effettuare le ricerche in base dati.
    'Attenzione:
    'Il campo chiave primaria (Comune.IDComune in questo caso) deve essere presente
    'nella SELECT
        sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
        sSQL = sSQL & "FROM Articolo "
        sSQL = sSQL & "WHERE ((IDTipoProdotto = " & Link_TipoImballo & ") "
        'sSQL = sSQL & "AND (GestioneLotti=" & fnNormBoolean(1) & ") "
        sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & "))"
        
    
    
    
    'Assegnazione della query di ricerca
    oSearch.SQL = fnAnsi2Jet(sSQL)
    
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
   
    
    Set oRes = oSearch.Exec
    
    
    If Not oRes.EOF Then
        Me.CDCodiceImballo.Load oRes!IDArticolo
        Me.CDCodiceImballo.Code = oRes!CodiceArticolo
        If Me.CDCodiceImballo.KeyFieldID > 0 Then
            Me.txtImballo.Text = oRes!Articolo
            Me.txtTaraUnitaria.Value = fnGetTaraImballo(Me.CDCodiceImballo.KeyFieldID)
            Me.txtLottoDiConferimento.SetFocus
        End If
    End If
            
End If
    
    Set oRes = Nothing
    Set oSearch = Nothing
End Sub


Private Sub LinkLavorazione_AfterRunServerApplication(ByVal lIDResultKey As Long)
    Me.txtQtaVenduta.Value = GET_RIEPILOGO_QUANTITA_VENDUTO(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value))
    Me.txtQtaAssegnata.Value = GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)) + GET_RIEPILOGO_QUANTITA_PROCESSO(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value))
    Me.txtQtaQuadrata.Value = GET_RIEPILOGO_QUANTITA_LAVORAZIONE(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value))
    Me.txtQtaDifferenza.Value = Me.txtQta_UM.Value - (Me.txtQtaQuadrata.Value + Me.txtQtaVenduta.Value)
    Me.txtQtaDifferenzaLavorazione.Value = Me.txtQta_UM.Value - (Me.txtQtaQuadrata.Value + Me.txtQtaAssegnata.Value)
End Sub

Private Sub LinkLavorazione_BeforeRunServerApplication()
    SaveSetting Trim(gResource.GetMessage(LBL_REGISTRY_KEY)), "RV_POAssegnazioneMerce", "LavorazioneAutomatica", 1
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
    
        'cbcx
        'QUESTA PARTE DEVE ESSERE RIVISTA
        '-----------------------------------------------------------------
        
'''''        Dim ErrorMsg As String
'''''
'''''        ErrorMsg = "No processes to execute" & vbCrLf
'''''        ErrorMsg = ErrorMsg & "This application is able to execute these processes:" & vbCrLf
'''''        '*
'''''        'Inserire i processi che l'applicazione sa eseguire
'''''        '*
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE & vbCrLf
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_EXTENDED_DATABASE & vbCrLf
'''''        'ErrorMsg = ErrorMsg & PROCESS_MANUTENZIONE_DA_SHELL & vbCrLf
'''''        Err.Raise ERR_NO_PROCESSES, , ErrorMsg


    End Select
    Exit Sub
ErrorHandler:
    SemaphoreUnlock
    ShowErrorLog
End Sub

Private Sub m_Document_OnReposition()
        
    'Viene creata (se non è già stato fatto) la collezione FormFields
    CreateFormFields
        
    If Not m_Document.TableNew Then
        'Se EOF = true o BOF = true vuol dire che si è andati oltre l'ultimo o
        'prima del primo record. In tal caso non si deve fare il refresh dei
        'controlli del form.
        If Not (m_Document.EOF Or m_Document.BOF) Then
            BrowseReposition
            
            'cbcx
            '---------------------------------------------
            'Gestione processo On_Extend
'            If Not m_ExtendApplication Is Nothing Then
'                'Notifica l'identificativo unico del documento corrente
'                m_ExtendApplication.PrimaryID = m_Document.Fields("ID" & m_App.TableName).Value
'            End If
            
        End If
    Else
        'Nel caso di inserimento nuovo record ripulisce i campi del form
        ClearFormFields
    End If
    
    
    
    'rif11 begin
    
   'If Me.GrigliaFasiIntervento.ColumnsHeader.Count > 0 Then
       On Error Resume Next
        'Binding mediante le proprietà DataMember e DataSource.
        'Me.GrigliaFasiIntervento.DataMember = m_DocumentsLink2.TableName
        'Set Me.GrigliaFasiIntervento.DataSource = m_Document

        'Binding mediante la proprietà Recordset
        Set Me.GrigliaCorpo.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink.TableName).Data
        Set Me.GrigliaImballi.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink1.TableName).Data
        Set Me.GrigliaAddebiti.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink2.TableName).Data
        
    'End If
'Variabile contatore che serve per il consolidamento delle righe eliminate

ContDelete = 0
ContDeleteConferimento = 0

RefreshArray
ParametriDiDefault Date
'INIZIALIZZAZIONE DEL DOCUMENTO
    If m_Document("IDRV_POCaricoMerceTesta").Value <= 0 Then
        Me.txtDataDocumento.Text = Date
        Link_Esercizio = fnGetEsercizio(Me.txtDataDocumento.Text)
        Link_PeriodoIVA = fnGetPeriodoIVA(Me.txtDataDocumento.Text)
        Me.cboMagazzinoConf.WriteOn Link_Magazzino_Conferimento
        Me.CboMagazzinoVend.WriteOn Link_Magazzino_Vendita
        Me.cboSezionale.WriteOn Link_Sezionale
        Me.txtDataDocumento.Value = Date
        Me.txtNumeroDocumento.Value = fnGetNumeroDocumento
        Me.txtDataConferimento.Value = Date
        Me.cboEsercizio.WriteOn Link_Esercizio
        
        'Me.txtDataDocumento.Enabled = True
        Me.txtNumeroDocumento.Enabled = True
        Me.CDSocio.Enabled = True
        Me.cboMagazzinoConf.Enabled = True
        Me.CboMagazzinoVend.Enabled = True
        Me.cboSezionale.Enabled = True
        Link_Oggetto = 0
        CambioSocio = False
        Me.CDSocio.SetFocus
        Link_Socio_OLD = 0
        LINK_OGGETTO_ACQ = 0
        Me.chkPreConferimento.Value = 0
    Else
        Link_Esercizio = fnGetEsercizio(Me.txtDataDocumento.Text)
        Link_PeriodoIVA = fnGetPeriodoIVA(Me.txtDataDocumento.Text)
        'Me.txtDataDocumento.Enabled = False
        CambioSocio = False
        Me.txtNumeroDocumento.Enabled = False
        Me.CDSocio.Enabled = False
        Me.cboMagazzinoConf.Enabled = False
        Me.CboMagazzinoVend.Enabled = False
        Me.cboSezionale.Enabled = False
        Link_Oggetto = fnNotNullN(m_Document("IDoggetto").Value)
        CodiceSocio = fnGetCodiceSocio(fnNotNullN(m_Document("IDAnagrafica")))
        Link_Socio_OLD = fnNotNullN(m_Document("IDAnagrafica").Value)
        LINK_OGGETTO_ACQ = GET_LINK_OGGETTO_ACQ(Me.cboTipoDocumentoAcq.CurrentID, fnNotNullN(m_Document(m_Document.PrimaryKey).Value))
        Me.chkPreConferimento.Value = Abs(fnNotNullN(m_Document("GeneratoDaPreConferimento")))
    End If
    
    Me.cboVettore.Enabled = True
    Me.txtImportoTrasporto.Enabled = True
    Me.cboLuogoPresaMerce.Enabled = True
    Me.txtNumeroDocumento.Enabled = True
    Me.txtDataConferimento.Enabled = True
    
    Me.txtNumeroConferimento.Enabled = True
    Me.txtPrefissoConferimento.Enabled = True
    Me.cboTipoDocumentoAcq.Enabled = True
    Me.txtNumeroDocumentoAcq.Enabled = True
    Me.txtDataDocumentoAcq.Enabled = True
    Me.txtTargaAutomezzo.Enabled = True
    Me.CDSocioFatt.Enabled = True
    
    
    txtIDOrdineCliente_Change
    txtIDLetteraIntento_Change
    fnGetMagazzinoCarico
    fnGetMagazzinoScarico
    
    'rif11 end

    If m_Document(m_Document.PrimaryKey).Value > 0 Then
       LINK_ORDINAMENTO = GET_LINK_ORDINAMENTO(m_DocumentsLink.TableName)
    Else
       LINK_ORDINAMENTO = 1
    End If
    
    SSTab1.Tab = 0


    If LINK_RIGA_CONFERIMENTO_DA_MOV > 0 Then
        If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
            m_DocumentsLink.MoveFirst
            While Not m_DocumentsLink.EOF
                If LINK_RIGA_CONFERIMENTO_DA_MOV = fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value) Then
                    Exit Sub
                End If
            m_DocumentsLink.MoveNext
            Wend
        End If
    End If

    If LINK_OGGETTO_ACQ > 0 Then
        Me.cboTipoDocumentoAcq.Enabled = False
        Me.txtNumeroDocumentoAcq.Enabled = False
        Me.txtDataDocumentoAcq.Enabled = False
        Me.cboVettore.Enabled = False
        Me.CDSocioFatt.Enabled = False
    Else
        Me.cboTipoDocumentoAcq.Enabled = True
        Me.txtNumeroDocumentoAcq.Enabled = True
        Me.txtDataDocumentoAcq.Enabled = True
        Me.cboVettore.Enabled = True
        Me.CDSocioFatt.Enabled = True
    End If
    
    If Me.chkPreConferimento.Value = vbChecked Then
        Me.cboVettore.Enabled = False
        Me.txtImportoTrasporto.Enabled = False
        Me.cboLuogoPresaMerce.Enabled = False
        Me.txtNumeroDocumento.Enabled = False
        Me.txtDataConferimento.Enabled = False
        
        Me.txtNumeroConferimento.Enabled = False
        Me.txtPrefissoConferimento.Enabled = False
        Me.cboTipoDocumentoAcq.Enabled = False
        Me.txtNumeroDocumentoAcq.Enabled = False
        Me.txtDataDocumentoAcq.Enabled = False
        Me.txtTargaAutomezzo.Enabled = False
        Me.CDSocioFatt.Enabled = False
    End If
    
    
    CREA_RECORDSET_LOTTO_IMBALLO_DELETE
    ADD_PESATURE_CONFERIMENTO fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    CONTROLLO_MOVIMENTAZIONE_CONFERIMENTO fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    m_Changed = False
    
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

    Set m_FormFields = Nothing
    Set m_Report = Nothing
    Set m_ActiveFilter = Nothing
    Set m_Document = Nothing
    Set m_Process = Nothing
    Set m_App = Nothing
    Set m_Semaphore = Nothing
    
    'cbcx
    'Set m_ExtendApplication = Nothing
End Sub



'rif8 begin

'**+
'Autore: Diamante s.p.a
'Data creazione: 03/11/00
'Autore ultima modifica:
'Data ultima modifica:
'
'Nome: m_DocumentsLink_OnReposition
'
'Parametri:
'
'Valori di ritorno:
'
'Funzionalità: Operazioni da effettuare al Reposition del sottodocumento.
'
'**/
Private Sub m_DocumentsLink_OnReposition()
Dim bValue As Boolean
Dim iIndex As Integer

On Error Resume Next
    
If AggiornamentoDocumento = 0 Then
    If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        'ControlloQuantitaDiCarico
        
        
        Me.CDArticolo.Load fnNotNull(m_DocumentsLink("IDArticolo").Value)
        Me.CDArticolo.Code = fnNotNull(m_DocumentsLink("CodiceArticolo").Value)
        Me.TxtArticolo.Text = fnNotNull(m_DocumentsLink("Articolo").Value)
        Link_UnitaDiMisura_Coop = fnNotNull(m_DocumentsLink("IDUnitaDiMisura").Value)
        Me.cboUM.WriteOn IIf(IsNull(m_DocumentsLink("IDUnitaDiMisuraDiamante").Value), 0, m_DocumentsLink("IDUnitaDiMisuraDiamante").Value)
        Me.txtColli.Value = fnNotNullN(m_DocumentsLink("Colli").Value)
        Me.txtPesoLordo.Value = fnNotNullN(m_DocumentsLink("PesoLordo").Value)
        Me.txtPesoNetto.Value = fnNotNullN(m_DocumentsLink("PesoNetto").Value)
        Me.txtTaraUnitaria.Value = fnNotNullN(m_DocumentsLink("TaraUnitaria").Value)
        Me.txtTara.Value = fnNotNullN(m_DocumentsLink("Tara").Value)
        Me.txtPezzi.Value = fnNotNullN(m_DocumentsLink("Pezzi").Value)
        Me.txtImportoUnitario.Value = fnNotNullN(m_DocumentsLink("ImportoUnitario").Value)
        Me.txtImpUniImb.Value = fnNotNullN(m_DocumentsLink("ImportoUnitarioImballo").Value)
        Me.txtQta_UM.Value = fnNotNullN(m_DocumentsLink("Qta_UM").Value)
        Me.CDCodiceImballo.Load fnNotNull(m_DocumentsLink("IDImballo").Value)
        Me.CDCodiceImballo.Code = fnNotNull(m_DocumentsLink("CodiceImballo").Value)
        Me.txtImballo.Text = fnNotNull(m_DocumentsLink("DescrizioneImballo").Value)
        Link_Movimento_Carico = fnNotNullN(m_DocumentsLink("IDMovimento_carico").Value)
        Link_Movimento_Scarico = fnNotNullN(m_DocumentsLink("IDMovimento_vendita").Value)
        Me.txtLottoDiConferimento.Text = fnNotNull(m_DocumentsLink("LottoDiConferimento").Value)
        Me.chkLottoChiuso.Value = IIf(IsNull(m_DocumentsLink("Chiuso").Value), 0, fnNormBoolean(m_DocumentsLink("Chiuso").Value))
        Me.txtLottoDiEntrata.Text = fnNotNull(m_DocumentsLink("CodiceLotto").Value)
        Me.txtNumeroLottoEntrata.Value = fnNotNullN(m_DocumentsLink("IDCodiceLotto").Value)
        Me.txtIDLottoCampagna.Value = fnNotNullN(m_DocumentsLink("IDLottoDiCampagna").Value)
        Me.txtDataSbloccoLotto.Value = fnNotNullN(m_DocumentsLink("DataSbloccoLottoCampagna").Value)
        Me.cboRegione.WriteOn fnNotNullN(m_DocumentsLink("IDRegione").Value)
        Me.cboNazione.WriteOn fnNotNullN(m_DocumentsLink("IDNazione").Value)
        Me.cboComune.WriteOn fnNotNullN(m_DocumentsLink("IDComune").Value)
        Me.cboProvincia.WriteOn fnNotNullN(m_DocumentsLink("IDProvincia").Value)
        Me.txtNomeMacchina.Text = fnNotNull(m_DocumentsLink("NomePC").Value)
        Me.txtUtenteMacchina.Text = fnNotNull(m_DocumentsLink("UtentePC").Value)
        Me.cboUtente.WriteOn fnNotNullN(m_DocumentsLink("IDUtente").Value)
        Me.txtCodiceUtente.Text = fnNotNull(m_DocumentsLink("CodiceUtente").Value)
        Me.CDPedana.Load fnNotNullN(m_DocumentsLink("IDArticoloPedana").Value)
        Me.txtTaraPedana.Value = fnNotNullN(m_DocumentsLink("TaraPedana").Value)
        Me.txtQuantitaPedana.Value = fnNotNullN(m_DocumentsLink("QuantitaPedana").Value)
        Me.txtTaraAutomezzo.Value = fnNotNullN(m_DocumentsLink("TaraMezzo").Value)
        Me.cboIVA.WriteOn fnNotNullN(m_DocumentsLink("IDIvaAcquisto").Value)
        Me.txtAliquotaIVA.Value = fnNotNullN(m_DocumentsLink("AliquotaIVA").Value)
        'Me.txtImportoUnitario.Value = fnNotNullN(m_DocumentsLink("ImportoUnitario").Value)
        Me.txtImponibileRiga.Value = fnNotNullN(m_DocumentsLink("TotaleImponibileRiga").Value)
        Me.txtImpostaRiga.Value = fnNotNullN(m_DocumentsLink("TotaleImpostaRiga").Value)
        Me.txtTotaleRiga.Value = fnNotNullN(m_DocumentsLink("TotaleLordoRiga").Value)

        Me.txtOraArrivo.Text = fnNotNull(m_DocumentsLink("OraArrivoMerce").Value)

        Me.cboTipoLavorazione.WriteOn fnNotNullN(m_DocumentsLink("IDRV_POTipoLavorazione").Value)
        Me.cboLiquidato.WriteOn fnNotNullN(m_DocumentsLink("IDRV_POTipoConfLiquidazione").Value)
        Me.chkPrezzoMedio.Value = Abs(fnNotNullN(m_DocumentsLink("PrezzoMedio").Value))
        
        Me.txtIDLottoImballo.Value = fnNotNullN(m_DocumentsLink("IDRV_POLottoImballo").Value)
        Me.chkTracciaImballo.Value = fnNotNullN(m_DocumentsLink("TracciaImballo").Value)
        Me.txtQtaPresunta.Value = fnNotNullN(m_DocumentsLink("QuantitaPresunta").Value)
        'Me.txtImpUniImb.Value = fnNotNullN(m_DocumentsLink("ImportoUnitarioImballo").Value)
        Me.txtNoteAgg.Text = fnNotNull(m_DocumentsLink("AnnotazioniAggiuntive").Value)
        Me.chkLottoImbEsclusivo.Value = fnNotNullN(m_DocumentsLink("LottoImballoUtilizzoEsclusivo").Value)
        bValue = True
        
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
       
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.CDArticolo.Load 0
        Me.CDArticolo.Code = ""
        Me.TxtArticolo.Text = ""
        Link_UnitaDiMisura_Coop = 0
        Me.cboUM.WriteOn 0
        Me.txtColli.Value = 0
        Me.txtPesoLordo.Value = 0
        Me.txtPesoNetto.Value = 0
        Me.txtTara.Value = 0
        Me.txtTaraUnitaria.Value = 0
        Me.txtPezzi.Value = 0
        Me.txtQta_UM.Value = 0
        Me.CDCodiceImballo.Load 0
        Me.CDCodiceImballo.Code = ""
        Me.txtImballo.Text = ""
        Link_Movimento_Carico = 0
        Link_Movimento_Scarico = 0
        Me.txtLottoDiConferimento.Text = ""
        Me.chkLottoChiuso.Value = 0
        Me.txtLottoDiEntrata.Text = ""
        Me.txtIDLottoCampagna.Value = 0
        Me.cboRegione.WriteOn 0
        Me.cboNazione.WriteOn 0
        Me.cboComune.WriteOn 0
        Me.cboProvincia.WriteOn 0
        Me.txtNomeMacchina.Text = ""
        Me.txtUtenteMacchina.Text = ""
        Me.cboUtente.WriteOn 0
        Me.txtCodiceUtente.Text = ""
        Me.CDPedana.Load 0
        Me.txtTaraPedana.Value = 0
        Me.txtQuantitaPedana.Value = 0

        Me.cboIVA.WriteOn 0
        Me.txtAliquotaIVA.Value = 0
        Me.txtImportoUnitario.Value = 0
        Me.txtImponibileRiga.Value = 0
        Me.txtImpostaRiga.Value = 0
        Me.txtTotaleRiga.Value = 0
        Me.txtOraArrivo.Text = ""


        Me.cboTipoLavorazione.WriteOn 0
        Me.cboLiquidato.WriteOn 0
        Me.chkPrezzoMedio.Value = 0
        Me.txtTaraAutomezzo.Value = 0
        Me.txtIDLottoImballo.Value = 0
        Me.chkTracciaImballo.Value = 0
        
        Me.txtImpUniImb.Value = 0
        Me.txtQtaPresunta.Value = 0
        Me.txtNoteAgg.Text = ""
        Me.chkLottoImbEsclusivo.Value = 0
        bValue = False
    End If
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
        If m_DocumentsLink("IDRV_POCaricoMerceRighe") < 0 Then
            Me.CDArticolo.Enabled = bValue
            Me.TxtArticolo.Enabled = bValue
        Else
            Me.CDArticolo.Enabled = False
            Me.TxtArticolo.Enabled = False
        End If
        Me.txtTaraUnitaria.Enabled = bValue
        Me.txtColli.Enabled = bValue
        Me.txtPesoLordo.Enabled = bValue
        Me.txtPesoNetto.Enabled = bValue
        Me.txtTara.Enabled = bValue
        Me.txtPezzi.Enabled = bValue
        'Me.txtQta_UM.Enabled = bValue
        Me.CDCodiceImballo.Enabled = bValue
        Me.txtImballo.Enabled = bValue
        Me.txtLottoDiConferimento.Enabled = bValue
        'Me.cboUM.Enabled = bValue
        Me.chkLottoChiuso.Enabled = bValue
        Me.lblArticolo.Enabled = bValue
        Me.lblImballo.Enabled = bValue
        Me.LabelLink1.Enabled = bValue
        Me.txtLottoDiEntrata.Enabled = bValue
        Me.txtLottoDiConferimento.Enabled = bValue
        Me.cmdSelezionaLottoCampagna.Enabled = bValue
        Me.cmdTracciabilitaLottoCampagna.Enabled = bValue
        Me.cmdArticoliDerivati.Enabled = bValue
        Me.cboRegione.Enabled = bValue
        Me.cboNazione.Enabled = bValue
        Me.cboComune.Enabled = bValue
        Me.CDPedana.Enabled = bValue
        Me.txtTaraPedana.Enabled = bValue
        Me.txtQuantitaPedana.Enabled = bValue
        Me.cboIVA.Enabled = bValue
'        Me.txtAliquotaIVA.Enabled = bValue
        Me.txtImportoUnitario.Enabled = bValue
'        Me.txtImponibileRiga.Enabled = bValue
'        Me.txtImpostaRiga.Enabled = bValue
'        Me.txtTotaleRiga.Enabled = bValue
        Me.txtOraArrivo.Enabled = bValue

        Me.cboTipoLavorazione.Enabled = bValue
        Me.cboLiquidato.Enabled = bValue
        Me.chkPrezzoMedio.Enabled = bValue

        Me.txtIDLottoImballo.Enabled = bValue
        Me.chkTracciaImballo.Enabled = bValue
        Me.txtTaraAutomezzo.Enabled = bValue
        
        Me.txtImpUniImb.Enabled = bValue
        Me.txtQtaPresunta.Enabled = bValue
        Me.txtNoteAgg.Enabled = bValue
        Me.chkLottoImbEsclusivo.Enabled = bValue
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovo.Enabled = True
        Me.cmdSalva.Enabled = bValue
        Me.cmdElimina.Enabled = bValue
        'Me.cmdRiepilogo.Enabled = bValue
      
        
        If Me.CDArticolo.KeyFieldID > 0 Then
            bVariazioneDettaglio = True
        Else
            bVariazioneDettaglio = False
        End If
        
        If Me.txtIDLottoCampagna.Value > 0 Then
            'If GET_ESISTENZA_BIO = True Then
            GET_DATI_LOTTO_CAMPAGNA Me.txtIDLottoCampagna.Value
            'End If
        Else
            Me.txtVarietà.Text = ""
            Me.txtFamiglia.Text = ""
            Me.txtTipoProduzione.Text = ""
        End If
        
    If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then

        Me.txtQtaVenduta.Value = GET_RIEPILOGO_QUANTITA_VENDUTO(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value))
        Me.txtQtaAssegnata.Value = GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value)) + GET_RIEPILOGO_QUANTITA_PROCESSO(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value))
        Me.txtQtaQuadrata.Value = GET_RIEPILOGO_QUANTITA_LAVORAZIONE(fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value))
        Me.txtQtaDifferenza.Value = Me.txtQta_UM.Value - (Me.txtQtaQuadrata.Value + Me.txtQtaVenduta.Value)
        Me.txtQtaDifferenzaLavorazione.Value = Me.txtQta_UM.Value - (Me.txtQtaQuadrata.Value + Me.txtQtaAssegnata.Value)
   Else
        Me.txtQtaVenduta.Value = 0
        Me.txtQtaAssegnata.Value = 0
        Me.txtQtaQuadrata.Value = 0
        Me.txtQtaDifferenza.Value = 0
        Me.txtQtaDifferenzaLavorazione.Value = 0
   End If
       
       
    If txtIDLottoImballo.Value > 0 Then
        Me.cmdRiepLottoImballoConf.Enabled = True
        Me.chkLottoImbEsclusivo.Enabled = True
    Else
        Me.cmdRiepLottoImballoConf.Enabled = False
        Me.chkLottoImbEsclusivo.Enabled = False
    End If
    
    GET_RIEPILOGO_PESI
    
End If
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
Private Sub ParametroSocio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaAnagrafica FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoSocio = fnNotNullN(rs!IDCategoriaAnagrafica)
Else
    Link_TipoSocio = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroGrezzo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoGrezzo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Link_TipoGrezzo = rs!IDTipoGrezzo
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
    Link_TipoLavorato = rs!IDTipoLavorato
Else
    Link_TipoLavorato = 0
End If

rs.CloseResultset
Set rs = Nothing
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

Private Sub ParametriDiDefault(DataDocumento As String)
    
    Link_Esercizio = fnGetEsercizio(DataDocumento)
    Link_PeriodoIVA = fnGetPeriodoIVA(DataDocumento)
    Link_Sezionale = fnGetSezionale(IDDocumento)
    
    Link_Magazzino_Conferimento = fnGetParametriMagazzino("IDMagazzino_Carico")
    Link_Magazzino_Vendita = fnGetParametriMagazzino("IDMagazzino_Vendita")
    Link_Causale_MagCar_Conf = fnGetParametriMagazzino("IDCausale_Carico_Mag_Carico")
    Link_Causale_MagScar_Conf = fnGetParametriMagazzino("IDCausale_Scarico_Mag_Carico")
    Link_Causale_MagCar_Vend = fnGetParametriMagazzino("IDCausale_Carico_Mag_Vendita")
    Link_Causale_MagScar_Vend = fnGetParametriMagazzino("IDCausale_Scarico_Mag_vendita")
    Numero_Decimali_Pesi = fnGetParametriMagazzino("IDRV_POTipoDecimaliPesiConferimento") - 1
    LINK_FUNZIONE_DDT = fnGetParametriMagazzino("IDFunzioneDDTAcqConf")
    LINK_FUNZIONE_FA = fnGetParametriMagazzino("IDFunzioneFAAcqConf")
    
    If Numero_Decimali_Pesi < 0 Then
        Numero_Decimali_Pesi = 2
    End If
    
End Sub


Public Function fnGetSezionale(Link_DocumentoCoop) As Long
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT RV_POSezionalePerDocumento.IDSezionale "
    sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
    sSQL = sSQL & "RV_POSezionalePerDocumento ON RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POSezionalePerDocumento.IDRV_POSchemaCoop "
    sSQL = sSQL & "WHERE ((IDFiliale=" & m_App.Branch & ") "
    sSQL = sSQL & "AND (IDDocumentoCoop=" & Link_DocumentoCoop & ") "
    sSQL = sSQL & "AND (Predefinito=" & fnNormBoolean(1) & "))"
    
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        fnGetSezionale = rsEse!IDSezionale
    Else
        fnGetSezionale = 0
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing
End Function
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
Public Function fnGetNumeroDocumento()
    Dim rsEse As DmtOleDbLib.adoResultset
    Dim sSQL As String
    
    sSQL = "SELECT ProgressivoDisponibile FROM ProgressivoSezionale "
    sSQL = sSQL & "WHERE ((IDPeriodoIva=" & Link_PeriodoIVA & ") "
    sSQL = sSQL & "AND (IDSezionale=" & Link_Sezionale & ") "
    sSQL = sSQL & "AND (IDTipoModulo=1))"
    
    
    Set rsEse = Cn.OpenResultset(sSQL)
    
    If rsEse.EOF = False Then
        EsistenzaValoreSezionale = True
        NumeroDocumentoDisponibile = rsEse!ProgressivoDisponibile
        fnGetNumeroDocumento = rsEse!ProgressivoDisponibile
    Else
        EsistenzaValoreSezionale = False
        NumeroDocumentoDisponibile = 1
        fnGetNumeroDocumento = 1
    End If
    
    rsEse.CloseResultset
    Set rsEse = Nothing


End Function




Private Sub m_DocumentsLink1_OnReposition()
Dim bValue As Boolean
Dim iIndex As Integer

On Error Resume Next
    

    If Not (m_DocumentsLink1.BOF And m_DocumentsLink1.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        'ControlloQuantitaDiCarico
        
        Me.CDImballoGestione.Load fnNotNullN(m_DocumentsLink1("IDArticolo").Value)
        Me.chkTracciaImballoGest.Value = fnNotNullN(m_DocumentsLink1("TracciaImballo").Value)
        Me.txtIDLottoImballoGest.Value = fnNotNullN(m_DocumentsLink1("IDRV_POLottoImballo").Value)
        Me.cboTipoProcessoImballo.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POTipoProcessoCoop").Value)
        Me.txtQuantitaImballo.Value = fnNotNullN(m_DocumentsLink1("Quantita").Value)
        Me.chkConfermaDaUtente.Value = fnNotNullN(m_DocumentsLink1("ConfermaDaUtente").Value)
        Me.txtImpUniAltriImb.Value = fnNotNullN(m_DocumentsLink1("Importo").Value)
        Me.txtLottoImballoRifEsterno.Text = fnNotNull(m_DocumentsLink1("RiferimentoEsterno").Value)
        bValue = True
        
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.CDImballoGestione.Load 0
        Me.chkTracciaImballoGest.Value = 0
        Me.txtIDLottoImballoGest.Value = 0
        Me.cboTipoProcessoImballo.WriteOn 0
        Me.txtQuantitaImballo.Value = 0
        Me.chkConfermaDaUtente.Value = 0
        Me.txtImpUniAltriImb.Value = 0
        Me.txtLottoImballoRifEsterno.Text = ""
        bValue = False
    End If
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
        Me.CDImballoGestione.Enabled = bValue
        Me.cboTipoProcessoImballo.Enabled = bValue
        Me.txtQuantitaImballo.Enabled = bValue
        Me.chkTracciaImballoGest.Enabled = bValue
        Me.txtIDLottoImballoGest.Enabled = bValue
        Me.chkConfermaDaUtente.Enabled = bValue
        Me.txtImpUniAltriImb.Enabled = bValue
        Me.txtLottoImballoRifEsterno.Enabled = bValue
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
        Me.cmdNuovoImballo.Enabled = True
        Me.cmdSalvaImballo.Enabled = bValue
        Me.cmdEliminaImballo.Enabled = bValue
      
        
        Me.txtGiacLottoImballo.Value = 0
        Me.txtDispLottoImballo.Value = 0
        Me.cmdLottoImballo.Enabled = False
        Me.cmdRiepilogoLottoImballo.Enabled = False

        CREA_RECORDSET_LOTTI_IMBALLI Me.CDImballoGestione.KeyFieldID, fnGetTipoOggetto("RV_POGestImbConf"), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value)
        
        
        If Me.chkTracciaImballoGest.Value = vbChecked Then
            If Me.cboTipoProcessoImballo.CurrentID = 2 Then
                Me.txtGiacLottoImballo.Value = GET_GIACENZA_LOTTO_IMBALLO(Me.CDImballoGestione.KeyFieldID)
                
                Me.txtDispLottoImballo.Value = Me.txtGiacLottoImballo.Value + GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(0, fnGetTipoOggetto("RV_POGestImbConf"), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
            
                If Me.txtIDLottoImballoGest.Value > 0 Then
                    Me.txtDispLottoImballo.Value = Me.txtDispLottoImballo.Value - GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(Me.txtIDLottoImballoGest.Value, fnGetTipoOggetto("RV_POGestImbConf"), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value), fnNotNullN(m_DocumentsLink1(m_DocumentsLink1.PrimaryKey).Value))
                Else
                    Me.cmdLottoImballo.Enabled = True
                End If
            End If
            If Me.cboTipoProcessoImballo.CurrentID = 1 Then
                If Me.txtIDLottoImballoGest.Value > 0 Then
                    Me.cmdRiepilogoLottoImballo.Enabled = True
                End If
            End If
        End If
End Sub

Private Sub m_DocumentsLink2_OnReposition()
Dim bValue As Boolean
Dim iIndex As Integer

On Error Resume Next
    

    If Not (m_DocumentsLink2.BOF And m_DocumentsLink2.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        'ControlloQuantitaDiCarico
        
        Me.CDArticoliAddebiti.Load fnNotNullN(m_DocumentsLink2("IDArticolo").Value)
        Me.cboTipoTrattenutaAggiuntiva.WriteOn fnNotNullN(m_DocumentsLink2("IDRV_POTipoTrattenutaAggiuntiva").Value)
        Me.txtQtaAddebiti.Value = fnNotNullN(m_DocumentsLink2("Quantita").Value)
        Me.txtImpUniAddebiti.Value = fnNotNullN(m_DocumentsLink2("ImportoUnitario").Value)
        Me.cboIvaAddebiti.WriteOn fnNotNullN(m_DocumentsLink2("IDIva").Value)
        Me.txtImponibileAddebiti.Value = fnNotNullN(m_DocumentsLink2("TotaleRigaNettoIva").Value)
        Me.txtImpostaAddebiti.Value = fnNotNullN(m_DocumentsLink2("ImpostaRiga").Value)
        Me.txtTotaleRigaAddebiti.Value = fnNotNullN(m_DocumentsLink2("TotaleRigaLordoIva").Value)
        
        bValue = True
        
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        Me.CDArticoliAddebiti.Load 0
        Me.cboTipoTrattenutaAggiuntiva.WriteOn 0
        Me.txtQtaAddebiti.Value = 0
        Me.txtImpUniAddebiti.Value = 0
        Me.cboIvaAddebiti.WriteOn 0
        Me.txtImponibileAddebiti.Value = 0
        Me.txtImpostaAddebiti.Value = 0
        Me.txtTotaleRigaAddebiti.Value = 0

        bValue = False
    End If
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento

        Me.CDArticoliAddebiti.Enabled = bValue
        Me.cboTipoTrattenutaAggiuntiva.Enabled = bValue
        Me.txtQtaAddebiti.Enabled = bValue
        Me.txtImpUniAddebiti.Enabled = bValue
        Me.cboIvaAddebiti.Enabled = bValue

    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
        Me.cmdNuovoAddebiti.Enabled = True
        Me.cmdSalvaAddebiti.Enabled = bValue
        Me.cmdEliminaAddebiti.Enabled = bValue
End Sub

Private Sub txtAnnotazioni_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub TxtArticolo_LostFocus()
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    Dim I As Long
    
 If Me.CDArticolo.Code = "" Then
    If Me.TxtArticolo.Text <> "" Then
        sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo, IDUnitaDiMisuraAcquisto "
        sSQL = sSQL & "FROM Articolo "
        sSQL = sSQL & "WHERE (((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL)) "

        sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & ") "
        sSQL = sSQL & "AND (Articolo LIKE " & fnNormString(Me.TxtArticolo.Text & "%") & "))"
    
        Set rs = Cn.OpenResultset(sSQL, adOpenKeyset)
    
        I = 0
        If rs.EOF = False Then
        
            While Not rs.EOF
                I = I + 1
                If I > 1 Then
                    rs.MoveLast
                End If
            rs.MoveNext
            Wend
        End If
    
        If I = 1 Then
            sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo, IDUnitaDiMisuraAcquisto "
            sSQL = sSQL & "FROM Articolo "
            sSQL = sSQL & "WHERE (((IDTipoProdotto <> " & Link_TipoImballo & ") OR (IDTipoProdotto IS NULL)) "

            sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & ") "
            sSQL = sSQL & "AND (Articolo LIKE " & fnNormString(Me.TxtArticolo.Text & "%") & "))"

            Set rs = Cn.OpenResultset(sSQL)
                Me.CDArticolo.Load rs!IDArticolo
                Me.CDArticolo.Code = rs!CodiceArticolo
                Me.TxtArticolo.Text = rs!Articolo
                Link_UnitaDiMisura_Coop = fnGetUMCoop(rs!IDUnitaDiMisuraAcquisto)
                Me.txtColli.SetFocus
            rs.CloseResultset
            Set rs = Nothing
        End If
        If I = 0 Then
            MsgBox "Non sono stati trovati i dati per i filtri impostati", vbInformation, "Articolo"
        End If
        If I > 1 Then
            rs.CloseResultset
            Set rs = Nothing
        
            lblArticolo_Click
            Me.txtColli.SetFocus
        End If
        
    End If
End If
    
End Sub

Private Sub txtColli_Change()
    If Link_UnitaDiMisura_Coop = 1 Then
        Me.txtQta_UM.Value = Me.txtColli.Value
    End If
    
    
End Sub

Private Sub txtColli_LostFocus()
        
    Me.txtTara.Value = ((Me.txtColli.Value * Me.txtTaraUnitaria.Value) + (Me.txtTaraPedana.Value * Me.txtQuantitaPedana.Value)) + Me.txtTaraAutomezzo.Value
    
    If ATTIVA_CALCOLO_PESO_LORDO = 1 Then
        If TIPO_PESO_ARTICOLO <= 1 Then
            If PESO_LORDO_ARTICOLO > 0 Then
                Me.txtPesoLordo.Value = PESO_LORDO_ARTICOLO * Me.txtColli.Value
            End If
        Else
            If PESO_LORDO_ARTICOLO > 0 Then
                Me.txtPesoNetto.Value = PESO_LORDO_ARTICOLO * Me.txtColli.Value
                Me.txtPesoLordo.Value = Me.txtPesoNetto.Value + Me.txtTara.Value
            End If
        End If
        

    End If
    
    If QUANTITA_PER_COLLO >= 1 Then
        Me.txtPezzi.Value = Me.txtColli.Value * QUANTITA_PER_COLLO
    End If
    
    CalcoloPesoNetto 1

End Sub

Private Sub txtDataConferimento_Change()



    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataConferimento_LostFocus()
On Error Resume Next
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then
        Link_Esercizio = fnGetEsercizio(Me.txtDataConferimento.Text)
        Me.cboEsercizio.WriteOn Link_Esercizio
        Link_PeriodoIVA = fnGetPeriodoIVA(Me.txtDataConferimento.Text)
        Me.txtNumeroDocumento.Value = fnGetNumeroDocumento
        Me.txtNumeroConferimento.Value = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "NumeroConferimento")
        Me.txtPrefissoConferimento.Text = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "PrefissoNumeroConferimento")
    Else
        If Me.txtDataConferimento.Value > 0 Then
            Link_Esercizio = fnGetEsercizio(Me.txtDataConferimento.Text)
            Me.cboEsercizio.WriteOn Link_Esercizio
            Link_PeriodoIVA = fnGetPeriodoIVA(Me.txtDataConferimento.Text)
            If Link_Esercizio <> fnNotNullN(m_Document("IDEsercizio").Value) Then
                Me.txtNumeroDocumento.Value = fnGetNumeroDocumento
                Me.txtNumeroConferimento.Value = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "NumeroConferimento")
                Me.txtPrefissoConferimento.Text = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "PrefissoNumeroConferimento")
            Else
                Me.txtNumeroDocumento.Value = fnNotNullN(m_Document("NumeroDocumento").Value)
                Me.txtNumeroConferimento.Value = fnNotNullN(m_Document("NumeroDocumentoSocio").Value)
                Me.txtPrefissoConferimento.Text = fnNotNull(m_Document("PrefissoNumeroConferimento").Value)
            End If
        End If
    End If
End Sub

Private Sub txtDataDocumento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataDocumentoAcq_Change()
    If Not (BrwMain.Visible) Then Change
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
    
    'Me.lblLetteraIntento.ToolTipText = ""
Else
    Me.txtNLetteraIntento.Text = fnNotNull(rs!Numero)
    Me.txtDataLetteraIntento.Value = fnNotNullN(rs!Data)

    'Me.lblLetteraIntento.ToolTipText = "Prot. N° " & fnNotNull(rs!NumeroCliFor) & " del " & fnNotNull(rs!DataEmissione)
End If

rs.CloseResultset
Set rs = Nothing

If Not (BrwMain.Visible) Then Change
Exit Sub
ERR_txtIDLetteraIntento_Change:
    MsgBox Err.Description, vbCritical, "txtIDLetteraIntento_Change"
End Sub


Private Sub txtIDLottoImballo_Change()
    Me.txtLottoImballo.Text = GET_DESCRIZIONE_LOTTO_IMBALLO(Me.txtIDLottoImballo.Value)
    If Me.txtIDLottoImballo.Value = 0 Then
        Me.cmdRiepLottoImballoConf.Enabled = False
    Else
        Me.cmdRiepLottoImballoConf.Enabled = True
    End If
End Sub

Private Sub txtIDLottoImballoGest_Change()
    Me.txtLottoImballoGest.Text = GET_DESCRIZIONE_LOTTO_IMBALLO(Me.txtIDLottoImballoGest.Value)

End Sub

Private Sub txtImballo_LostFocus()
    Dim rs As DmtOleDbLib.adoResultset
    Dim sSQL As String
    Dim I As Long
    
 If Me.CDCodiceImballo.Code = "" Then
    If Me.txtImballo.Text <> "" Then
        sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
        sSQL = sSQL & "FROM Articolo "
        sSQL = sSQL & "WHERE ((IDTipoProdotto = " & Link_TipoImballo & ") "
        sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & ") "
        sSQL = sSQL & "AND (Articolo LIKE " & fnNormString(Me.txtImballo.Text & "%") & "))"
    
        Set rs = Cn.OpenResultset(sSQL, adOpenKeyset)
    
        I = 0
        If rs.EOF = False Then
        
            While Not rs.EOF
                I = I + 1
                If I > 1 Then
                    rs.MoveLast
                End If
            rs.MoveNext
            Wend
        End If
    
        If I = 1 Then
        sSQL = "SELECT IDArticolo, Articolo, CodiceArticolo "
        sSQL = sSQL & "FROM Articolo "
        sSQL = sSQL & "WHERE ((IDTipoProdotto = " & Link_TipoImballo & ") "
        sSQL = sSQL & "AND (IDAzienda=" & m_App.IDFirm & ") "
        sSQL = sSQL & "AND (Articolo LIKE " & fnNormString(Me.txtImballo.Text & "%") & "))"

            Set rs = Cn.OpenResultset(sSQL)
                Me.CDCodiceImballo.Load rs!IDArticolo
                Me.CDCodiceImballo.Code = rs!CodiceArticolo
                Me.txtImballo.Text = rs!Articolo
                Me.txtTaraUnitaria.Value = fnGetTaraImballo(Me.CDCodiceImballo.KeyFieldID)
            rs.CloseResultset
            Set rs = Nothing
            
            Me.txtLottoDiConferimento.SetFocus
        End If
        If I = 0 Then
            MsgBox "Non sono stati trovati i dati per i filtri impostati", vbInformation, "Articolo"
        End If
        If I > 1 Then
            rs.CloseResultset
            Set rs = Nothing
        
            lblImballo_Click
            
        End If
        
    End If
End If
End Sub



Private Sub txtImportoTrasporto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtImportoUnitario_LostFocus()
    fnTotaleRiga
End Sub

Private Sub txtImpUniAddebiti_LostFocus()
    GET_TOTALI_RIGA_ALTRI_ADDEBITI
End Sub

Private Sub txtImpUniImb_LostFocus()
fnTotaleRiga
End Sub

Private Sub txtLottoDiConferimento_LostFocus()
Dim Link_Lotto As Long
Dim Link_LottoCampagna As Long

If (Me.CDArticolo.KeyFieldID > 0) Then
    If (m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value <= 0) Then
        If Me.txtNumeroLottoEntrata.Value = 0 Then
            Link_Lotto = GET_NUMERO_LOTTO
            Me.txtNumeroLottoEntrata.Value = Link_Lotto
        Else
            Link_Lotto = Me.txtNumeroLottoEntrata.Value
        End If
        Me.txtLottoDiEntrata.Text = fnGetCodiceLotto(1, 1, Me.txtLottoDiEntrata.Text, Link_Lotto)
    End If
End If

If m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value <= 0 Then
    If Len(Trim(Me.txtLottoDiConferimento.Text)) > 0 Then
        'If GET_ESISTENZA_BIO = True Then
            Link_LottoCampagna = GET_LINK_LOTTO_DI_CAMPAGNA
            If Link_LottoCampagna = 0 Then
                MsgBox "ATTENZIONE!!!!" & vbCrLf & "Non è stato possibile reperire le informazioni del codice lotto di campagna per questo socio"
                Me.txtIDLottoCampagna = 0
            Else
                Me.txtIDLottoCampagna = Link_LottoCampagna
                GET_DATI_LOTTO_CAMPAGNA Me.txtIDLottoCampagna.Value
            End If
        'End If
    End If
End If

End Sub
Private Function GET_ESISTENZA_BIO() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POProgramma FROM RV_POProgramma "
sSQL = sSQL & "WHERE IDRV_POProgramma=2"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_BIO = False
Else
    GET_ESISTENZA_BIO = True
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_LINK_LOTTO_DI_CAMPAGNA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_PO01_LottoCampagna FROM RV_PO01_LottoCampagna "
sSQL = sSQL & "WHERE IDSocio=" & Me.CDSocio.KeyFieldID
sSQL = sSQL & " AND CodiceLotto=" & fnNormString(Me.txtLottoDiConferimento.Text)
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_LOTTO_DI_CAMPAGNA = 0
Else
    GET_LINK_LOTTO_DI_CAMPAGNA = fnNotNullN(rs!IDRV_PO01_LottoCampagna)
End If

rs.CloseResultset
Set rs = Nothing
End Function


Private Sub txtNoteAgg_LostFocus()
Dim Link_Lotto As Long

    If (Len(txtNoteAgg.Text) > 0) Then
        If (m_DocumentsLink(m_DocumentsLink.PrimaryKey).Value <= 0) Then
            If Me.txtNumeroLottoEntrata.Value = 0 Then
                Link_Lotto = GET_NUMERO_LOTTO
                Me.txtNumeroLottoEntrata.Value = Link_Lotto
            Else
                Link_Lotto = Me.txtNumeroLottoEntrata.Value
            End If
            Me.txtLottoDiEntrata.Text = fnGetCodiceLotto(1, 1, Me.txtLottoDiEntrata.Text, Link_Lotto)
        End If
    End If
End Sub

Private Sub txtNumeroConferimento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumeroDocumento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumeroDocumentoAcq_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtOraArrivo_DblClick()
    frmOraArrivoMerce.Show vbModal
End Sub

Private Sub txtPesoLordo_Change()
    If Link_UnitaDiMisura_Coop = 2 Then
        Me.txtQta_UM.Value = Me.txtPesoLordo.Value
    End If
    
    
End Sub

Private Sub txtPesoLordo_LostFocus()
    CalcoloPesoNetto 1
End Sub

Private Sub txtPesoNetto_Change()
    If Link_UnitaDiMisura_Coop = 3 Then
        Me.txtQta_UM.Value = Me.txtPesoNetto.Value
    End If
    
    
End Sub

Private Sub txtPesoNetto_LostFocus()
    CalcoloPesoNetto 2
End Sub

Private Sub txtPezzi_Change()
    If Link_UnitaDiMisura_Coop = 5 Then
        Me.txtQta_UM.Value = Me.txtPezzi.Value
    End If
    
    
End Sub





Private Sub AggiornamentoProgressivoSezionale()
    Dim sSQL As String
    Dim rsCtrl As DmtOleDbLib.adoResultset
    If EsistenzaValoreSezionale = True Then
        If NumeroDocumentoDisponibile = Me.txtNumeroDocumento.Value Then
            sSQL = "UPDATE ProgressivoSezionale SET "
            sSQL = sSQL & "ProgressivoDisponibile=" & Me.txtNumeroDocumento.Value + 1 & " "
            sSQL = sSQL & "WHERE ((IDPeriodoIva=" & Link_PeriodoIVA & ") "
            sSQL = sSQL & "AND (IDSezionale=" & Link_Sezionale & ") "
            sSQL = sSQL & "AND (IDTipoModulo=1))"
        ElseIf NumeroDocumentoDisponibile < Me.txtNumeroDocumento.Value Then
            sSQL = "UPDATE ProgressivoSezionale SET "
            sSQL = sSQL & "ProgressivoDisponibile=" & Me.txtNumeroDocumento.Value + 1 & " "
            sSQL = sSQL & "WHERE ((IDPeriodoIva=" & Link_PeriodoIVA & ") "
            sSQL = sSQL & "AND (IDSezionale=" & Link_Sezionale & ") "
            sSQL = sSQL & "AND (IDTipoModulo=1))"
        End If
            
    Else
        sSQL = "INSERT INTO ProgressivoSezionale (IDProgressivoSezionale, IDPeriodoIva, IDTipoModulo, IDSezionale, "
        sSQL = sSQL & "ProgressivoDisponibile, IDUtenteUltimaVariazione, VirtualDelete, DataUltimaVariazione) "
        sSQL = sSQL & "VALUES ("
        sSQL = sSQL & fnGetNewKey("ProgressivoSezionale", "IDProgressivoSezionale") & ", "
        sSQL = sSQL & Link_PeriodoIVA & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & Link_Sezionale & ", "
        sSQL = sSQL & Me.txtNumeroDocumento.Value + 1 & ", "
        sSQL = sSQL & 1 & ", "
        sSQL = sSQL & 0 & ", "
        sSQL = sSQL & fnNormDate(Date) & ")"
    End If
    
    If Len(sSQL) > 0 Then
        Cn.Execute sSQL
    End If
    
End Sub
Private Function InserimentoOggetto() As Boolean
    
    Dim sSQL As String
    
    sSQL = "INSERT Oggetto (IDOggetto, IDTipoOggetto, IDAzienda, IDAttivitaAzienda, "
    sSQL = sSQL & "IDSezionale, Oggetto, DataEmissione, Numero, DataUltimaVariazione, "
    sSQL = sSQL & "IDUtenteUltimaVariazione, VirtualDelete, IDFunzione) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & Link_Oggetto & ", "
    sSQL = sSQL & fnGetTipoOggetto & ", "
    sSQL = sSQL & m_App.IDFirm & ", "
    sSQL = sSQL & 6 & ", "
    sSQL = sSQL & Link_Sezionale & ", "
    sSQL = sSQL & fnNormString("Conferimento merce") & ", "
    sSQL = sSQL & fnNormDate(Me.txtDataDocumento.Text) & ", "
    sSQL = sSQL & fnNormString(Me.txtNumeroDocumento.Value) & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & m_App.IDUser & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & m_App.FunctionID & ")"
    
    Cn.Execute sSQL
    
    InserimentoOggetto = True
    
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
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Public Sub fnEliminaOggetto()
    'Dim sSQL As String
    
    'sSQL = "DELETE FROM Oggetto WHERE IDOggetto=" & m_Document("IDOggetto").Value
    
    'Cn.Execute sSQL
End Sub
Private Function fnGetUMCoop(Link_UMAcq As Long) As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
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

Private Sub txtPezzi_LostFocus()
    GET_RIEPILOGO_PESI
End Sub

Private Sub txtPrefissoConferimento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtQta_UM_Change()
On Error GoTo ERR_txtQta_UM_Change
    fnTotaleRiga
    On Error Resume Next
    If fnNotNullN(m_DocumentsLink(m_DocumentsLink.PrimaryKey) <= 0) Then
        Me.txtQtaPresunta.Value = Me.txtQta_UM.Value
    End If
    
Exit Sub
ERR_txtQta_UM_Change:
    MsgBox Err.Description, vbCritical, "txtQta_UM_Change"
End Sub

Private Sub txtQtaAddebiti_LostFocus()
    GET_TOTALI_RIGA_ALTRI_ADDEBITI
End Sub

Private Sub txtQuantitaPedana_Change()
    CalcoloPesoNetto 1
    
End Sub

Private Sub txtTara_Change()
    If Link_UnitaDiMisura_Coop = 4 Then
        Me.txtQta_UM.Value = Me.txtTara.Value
    End If
End Sub
Private Function fnGetCodiceLotto(TipoLotto As Integer, TipoStringaLotto As Integer, StringaLotto As String, Link_LottoArticolo As Long) As String
On Error GoTo ERR_fnGetCodiceLotto
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim Codice As String
Dim I As Integer
Dim PosizioneCodice As Integer
Dim Stringa As String
Dim StringaElaborata As String
Dim SenzaIndentificativo As Boolean

sSQL = "SELECT RV_POLottoCostruzioneRighe.IDRV_POLottoComp, RV_POLottoCostruzioneRighe.Posizione, RV_POLottoCostruzioneRighe.Lunghezza, "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.Testo, RV_POLottoCostruzioneTesta.PosVendita, RV_POLottoCostruzioneTesta.PosConferimento, "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.SXDX, RV_POLottoCostruzioneTesta.SenzaCodiceRiferimento_Conf "
sSQL = sSQL & "FROM RV_POLottoCostruzioneRighe INNER JOIN "
sSQL = sSQL & "RV_POLottoCostruzioneTesta ON "
sSQL = sSQL & "RV_POLottoCostruzioneRighe.IDRV_POLottoCostruzioneTesta = RV_POLottoCostruzioneTesta.IDRV_POLottoCostruzioneTesta "
sSQL = sSQL & "WHERE RV_POLottoCostruzioneTesta.IDFiliale=" & TheApp.Branch
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoLotto=" & TipoLotto
sSQL = sSQL & " AND RV_POLottoCostruzioneRighe.TipoStringaLotto=" & TipoStringaLotto
sSQL = sSQL & " ORDER BY Posizione"


fnGetCodiceLotto = ""
StringaElaborata = ""
Codice = ""
PosizioneCodice = 0
SenzaIndentificativo = False
Set rs = Cn.OpenResultset(sSQL)




If Link_LottoArticolo > 0 Then
    For I = Len(CStr(Link_LottoArticolo)) To 7
        Codice = Codice & "0"
    Next
    Codice = Codice & Link_LottoArticolo
    
    
End If
            

If rs.EOF Then
    StringaElaborata = ""
    
Else
    
    PosizioneCodice = fnNotNullN(rs!PosConferimento)
    SenzaIndentificativo = fnNotNullN(rs!SenzaCodiceRiferimento_Conf)
    
    fnGetCodiceLotto = StringaLotto
    
    While Not rs.EOF
        Select Case fnNotNullN(rs!IDRV_POLottoComp)
            Case 1 'Codice socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.CDSocio.Code), 0, fnNotNullN(rs!SXDX))
            Case 2 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.CDSocio.Description), 1, fnNotNullN(rs!SXDX))
            Case 3 'Ragione sociale
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtNomeSocio.Text), 1, fnNotNullN(rs!SXDX))
            Case 4 'Giorno conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("d", Me.txtDataDocumento.Text)), 0, fnNotNullN(rs!SXDX))
            Case 5 'Mese del conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("m", Me.txtDataDocumento.Text)), 0, fnNotNullN(rs!SXDX))
            Case 6 'Anno del conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("yyyy", Me.txtDataDocumento.Text)), 0, fnNotNullN(rs!SXDX))
            Case 7 'Giorno lavorazione
                
            Case 8 'Mese lavorazione
                
            Case 9 'Anno lavorazione
                
            Case 10 'calibro
                
            Case 11 'Tipo lavorazione
            
            Case 12 'Tipo categoria
            
            Case 13 'Carattere speciale "_"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("_"), 1, fnNotNullN(rs!SXDX))
            Case 14 'Carattere speciale "-"
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr("-"), 1, fnNotNullN(rs!SXDX))
            Case 15 'Stringa personalizzata
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(fnNotNull(rs!Testo)), 1, fnNotNullN(rs!SXDX))
            Case 16 'Codice imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), Me.CDCodiceImballo.Code, 1, fnNotNullN(rs!SXDX))
            Case 17 'Descrizione imballo
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), Me.CDCodiceImballo.Description, 1, fnNotNullN(rs!SXDX))
            Case 18 'Codice pedana
                
            Case 19
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), Me.CDArticolo.Code, 1, fnNotNullN(rs!SXDX))
            Case 20
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), Me.CDArticolo.Description, 1, fnNotNullN(rs!SXDX))
            Case 22
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("ww", Me.txtDataDocumento.Text)), 0, fnNotNullN(rs!SXDX))
            Case 23
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", Me.txtDataDocumento.Text)), 0, fnNotNullN(rs!SXDX))
            Case 24
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtLottoDiConferimento.Text), 1, fnNotNullN(rs!SXDX))
            Case 25
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("y", Me.txtDataDocumento.Text)), 0, fnNotNullN(rs!SXDX))
            Case 26
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtNumeroDocumento.Value), 0, fnNotNullN(rs!SXDX))
            Case 27 'Codice certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_SOCIO(Me.CDSocio.KeyFieldID, "CodiceCertificazione")), 1, fnNotNullN(rs!SXDX))
            Case 28 'Descrizione certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_SOCIO(Me.CDSocio.KeyFieldID, "DescrizioneCertificazione")), 1, fnNotNullN(rs!SXDX))
            Case 29 'Protocollo certificazione del socio
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_SOCIO(Me.CDSocio.KeyFieldID, "ProtocolloCertificazione")), 1, fnNotNullN(rs!SXDX))
            Case 30 'Codice certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value, "CodiceCertificazione")), 1, fnNotNullN(rs!SXDX))
            Case 31 'Descrizione certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value, "DescrizioneCertificazione")), 1, fnNotNullN(rs!SXDX))
            Case 32 'Protocollo certificazione del lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value, "ProtocolloCertificazione")), 1, fnNotNullN(rs!SXDX))
            Case 33 'Giorno della settimana
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(DatePart("w", Me.txtDataDocumento.Text) - 1), 0, fnNotNullN(rs!SXDX))
            Case 34 'Codice utente
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(Me.txtCodiceUtente.Text), 1, fnNotNullN(rs!SXDX))
            Case 35 'Descrizione lotto di campagna
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_DESCRIZIONE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value)), 1, fnNotNullN(rs!SXDX))
            Case 36 'Codice certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_FAMIGLIA_PRODOTTO(Me.CDSocio.KeyFieldID, "CodiceCertificazione", Me.CDArticolo.KeyFieldID)), 1, fnNotNullN(rs!SXDX))
            Case 37 'Descrizione certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_FAMIGLIA_PRODOTTO(Me.CDSocio.KeyFieldID, "DescrizioneCertificazione", Me.CDArticolo.KeyFieldID)), 1, fnNotNullN(rs!SXDX))
            Case 38 'Protocollo certificazione della famiglia del prodotto venduto
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CERTIFICAZIONE_FAMIGLIA_PRODOTTO(Me.CDSocio.KeyFieldID, "ProtocolloCertificazione", Me.CDArticolo.KeyFieldID)), 1, fnNotNullN(rs!SXDX))
            Case 39 'Note aggiuntive della riga conferimento
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), Me.txtNoteAgg.Text, 1, fnNotNullN(rs!SXDX))
            Case 43 'Varietà dell'articolo conferito
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_VARIETA_ART(Me.CDArticolo.KeyFieldID, "Varieta")), 1, fnNotNullN(rs!SXDX))
            Case 44 'Codice della varietà dell'articolo conferito
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_VARIETA_ART(Me.CDArticolo.KeyFieldID, "CodiceImEx")), 1, fnNotNullN(rs!SXDX))
            Case 47 'Famiglia dell'articolo conferito
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_FAMIGLIA_ART(Me.CDArticolo.KeyFieldID, "FamigliaProdotti")), 1, fnNotNullN(rs!SXDX))
            Case 48 'Codice della famiglia dell'articolo conferito
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_FAMIGLIA_ART(Me.CDArticolo.KeyFieldID, "CodiceImEx")), 1, fnNotNullN(rs!SXDX))
            Case 51 'Varietà del lotto di produzione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value, "Varieta")), 1, fnNotNullN(rs!SXDX))
            Case 52 'Codice varietà del lotto di produzione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value, "CodiceImExVarieta")), 1, fnNotNullN(rs!SXDX))
            Case 53 'Famiglia del lotto di produzione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value, "FamigliaProdotti")), 1, fnNotNullN(rs!SXDX))
            Case 54 'Codice famiglia del lotto di produzione
                StringaElaborata = StringaElaborata & GET_STRINGALOTTO(fnNotNullN(rs!IDRV_POLottoComp), fnNotNullN(rs!Lunghezza), CStr(GET_CODICE_LOTTO_CAMPAGNA(Me.txtIDLottoCampagna.Value, "CodiceImEx")), 1, fnNotNullN(rs!SXDX))
        
        End Select
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing
    
    If StringaElaborata = "" Then
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Mid(Codice, 1, Len(Codice) - 1)
            Else
                fnGetCodiceLotto = Mid(Codice, 2, Len(Codice))
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    Else
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Codice & StringaElaborata
            Else
                fnGetCodiceLotto = StringaElaborata & Codice
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    End If
Exit Function
ERR_fnGetCodiceLotto:
    If StringaElaborata = "" Then
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Mid(Codice, 1, Len(Codice) - 1)
            Else
                fnGetCodiceLotto = Mid(Codice, 2, Len(Codice))
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    Else
        If SenzaIndentificativo = False Then
            If PosizioneCodice = 1 Then
                fnGetCodiceLotto = Codice & StringaElaborata
            Else
                fnGetCodiceLotto = StringaElaborata & Codice
            End If
        Else
            fnGetCodiceLotto = StringaElaborata
        End If
    End If
    
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
Private Function fnGetLottoArticolo() As Long

    Dim sSQL As String
    Dim IDLotto As Long
    IDLotto = fnGetNewKey("LottoArticolo", "IDLottoArticolo")
    Link_LottoArticolo = IDLotto
    sSQL = "INSERT INTO LottoArticolo (IDLottoArticolo, IDArticolo, Codice, LottoArticolo, DataScadenza) "
    sSQL = sSQL & "VALUES ("
    sSQL = sSQL & IDLotto & ", "
    sSQL = sSQL & m_DocumentsLink("IDArticolo").Value & ", "
    sSQL = sSQL & fnNormString(fnGetCodiceLotto(1, 1, m_DocumentsLink("CodiceLotto").Value, Link_LottoArticolo)) & ", "
    sSQL = sSQL & fnNormString(fnGetCodiceLotto(1, 2, m_DocumentsLink("DescrizioneLotto").Value, Link_LottoArticolo)) & ", "
    sSQL = sSQL & fnNormDate(DateAdd("m", 1, Date)) & ")"
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


Public Sub Movimenti(IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double)

     
End Sub
Private Function GeneraMovimentoDiCarico(IDRigaConferimento As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, _
PrezzoUnitario As Double, PrezzoImponibile As Double, Colli As Double, PesoLordo As Double, PesoNetto As Double, Tara As Double, Pezzi As Double) As Boolean

Cn.BeginTrans

Mov.DataMovimento = Me.txtDataDocumento.Text
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = Link_Esercizio
Mov.IDTipoOggetto = m_DocType.ID
Mov.IDOggetto = m_Document(m_Document.PrimaryKey).Value
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(IDDocumento, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Link_Magazzino_Carico
Mov.IDMagazzinoUscita = Link_Magazzino_Carico
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUMDiamante
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", PrezzoImponibile
Mov.Field "PrezzoUnitario", PrezzoUnitario
Mov.Field "DataDocumento", Me.txtDataDocumento.Text
Mov.Field "Oggetto", Mid(m_App.FunctionName & " del " & Me.txtDataDocumento & " Numero " & Me.txtNumeroDocumento, 1, 100)
Mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDRigaConferimento
Mov.Field "RV_POTipoRiga", 1
Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
Mov.Field "RV_POIDAssegnazioneMerce", 0
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", Me.CDSocio.KeyFieldID
Mov.Field "RV_PODataConferimento", Me.txtDataDocumento.Text
Mov.Field "RV_PONumeroConferimento", Me.txtNumeroDocumento.Value
Mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
Mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POQuantitaMovimentata", Qta_UM
Mov.Field "RV_PONumeroColli", Colli
Mov.Field "RV_POPesoLordo", PesoLordo
Mov.Field "RV_POPesoNetto", PesoNetto
Mov.Field "RV_POTara", Tara
Mov.Field "RV_POQuantitaPezzi", Pezzi
Mov.Field "RV_POIDAnagraficaFatturazione", Me.CDSocioFatt.KeyFieldID

Mov.Field "TipoRiga", trcNessuno

GeneraMovimentoDiCarico = Mov.Insert

If GeneraMovimentoDiCarico = False Then
    Cn.RollbackTrans
Else
    Cn.CommitTrans
End If



End Function
Private Function GeneraMovimentoDiScarico(IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double) As Boolean
Dim Mov As DmtMovim.cMovimentazione

Cn.BeginTrans

Mov.DataMovimento = Me.txtDataDocumento.Text
Mov.FattoreDiConversione = Null
Mov.IDEsercizio = Link_Esercizio
Mov.IDOggetto = Link_Oggetto
Mov.IDTipoOggetto = fnGetTipoOggetto
Mov.IDFunzione = Link_CausaleScarico
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoUscita = Link_Magazzino_Scarico
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUMDiamante
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", 0
Mov.Field "DataDocumento", Date
Mov.Field "Oggetto", Mid("Conferimento del " & Me.txtDataDocumento & " Numero " & Me.txtNumeroDocumento, 1, 100)
Mov.Field "IDTipoMovimento", 3
Mov.Field "TipoRiga", trcNessuno


GeneraMovimentoDiScarico = Mov.Insert

If GeneraMovimentoDiScarico = False Then
    Cn.RollbackTrans
Else
    Cn.CommitTrans
End If


End Function
Public Function GeneraMovimentoCaricoImballo(IDRigaConferimento As Long, CodiceLottoEntrata As String, CodiceLottoCampagna As String, IDArticolo As Long, IDUMDiamante As Long, Articolo As String, Qta_UM As Double, IDLottoImballo As Long, PrezzoUnitario As Double, PrezzoImponibile As Double) As Boolean
Cn.BeginTrans

Mov.DataMovimento = Me.txtDataDocumento.Text
Mov.FattoreDiConversione = Null

Mov.GestioneMatricole = False
Mov.IDEsercizio = Link_Esercizio
Mov.IDTipoOggetto = m_DocType.ID
Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(IDDocumento, 1)
Mov.IDUtente = TheApp.IDUser
Mov.IDMagazzinoEntrata = Link_Magazzino_Carico
Mov.IDMagazzinoUscita = Link_Magazzino_Carico
Mov.Cessione = 0
Mov.Field "IDAzienda", TheApp.IDFirm
If Me.chkPreConferimento.Value = vbUnchecked Then
    Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
Else
    Mov.Field "IDAnagrafica", IDSOCIO_PRE_CONF
End If
Mov.Field "IDTipoAnagrafica", 3
Mov.Field "IDArticolo", IDArticolo
Mov.Field "IDUnitaDiMisura", IDUMDiamante
Mov.Field "IDcambio", Null
Mov.Field "DescrizioneArticolo", Articolo
Mov.Field "QuantitaTotale", Qta_UM
Mov.Field "Importo", PrezzoImponibile
Mov.Field "PrezzoUnitario", PrezzoUnitario
Mov.Field "DataDocumento", Me.txtDataDocumento.Text
Mov.Field "Oggetto", Mid(TheApp.FunctionName & " del " & Me.txtDataDocumento & " Numero " & Me.txtNumeroDocumento, 1, 100)
Mov.Field "IDTipoMovimento", 1

'DATI DI CONFERIMENTO
Mov.Field "IDValoriOggettoDettaglio", IDRigaConferimento
Mov.Field "RV_POTipoRiga", 2
Mov.Field "RV_POIDCaricoMerceRighe", IDRigaConferimento
Mov.Field "RV_POIDAssegnazioneMerce", 0
Mov.Field "RV_POIDProcessoIVGamma", 0
Mov.Field "RV_POIDAnagraficaSocio", Me.CDSocio.KeyFieldID
Mov.Field "RV_PODataConferimento", Me.txtDataDocumento.Text
Mov.Field "RV_PONumeroConferimento", Me.txtNumeroDocumento.Value
Mov.Field "RV_POCodiceLotto", CodiceLottoEntrata
Mov.Field "RV_POCodiceLottoCampagna", CodiceLottoCampagna
Mov.Field "RV_POQuantitaLiquidazione", 0
Mov.Field "RV_POImportoInclusoImballo", 0
Mov.Field "RV_POImportoLiquidazione", 0
Mov.Field "RV_POIDAnagraficaFatturazione", Me.CDSocioFatt.KeyFieldID

Mov.Field "RV_POIDLottoImballo", IDLottoImballo
Mov.Field "LottoImballo", GET_DESCRIZIONE_LOTTO_IMBALLO(IDLottoImballo)



Mov.Field "TipoRiga", trcNessuno

GeneraMovimentoCaricoImballo = Mov.Insert

If GeneraMovimentoCaricoImballo = False Then
    Cn.RollbackTrans
Else
    Cn.CommitTrans
End If


End Function
Private Function fnGetIDMovimentoMagazzino() As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT IDMovimento FROM Movimento ORDER BY IDMovimento DESC"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        fnGetIDMovimentoMagazzino = rs!IDMovimento
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function
Public Function fnGetMagazzinoCarico() As Long
   Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POSchemaCoop.IDFiliale, RV_POProcessiDocumentoCoop.IDRV_POProcessiDocumentoCoop, "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop, RV_POProcessiDocumentoCoop.IDDocumentoCoop,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoProcessoCoop, RV_POProcessiDocumentoCoop.IDMagazzino,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoMagazzino , RV_POSchemaCoop.IDUtente "
    sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop ON "
    sSQL = sSQL & "RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop "
    sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale = " & m_App.Branch & ") AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDDocumentoCoop =" & IDDocumento & ") AND (RV_POSchemaCoop.IDUtente = 0) AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDTipoProcessoCoop = 1)"
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Select Case rs!IDTipoMagazzino
            Case 1
                Link_Magazzino_Carico = Me.cboMagazzinoConf.CurrentID
                Link_CausaleCarico = Link_Causale_MagCar_Conf
            Case 2
                Link_Magazzino_Carico = Me.CboMagazzinoVend.CurrentID
                Link_CausaleCarico = Link_Causale_MagCar_Vend
                
        End Select
    Else
        Link_Magazzino_Carico = Me.cboMagazzinoConf.CurrentID
        Link_CausaleCarico = Link_Causale_MagCar_Conf
       
    End If
        
    
End Function
Public Function fnGetMagazzinoScarico() As Long
   Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT RV_POSchemaCoop.IDFiliale, RV_POProcessiDocumentoCoop.IDRV_POProcessiDocumentoCoop, "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop, RV_POProcessiDocumentoCoop.IDDocumentoCoop,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoProcessoCoop, RV_POProcessiDocumentoCoop.IDMagazzino,"
    sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDTipoMagazzino , RV_POSchemaCoop.IDUtente "
    sSQL = sSQL & "FROM RV_POSchemaCoop INNER JOIN "
    sSQL = sSQL & "RV_POProcessiDocumentoCoop ON "
    sSQL = sSQL & "RV_POSchemaCoop.IDRV_POSchemaCoop = RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop "
    sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale = " & m_App.Branch & ") AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDDocumentoCoop =" & IDDocumento & ") AND (RV_POSchemaCoop.IDUtente = 0) AND "
    sSQL = sSQL & "(RV_POProcessiDocumentoCoop.IDTipoProcessoCoop = 2)"
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        Select Case rs!IDTipoMagazzino
            Case 1
                Link_Magazzino_Scarico = Me.cboMagazzinoConf.CurrentID
                Link_CausaleScarico = Link_Causale_MagScar_Conf
            Case 2
                Link_Magazzino_Scarico = Me.CboMagazzinoVend.CurrentID
                Link_CausaleScarico = Link_Causale_MagScar_Vend
                
        End Select
    Else
        Link_Magazzino_Carico = Me.cboMagazzinoConf.CurrentID
        Link_CausaleScarico = Link_Causale_MagCar_Conf
       
    End If
        
    
End Function
Private Sub AggiornamentoMovimenti()
Dim Messagggio As String

Dim Mov As DmtMovim.cMovimentazione
Set Mov = New DmtMovim.cMovimentazione
Dim prg As dmtProg.cRicostruzione
If m_DocumentsLink("IDMovimento_Carico").Value > 0 Then
    Set Mov.Connection = TheApp.Database.Connection
    Mov.Read m_DocumentsLink("IDMovimento_Carico").Value
    
    
    Mov.DataMovimento = Me.txtDataDocumento.Text
    Mov.IDEsercizio = Link_Esercizio
    Mov.IDTipoOggetto = fnGetTipoOggetto
    
    Mov.Field "QuantitaTotale", m_DocumentsLink("Qta_UM").Value
    Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
    Mov.Field "DataDocumento", Me.txtDataDocumento.Text
    Mov.Insert
   
End If

    

If m_DocumentsLink("IDMovimento_vendita").Value > 0 Then
    Set Mov.Connection = TheApp.Database.Connection
    Mov.Read m_DocumentsLink("IDMovimento_Carico").Value

    Mov.IDEsercizio = Link_Esercizio
    Mov.IDOggetto = Link_Oggetto
    Mov.IDTipoOggetto = fnGetTipoOggetto
    
    Mov.Field "QuantitaTotale", m_DocumentsLink("Qta_UM").Value
    
    Mov.Insert
    
End If

Set prg = New dmtProg.cRicostruzione

Set prg.Connessione = TheApp.Database.Connection
    prg.Filtri.Add m_App.IDFirm, "IDAzienda"

    prg.Filtri.Add Link_Esercizio, "IDEsercizio"

    prg.Filtri.Add Link_Magazzino_Carico, "IDMagazzino"

    prg.Where = "Articolo.CodiceArticolo = " & fnNormString(m_DocumentsLink("CodiceArticolo").Value)
    'Vengono ricostruiti in questo caso le giacenze dei lotti degli articoli che hanno codice compreso tra 'A000' e 'A999', di esercizio, azienda e magazzino specificati

    prg.RicostruzioneLotti
    prg.RicostruzioneProgressivi
    
    


End Sub
Private Function ControlloAltraMovimentazioneRiga() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

    If Link_LottoArticolo > 0 Then
        sSQL = "SELECT * FROM Movimento WHERE ("
        sSQL = sSQL & "(IDTipoOggetto<>" & fnGetTipoOggetto & ") "
        sSQL = sSQL & "AND (IDLottoArticolo=" & Link_LottoArticolo & "))"
    
        Set rs = Cn.OpenResultset(sSQL)

        If rs.EOF = False Then
            ControlloAltraMovimentazioneRiga = True
        Else
            ControlloAltraMovimentazioneRiga = False
        End If

        rs.CloseResultset
        Set rs = Nothing
    End If
End Function
Private Function fnGetTaraImballo(IDArticolo As Long) As Double
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT Tara FROM Articolo WHERE "
    sSQL = sSQL & "IDArticolo = " & IDArticolo
    
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
Private Sub CalcoloPesoNetto(TipoCalcolo As Long)
On Error Resume Next
Dim ArrayPesoNetto() As String
Dim PesoNetto As Double
Dim Decimal_PesoNetto As Double
Dim Tara As Double

Me.txtTara.Value = ((Me.txtColli.Value * Me.txtTaraUnitaria.Value) + (Me.txtTaraPedana.Value * Me.txtQuantitaPedana.Value)) + Me.txtTaraAutomezzo.Value

If ATTIVA_CALCOLO_PESO_LORDO = 0 Then
    Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
Else
    If PESO_LORDO_ARTICOLO = 0 Then
    
        Select Case TipoCalcolo
            Case 1
                Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
            Case 2
                Me.txtPesoLordo.Value = Me.txtPesoNetto.Value + Me.txtTara.Value
            Case Else
                Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
        End Select
    Else
        If TIPO_PESO_ARTICOLO <= 1 Then
            Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
        Else
            If TipoCalcolo = 0 Then
                Me.txtPesoLordo.Value = Me.txtTara.Value + Me.txtPesoNetto.Value
            Else
                Me.txtPesoNetto.Value = Me.txtColli.Value * PESO_LORDO_ARTICOLO
                Me.txtPesoLordo.Value = Me.txtTara.Value + Me.txtPesoNetto.Value
            End If
        End If
    End If
End If

Select Case LINK_TIPO_ARROTONDAMENTO
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
    'Me.txtPesoNetto.Value = Me.txtPesoLordo.Value - Me.txtTara.Value
    
GET_RIEPILOGO_PESI

End Sub

Private Sub txtTara_LostFocus()
    CalcoloPesoNetto 1
End Sub
Private Function ControlloRigheMovimentate() As Boolean
'RESTITUISCE TRUE SE NON SONO STATE MOVIMENTATE LE RIGHE DA ALTRI GESTORI
'ALTRIMENTI RESTITUISCE FALSE

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & m_Document("IDRV_POCaricoMerceTesta").Value

Set rs = Cn.OpenResultset(sSQL, adOpenKeyset)

If rs.EOF = True Then
    ControlloRigheMovimentate = True

Else
    While Not rs.EOF
        If IsNull(rs!IDRV_POCaricoMerceRighe) Then
            ControlloRigheMovimentate = True
        Else
            If ControlloAltraMovimentazioneRigaPerEliminazione(rs!IDRV_POCaricoMerceRighe) = True Then
                ControlloRigheMovimentate = True
            Else
                ControlloRigheMovimentate = False
                rs.MoveLast
            End If
        End If
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing

ControlloRigheMovimentate = ControlloRigheMovimentate


End Function
Private Function ControlloAltraMovimentazioneRigaPerEliminazione(IDRigaConferimento) As Boolean
'RESTITUSCE TRUE SE NON E STATA SE IL RECORDSET E' VUOTO

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT * FROM Movimento WHERE ("
    sSQL = sSQL & "(IDTipoOggetto<>" & fnGetTipoOggetto & ") "
    sSQL = sSQL & "AND (RV_POIDCaricoMerceRighe=" & IDRigaConferimento & "))"

    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = True Then
        ControlloAltraMovimentazioneRigaPerEliminazione = True
    Else
        ControlloAltraMovimentazioneRigaPerEliminazione = False
    End If
    rs.CloseResultset
    Set rs = Nothing

End Function

Private Sub EliminaRigheDaDocumento()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection


If ATTIVAZIONE_NUOVO_CALCOLO = False Then
    sSQL = "SELECT * FROM RV_POCaricoMerceRighe "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & m_Document("IDRV_POCaricoMerceTesta").Value

    
    Set rs = Cn.OpenResultset(sSQL)
    While Not rs.EOF
    
        If rs!IDMovimento_Carico > 0 Then
            EliminaMovimento 0, fnNotNullN(rs!IDMovimento_Carico)
        End If
    
        If fnNotNullN(rs!IDMovimentoImballo) > 0 Then
            EliminaMovimento 0, fnNotNullN(rs!IDMovimentoImballo)
        End If
    

    rs.MoveNext
    Wend

    rs.CloseResultset
    Set rs = Nothing
Else


    Mov.IDTipoOggetto = fnGetTipoOggetto
    Mov.IDOggetto = m_Document("IDRV_POCaricoMerceTesta").Value
    Mov.Delete
End If
Set Mov = Nothing
End Sub
Private Sub EliminaMovimento(IDRigaConferimento As Long, IDMovimento As Long)
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim IDMovimento_Local As Long

If ATTIVAZIONE_NUOVO_CALCOLO = True Then
    sSQL = "SELECT IDMovimento FROM Movimento "
    sSQL = sSQL & "WHERE IDTipoOggetto=" & m_DocType.ID
    sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDRigaConferimento
    
    Set rs = Cn.OpenResultset(sSQL)

    While Not rs.EOF
        If Mov.Delete(fnNotNull(rs!IDMovimento)) = False Then
            MsgBox "Impossibile eliminare il movimento (IDMovimento: " & IDMovimento & ")", vbCritical, "Eliminazione movimento di carico"
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    
Else

    If Mov.Delete(IDMovimento) = False Then
        MsgBox "Impossibile eliminare il movimento (IDMovimento: " & IDMovimento & ")", vbCritical, "Eliminazione movimento di carico"
    End If

End If
End Sub
Private Sub EliminaLotto(IDLottoArticolo As Long)
'On Error GoTo ERR_EliminaLotto
'Dim Errore As String
'Dim sSQL As String'

'Errore = "Eliminazione lotto articolo"
'sSQL = "DELETE FROM LottoArticolo WHERE IDLottoArticolo=" & IDLottoArticolo
'Cn.Execute sSQL

'Errore = "Eliminazione lotto articolo per magazzino"
'sSQL = "DELETE FROM LottoArticoloPerMagazzino WHERE IDLottoArticolo=" & IDLottoArticolo
'Cn.Execute sSQL
'Exit Sub
'ERR_EliminaLotto:
'    MsgBox Err.Description, vbCritical, Errore
End Sub
Private Function ControlloQuantitaDiCarico()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT * FROM Movimento WHERE IDMovimento=" & m_DocumentsLink("IDMovimento_Carico").Value
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    If rs!IDTipoOggetto <> fnGetTipoOggetto Then
        If rs!QuantitaTotale <> m_DocumentsLink("Qta_UM").Value Then
        'NumeroRiga = m_DocumentsLink("IDRV_POCaricoMerceRighe").Value
            Me.txtQta_UM.Value = rs!QuantitaTotale
            m_DocumentsLink("Qta_UM").Value = rs!QuantitaTotale
        
            Select Case Link_UnitaDiMisura_Coop
            Case 1
                Me.txtColli.Value = Me.txtQta_UM.Value
            Case 2
                Me.txtPesoLordo.Value = Me.txtQta_UM.Value
            Case 3
                Me.txtPesoNetto.Value = Me.txtQta_UM.Value
            Case 4
                Me.txtTara.Value = Me.txtQta_UM.Value
            Case 5
                Me.txtPezzi.Value = Me.txtQta_UM.Value
            End Select
            cmdSalva_Click
        
        End If
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function

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
    
    'Connessione di tipo dmtDataLayer
        'ConnessioneDiamanteOLD
    'Connessione di tipo DMTADODBLib
        ConnessioneDiamanteADO
        
        GET_PARAMATRI_FILIALE
        
'        ParametroSocio
'        ParametroImballo
'        ParametroGrezzo
'        ParametroLavorato
        RecuperaOperazioniDocumento
        fnRecuperaAnnotazioniPerDoc
'        ParametroTipoScarto
'        ParametroTipoCaloPeso
'        ParametroTipoAumentoPeso
'        ParametroNuovoCalcolo
'        ParametroTipoSceltaArticoloConferito
'        ParametroLottoObbligatorio
'        ParametroTipoArrotondamento
'        ParametroPrezzoMedioAutomatico
'        ParametroAggiornaPrezzoMedioDaConf
'        ParametroAggiornaTipoLavorazioneDaConf
        GET_TIPO_LIQUIDAZIONE_PER_CONFERIMENTO TheApp.Branch
'        ParametroStampaSituazioneImballi
'        ParametroAttivaCalcoloPesoLordoConf
'        ParametroAttivaImportoRiepConf
        'GET_CONTROLLO_LICENZA
        GET_MODULO_ATTIVATO MODULO_CODICE, 80
        LINK_LOTTO_IMBALLO_PRED = GET_LINK_LOTTO_IMBALLO_PRED(TheApp.IDFirm)
    'Inizializzazioni da fare prima dell'apertura del documento
        OnBeforeOpenDoc
        
        
    'rif12
    'Altre inizializzazioni
        OnStart
    

    
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
        m_DocType.Fields("ID" & m_App.TableName).Value = m_App.CallerFieldValue '872
        'MsgBox m_App.CallerFieldValue
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"
        
        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")
        
        'Inidica, nel caso di esegui gestione, se riportare il valore corrente al chiamante
        bNotReturnValue = CBool(Val(GetSetting(REGISTRY_KEY, App.EXEName, "NoReturnValue", "0")))
        
        LINK_RIGA_CONFERIMENTO_DA_MOV = Val(GetSetting(REGISTRY_KEY, App.EXEName, "IDRigaConferimento"))  '1880
        'MsgBox LINK_RIGA_CONFERIMENTO_DA_MOV
        Set m_Document.ActiveFilter = m_ActiveFilter

    Else
        '---------------------------------------------------
        '     Il programma non è stato chiamato da un link.
        '---------------------------------------------------
   '
        'Il filtro attivo alla partenza è quello predefinito
        For Each oFilter In m_DocType.Filters
            If oFilter.ID = oFiltersActivity.DefaultFilterID Then
                Set m_ActiveFilter = m_DocType.Filters(oFilter.Name)
                Exit For
            End If
        Next '

        'Si comunica al documento quale filtro eseguire all'avvio.
        Set m_Document.ActiveFilter = m_ActiveFilter
        Set BrwMain.Recordset = m_Document.Data
        m_Document.Dataset.Recordset.Sort = "DataDocumento DESC, NumeroDocumento DESC"
        Set Me.BrwMain.Recordset = m_Document.Dataset.Recordset
    End If
    
        
    'Si comunica al documento quale filtro eseguire all'avvio.
    'Set m_Document.ActiveFilter = m_ActiveFilter
    '    m_Document.Dataset.Recordset.Sort = "DataDocumento DESC, NumeroDocumento DESC"
    'Set Me.BrwMain.Recordset = m_Document.Dataset.Recordset
    'Prima di aprire il documento occorre comunicargli qual'è il campo chiave primaria.
    m_Document.PrimaryKey = "ID" & m_Document.TableName
    'Apertura del documento.
    AggiornamentoDocumento = 1
    m_Document.OpenDoc
    AggiornamentoDocumento = 0
    
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
        BrwMain.LoadUserSettings
        SetVisibilityIDFields
    End If
    
    
    'Crea i campi per la ricerca.
    CreateBrowserConditions
    'Assegnazione del riferimento alla fonte dati (binding sul recordset del documento)
    
    'rif14

    
    'Set BrwMain.Recordset = m_Document.Dataset.Recordset
    Set BrwMain.Recordset = m_Document.Data
    
    
            
     'Viene inizializzato il dialogo di stampa
  '  With DmtPrnDlg
  '      Set .Application = m_App
  '      Set .DocType = m_DocType
  '  End With
    
    
    
    'Ripulisco la tabella semaforo.
    'Se era avvenuto un crash di sistema questo garantisce il ripristino della situazione.
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    
    'Evita il blocco della toolbar
    'BarMenu.ResetHooks
    
    Screen.MousePointer = OLDCursor


    If LINK_RIGA_CONFERIMENTO_DA_MOV > 0 Then
        m_DocumentsLink_OnReposition
        m_Changed = False
    End If
End Sub
Private Function SalvataggioRighe(IDConferimentoTesta As Long) As Boolean
On Error GoTo ERR_SalvataggioRighe
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsCount As DmtOleDbLib.adoResultset
Dim Unita_progresso As Double

Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection

Screen.MousePointer = 11
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 100


    Mov.IDTipoOggetto = fnGetTipoOggetto
    Mov.IDOggetto = IDConferimentoTesta
    Mov.Delete

    
    sSQL = "SELECT * FROM RV_POCaricoMerceRighe "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDConferimentoTesta
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection
    
    While Not rs.EOF
        If fnNotNullN(rs!IDArticolo) > 0 Then
                 
'            If fnNotNullN(rs!IDMovimento_Carico) > 0 Then
'                EliminaMovimento fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDMovimento_Carico)
'            End If
'
'
'            If fnNotNullN(rs!IDMovimentoImballo) > 0 Then
'                EliminaMovimento fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDMovimentoImballo)
'            End If
            
            If GeneraMovimentoDiCarico(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNull(rs!CodiceLotto), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDArticolo), fnNotNullN(rs!IDUnitaDiMisuraDiamante), fnNotNull(rs!Articolo), fnNotNullN(rs!Qta_UM), _
                fnNotNullN(rs!ImportoUnitario), fnNotNullN(rs!ImportoUnitario) * fnNotNullN(rs!Qta_UM), _
                fnNotNullN(rs!Colli), fnNotNullN(rs!PesoLordo), fnNotNullN(rs!PesoNetto), _
                fnNotNullN(rs!Tara), fnNotNullN(rs!Pezzi)) = False Then
                
                SalvataggioRighe = False
                rs.Close
                Set rs = Nothing
                Set Mov = Nothing
                Screen.MousePointer = 0
                Exit Function
            End If

            If fnNotNullN(rs!IDImballo > 0) Then
                If GeneraMovimentoCaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNull(rs!CodiceLotto), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDImballo), fnNotNullN(GET_UM_ARTICOLO(rs!IDImballo)), fnNotNull(rs!DescrizioneImballo), fnNotNullN(rs!Colli), fnNotNullN(rs!IDRV_POLottoImballo), fnNotNullN(rs!ImportoUnitarioImballo), fnNotNullN(rs!ImportoUnitarioImballo) * fnNotNullN(rs!Colli)) = False Then
                    SalvataggioRighe = False
                    rs.Close
                    Set rs = Nothing
                    Set Mov = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
                
            End If
            If ((fnNotNullN(rs!QuantitaPedana) > 0) And (fnNotNullN(rs!IDArticoloPedana) > 0)) Then
                If GeneraMovimentoCaricoImballo(fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNull(rs!CodiceLotto), fnNotNull(rs!LottoDiConferimento), fnNotNullN(rs!IDArticoloPedana), fnNotNullN(GET_UM_ARTICOLO(rs!IDArticoloPedana)), fnNotNull(rs!ArticoloPedana), fnNotNullN(rs!QuantitaPedana), 0, 0, 0) = False Then
                    SalvataggioRighe = False
                    rs.Close
                    Set rs = Nothing
                    Set Mov = Nothing
                    Screen.MousePointer = 0
                    Exit Function
                End If
            End If
            
'            If (fnNotNullN(rs!IDMovimento_Carico) > 0) Or (fnNotNullN(rs!IDMovimentoImballo) > 0) Then
'                sSQL = "UPDATE RV_POCaricoMerceRighe SET "
'                sSQL = sSQL & "IDMovimento_Carico=0 , "
'                sSQL = sSQL & "IDMovimentoImballo=0 "
'                sSQL = "WHERE IDRV_POCaricoMerceRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
'                Cn.Execute sSQL
'            End If
        End If
       
        RICALCOLA_CAMPIONATURA fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!Qta_UM)
        
        AGGIORNA_PESATURA fnNotNullN(rs!IDRV_POCaricoMerceRighe), fnNotNullN(rs!IDRV_POcaricoMercetesta), fnNotNullN(rs!LINK_ORDINAMENTO)
    
    DoEvents
    rs.MoveNext
    Wend

    rs.Close
    Set rs = Nothing
    
 
Set Mov = Nothing
Me.ProgressBar1.Value = 0
Screen.MousePointer = 0

SalvataggioRighe = True

ContDelete = 0
ContDeleteConferimento = 0

RefreshArray
Exit Function
ERR_SalvataggioRighe:
    MsgBox Err.Description, vbCritical, "Salvataggio Righe"
    SalvataggioRighe = False
End Function
Private Function fnEsistenza_Lavorazione(IDRiga As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POLavorazione WHERE IDRV_POCaricoMerceRighe=" & IDRiga

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = True Then
    fnEsistenza_Lavorazione = False
Else
    fnEsistenza_Lavorazione = True
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub EliminazioneRighe()
Dim I As Integer
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
If ArrayDelete(0) = 0 Then
    Exit Sub
End If

For I = 0 To 150
    If ArrayDelete(I) > 0 Then
        EliminaMovimento 0, ArrayDelete(I)
    End If
Next

End Sub
Private Sub fnEliminaLavorazione(IDConferimentoRiga As Long)
Dim sSQL As String

sSQL = "DELETE FROM RV_POLavorazione "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Cn.Execute sSQL
End Sub
Private Sub fnEliminaAssegnazione(IDConferimentoRiga As Long)
Dim sSQL As String

sSQL = "DELETE FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga

Cn.Execute sSQL
End Sub
Private Sub fnEliminaRiferimentiDDT(IDConferimentoRiga As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoDettaglio0004 SET "
sSQL = sSQL & "RV_POIDConferimentoRighe=0 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga

Cn.Execute sSQL
End Sub
Private Sub fnEliminaRiferimentiFA(IDConferimentoRiga As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoDettaglio0001 SET "
sSQL = sSQL & "RV_POIDConferimentoRighe=0 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga

Cn.Execute sSQL
End Sub
Private Sub fnEliminaRiferimentiSNF(IDConferimentoRiga As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoDettaglio0034 SET "
sSQL = sSQL & "RV_POIDConferimentoRighe=0 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga

Cn.Execute sSQL
End Sub

Private Sub fnEliminaRiferimentiNC(IDConferimentoRiga As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoDettaglio0016 SET "
sSQL = sSQL & "RV_POIDConferimentoRighe=0 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga

Cn.Execute sSQL
End Sub
Private Sub fnEliminaRiferimentiND(IDConferimentoRiga As Long)
Dim sSQL As String

sSQL = "UPDATE ValoriOggettoDettaglio0007 SET "
sSQL = sSQL & "RV_POIDConferimentoRighe=0 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga

Cn.Execute sSQL
End Sub
Public Function ControlloQuantita() As Boolean
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String

sSQL = "SELECT Qta_UM From RV_POCaricoMerceRighe WHERE IDRV_POCaricoMerceRighe=" & m_DocumentsLink("IDRV_POCaricoMerceRighe").Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = True Then
    ControlloQuantita = False
Else
    If m_DocumentsLink("Qta_UM").Value <> rs!Qta_UM Then
        ControlloQuantita = True
    Else
        ControlloQuantita = False
    End If
End If

rs.CloseResultset
Set rs = Nothing



End Function
Public Function PermessoSalvataggio() As Boolean
Dim Testo As String
Dim NumeroColliPes As Double
Dim TotalePesoLordoPesatura As Double
Dim TotalePezziPesatura As Double

PermessoSalvataggio = True

If Me.CDArticolo.KeyFieldID <= 0 Then
    MsgBox "Inserire l'articolo", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.CDArticolo.SetFocus
    Exit Function
End If
If Me.txtQta_UM.Value <= 0 Then
    MsgBox "La quantità deve essere maggiore di zero", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    'Me.txtQta_UM.SetFocus
    Exit Function
End If
If Me.txtColli.Value < 0 Then
    MsgBox "La quantità dei colli deve essere maggiore o uguale a zero", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtColli.SetFocus
    Exit Function
    
End If

If Me.txtPesoLordo.Value < 0 Then
    MsgBox "La quantità del peso lordo deve essere maggiore o uguale a zero", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtPesoLordo.SetFocus
    Exit Function
    
End If
If Me.txtPesoNetto.Value < 0 Then
    MsgBox "La quantità del peso netto deve essere maggiore o uguale a zero", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtPesoNetto.SetFocus
    Exit Function
    
End If
If Me.txtPezzi.Value < 0 Then
    MsgBox "La quantità dei pezzi deve essere maggiore o uguale a zero", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtPezzi.SetFocus
    Exit Function
    
End If
If Me.txtTara.Value < 0 Then
    MsgBox "La quantità della tara deve essere maggiore o uguale a zero", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtTara.SetFocus
    Exit Function
    
End If
If (Me.txtPesoLordo.Value < Me.txtPesoNetto.Value) Then
    MsgBox "La quantità del peso netto deve essere minore o uguale al valore del peso lordo", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.txtPesoLordo.SetFocus
    Exit Function

End If
If Me.cboUM.CurrentID = 0 Then
    MsgBox "Non è stata impostata l'unità di misura dell'articolo", vbInformation, "Errore salvataggio"
    PermessoSalvataggio = False
    Me.CDArticolo.SetFocus
    Exit Function
End If

If LOTTO_CAMPAGNA_OBBLIGATORIO = 1 Then
    If Me.txtIDLottoCampagna.Value = 0 Then
        MsgBox "Il lotto di campagna è obbligatorio", vbInformation, "Impossibile salvare"
        PermessoSalvataggio = False
        Me.txtLottoDiConferimento.SetFocus
        Exit Function
    End If
End If

If Me.txtIDLottoCampagna.Value > 0 Then
    If Me.txtDataSbloccoLotto.Value > 0 Then
        If DateDiff("d", Me.txtDataSbloccoLotto.Text, Date) < 0 Then
            MsgBox "Il lotto di campagna risulta bloccato", vbInformation, "Impossibile salvare"
            PermessoSalvataggio = False
            Me.txtLottoDiConferimento.SetFocus
            Exit Function
        End If
    End If
End If

NumeroColliPes = GET_SOMMA_COLLI_PES
TotalePesoLordoPesatura = GET_SOMMA_PESO_PES
TotalePezziPesatura = GET_SOMMA_PEZZI_PES

If NumeroColliPes > 0 Then
    If ((Me.txtColli.Value <> NumeroColliPes) Or (Me.txtPesoLordo.Value <> TotalePesoLordoPesatura) Or (Me.txtPezzi.Value <> TotalePezziPesatura)) Then
        Testo = "ATTENZIONE!!!" & vbCrLf
        Testo = Testo & "Il numero dei colli o il peso lordo o i pezzi del conferimento non sono uguali a quelli che risultano dalle pesature." & vbCrLf
        Testo = Testo & "Continuare a salvare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Validazione dati") = vbNo Then
            PermessoSalvataggio = False
            Exit Function
        End If
    End If
End If

End Function

Private Sub txtTaraAutomezzo_LostFocus()
    CalcoloPesoNetto 1
End Sub

Private Sub txtTaraPedana_LostFocus()
    CalcoloPesoNetto 1
End Sub

Private Sub RecuperaOperazioniDocumento()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POOperazionePerDoc.GestioneArticoli, RV_POOperazionePerDoc.CreazioneAutomaticaLottoVend "
sSQL = sSQL & "FROM RV_POOperazionePerDoc INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_POOperazionePerDoc.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE (RV_POSchemaCoop.IDFiliale =" & m_App.Branch & ") And (RV_POSchemaCoop.IDUtente = 0) And (RV_POOperazionePerDoc.IDRV_PODocumentoCoop = " & IDDocumento & ")"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Flag_GestioneArticoli = rs!GestioneArticoli
    Flag_AutomazioneLotti = rs!CreazioneAutomaticaLottoVend
Else
    Flag_GestioneArticoli = 0
    Flag_AutomazioneLotti = 0
End If

rs.CloseResultset
Set rs = Nothing

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
    S_Annotazioni1 = fnNotNull(rs!Annotazioni1)
    S_Annotazioni2 = fnNotNull(rs!Annotazioni2)
    S_Annotazioni3 = fnNotNull(rs!Annotazioni3)
Else
    S_Annotazioni1 = ""
    S_Annotazioni2 = ""
    S_Annotazioni3 = ""
End If

rs.CloseResultset
Set rs = Nothing

End Sub

Private Function fnGetCodiceSocio(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Codice FROM Fornitore WHERE "
sSQL = sSQL & "IDAzienda=" & TheApp.IDFirm & " AND "
sSQL = sSQL & "IDAnagrafica=" & IDAnagrafica

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    fnGetCodiceSocio = 0
Else
    fnGetCodiceSocio = fnNotNull(rs!Codice)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CambioMovimentazioneSocioPerConferimento()
Dim sSQL As String
Dim rsRighe As DmtOleDbLib.adoResultset
Dim rsLav As DmtOleDbLib.adoResultset

Dim Mov As DmtMovim.cMovimentazione
Set Mov = New DmtMovim.cMovimentazione
Set Mov.Connection = TheApp.Database.Connection


sSQL = "SELECT * FROM RV_POCaricoMerceRighe WHERE IDRV_POCaricoMerceTesta=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
Set rsRighe = Cn.OpenResultset(sSQL)

While Not rsRighe.EOF
    
    sSQL = "SELECT * FROM RV_POLavorazione WHERE IDRV_POCaricoMerceRighe=" & fnNotNullN(rsRighe!IDRV_POCaricoMerceRighe)
    
    Set rsLav = Cn.OpenResultset(sSQL)
    
    While Not rsLav.EOF
        If fnNotNullN(rsLav!IDMovimento_Carico) > 0 Then
            Mov.Read fnNotNullN(rsLav!IDMovimento_Carico)
            Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
            Mov.Insert
        End If
        If fnNotNullN(rsLav!IDMovimento_Vendita) > 0 Then
            Mov.Read fnNotNullN(rsLav!IDMovimento_Vendita)
            Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
            Mov.Insert
        
        End If
        
    rsLav.MoveNext
    Wend
    
    rsLav.CloseResultset
    Set rsLav = Nothing

If fnNotNullN(rsRighe!IDMovimento_Carico) > 0 Then
    Mov.Read fnNotNullN(rsRighe!IDMovimento_Carico)
    Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
    Mov.Insert

End If
    
rsRighe.MoveNext
Wend

rsRighe.CloseResultset
Set rsRighe = Nothing
End Sub

Private Function GET_RIEPILOGO_QUANTITA_LAVORAZIONE(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim SegnoMovimento As String

GET_RIEPILOGO_QUANTITA_LAVORAZIONE = 0


If ATTIVAZIONE_NUOVO_CALCOLO = True Then
    sSQL = "SELECT IDArticolo, IDFunzione, RV_POQuantitaMovimentata FROM Movimento "
    sSQL = sSQL & "WHERE RV_POIDCaricoMerceRighe=" & IDConferimentoRiga
    sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_POLavorazioneL")
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    Set rs = Cn.OpenResultset(sSQL)
    
    
    
    
    If rs.EOF Then
        GET_RIEPILOGO_QUANTITA_LAVORAZIONE = 0
    Else

        Select Case GET_TIPO_PRODOTTO(rs!IDArticolo)
        
            Case Link_TipoCaloPeso
                GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!RV_POQuantitaMovimentata)
            Case Link_TipoScarto
                GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!RV_POQuantitaMovimentata)
            Case Link_TipoAumentoPeso
                GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE - fnNotNullN(rs!RV_POQuantitaMovimentata)
            Case Else
                GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!RV_POQuantitaMovimentata)
        End Select


        'While Not rs.EOF
        '    SegnoMovimento = RecuperaSegnoPerDisponibilita(fnNotNullN(rs!IDFunzione))
        '    Select Case SegnoMovimento
        '        Case "-"
        '            GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!RV_POQuantitaMovimentata)
        '        Case "+"
        '            GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE - fnNotNullN(rs!RV_POQuantitaMovimentata)
        '        Case Else
        '            GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!RV_POQuantitaMovimentata)
        '    End Select
        'rs.MoveNext
        'Wend
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    


Else



    sSQL = "SELECT SegnoMovimento, Colli, PesoLordo, PesoNetto, Tara, Pezzi "
    sSQL = sSQL & "FROM RV_POLavorazione "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga
    
    
    
    Set rs = Cn.OpenResultset(sSQL)
    
    
    While Not rs.EOF
        Select Case m_DocumentsLink("IDUnitaDiMisura").Value
            Case 1
                If Trim(fnNotNull(rs!SegnoMovimento)) = "+" Then
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE - fnNotNullN(rs!Colli)
                Else
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!Colli)
                End If
            Case 2
                If Trim(fnNotNull(rs!SegnoMovimento)) = "+" Then
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE - fnNotNullN(rs!PesoLordo)
                Else
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!PesoLordo)
                End If
    
            Case 3
                If Trim(fnNotNull(rs!SegnoMovimento)) = "+" Then
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE - fnNotNullN(rs!PesoNetto)
                Else
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!PesoNetto)
                End If
    
            Case 4
                If Trim(fnNotNull(rs!SegnoMovimento)) = "+" Then
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE - fnNotNullN(rs!Tara)
                Else
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!Tara)
                End If
    
            Case 5
                If Trim(fnNotNull(rs!SegnoMovimento)) = "+" Then
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE - fnNotNullN(rs!Pezzi)
                Else
                    GET_RIEPILOGO_QUANTITA_LAVORAZIONE = GET_RIEPILOGO_QUANTITA_LAVORAZIONE + fnNotNullN(rs!Pezzi)
                End If
    
        End Select
    rs.MoveNext
    Wend
    
    
    rs.CloseResultset
    Set rs = Nothing

End If


End Function
Private Function GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = 0

If ATTIVAZIONE_NUOVO_CALCOLO = False Then
''''''''''CALCOLO DELLE LAVORAZIONI MERCE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT SUM(Qta_UM) AS Qta_UM, "
    sSQL = sSQL & "SUM(Colli) AS Colli, "
    sSQL = sSQL & "SUM(PesoLordo) AS PesoLordo, "
    sSQL = sSQL & "SUM(PesoNetto) AS PesoNetto, "
    sSQL = sSQL & "SUM(Tara) AS Tara, "
    sSQL = sSQL & "SUM(Pezzi) AS Pezzi "
    sSQL = sSQL & "FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = 0
    Else
        Select Case m_DocumentsLink("IDUnitaDiMisura").Value
            Case 1
                GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = fnNotNullN(rs!Colli)
            Case 2
                GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = fnNotNullN(rs!PesoLordo)
            Case 3
                GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = fnNotNullN(rs!PesoNetto)
            Case 4
                GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = fnNotNullN(rs!Tara)
            Case 5
                GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = fnNotNullN(rs!Pezzi)
        End Select
            
    End If
    
    rs.CloseResultset
    Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    


Else

    sSQL = "SELECT Sum (RV_POQuantitaMovimentata) AS Quantita FROM Movimento "
    sSQL = sSQL & "WHERE RV_POIDCaricoMerceRighe=" & IDConferimentoRiga
    sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_POAssegnazioneMerce")
    sSQL = sSQL & " AND RV_POTipoRiga=1"
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = 0
    Else
        GET_RIEPILOGO_QUANTITA_ASSEGNAZIONE = fnNotNullN(rs!Quantita)
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End If


End Function
Private Function GET_RIEPILOGO_QUANTITA_PROCESSO(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset



GET_RIEPILOGO_QUANTITA_PROCESSO = 0

If ATTIVAZIONE_NUOVO_CALCOLO = False Then
    sSQL = "SELECT SUM(Quantita) AS Quantita "
    sSQL = sSQL & "FROM RV_POProcessoIVGammaRighe "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDConferimentoRiga
    
    Set rs = Cn.OpenResultset(sSQL)
    
    
    
    If rs.EOF Then
        GET_RIEPILOGO_QUANTITA_PROCESSO = 0
    Else
        GET_RIEPILOGO_QUANTITA_PROCESSO = fnNotNullN(rs!Quantita)
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Else
    
    sSQL = "SELECT Sum (RV_POQuantitaMovimentata) AS Quantita FROM Movimento "
    sSQL = sSQL & "WHERE RV_POIDCaricoMerceRighe=" & IDConferimentoRiga
    sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto("RV_POIVGamma")
    'sSQL = sSQL & " AND RV_POTipoRiga=1 "
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_RIEPILOGO_QUANTITA_PROCESSO = 0
    Else
        GET_RIEPILOGO_QUANTITA_PROCESSO = fnNotNullN(rs!Quantita)
    End If
    
    rs.CloseResultset
    Set rs = Nothing

End If


End Function




Private Function GET_RIEPILOGO_QUANTITA_VENDUTO(IDConferimentoRiga As Long) As Double
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim param As ADODB.Parameter



GET_RIEPILOGO_QUANTITA_VENDUTO = 0


If ATTIVAZIONE_NUOVO_CALCOLO = True Then


    sSQL = "SELECT Sum (RV_POQuantitaMovimentata) AS Quantita FROM Movimento "
    sSQL = sSQL & "WHERE RV_POIDCaricoMerceRighe=" & IDConferimentoRiga
    sSQL = sSQL & " AND (IDTipoOggetto=114 OR IDTipoOggetto=2 OR IDTipoOggetto=8) "
    sSQL = sSQL & " AND RV_POTipoRiga=1 "
    
    Set rs = New ADODB.Recordset
    
    rs.Open sSQL, Cn.InternalConnection
    
    
    If rs.EOF Then
        GET_RIEPILOGO_QUANTITA_VENDUTO = 0
    Else
        GET_RIEPILOGO_QUANTITA_VENDUTO = fnNotNullN(rs!Quantita)
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function

End If


'DOCUMENTO DI TRASPORTO
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale, Sum(Art_numero_colli) as Colli, "
sSQL = sSQL & "Sum(Art_quantita_Pezzi) as Pezzi, Sum(Art_tara) as Tara,Sum(Art_peso) as PesoLordo "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"

'Set rs = Cn.OpenResultset(sSQL)
Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    Select Case Link_UnitaDiMisura_Coop
        Case 1
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Colli)
        Case 2
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!PesoLordo)
        Case 3
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + (fnNotNullN(rs!PesoLordo) - fnNotNullN(rs!Tara))
        Case 4
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Tara)
        Case 5
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Pezzi)
    End Select
    
    
End If

rs.Close
Set rs = Nothing

DoEvents
'FATTURA ACCOMPAGNATORIA
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale, Sum(Art_numero_colli) as Colli, "
sSQL = sSQL & "Sum(Art_quantita_Pezzi) as Pezzi, Sum(Art_tara) as Tara,Sum(Art_peso) as PesoLordo "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"

'Set rs = Cn.OpenResultset(sSQL)
Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    Select Case Link_UnitaDiMisura_Coop
        Case 1
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Colli)
        Case 2
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!PesoLordo)
        Case 3
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + (fnNotNullN(rs!PesoLordo) - fnNotNullN(rs!Tara))
        Case 4
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Tara)
        Case 5
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Pezzi)
    End Select
    
    
End If

rs.Close
Set rs = Nothing

DoEvents

'CORRISPETTIVI
sSQL = "SELECT Sum(Art_quantita_totale) as QtaTotale, Sum(Art_numero_colli) as Colli, "
sSQL = sSQL & "Sum(Art_quantita_Pezzi) as Pezzi, Sum(Art_tara) as Tara,Sum(Art_peso) as PesoLordo "
sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 "
sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND RV_POTipoRiga=1"

'Set rs = Cn.OpenResultset(sSQL)
Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection
If rs.EOF Then
    GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO
Else
    Select Case Link_UnitaDiMisura_Coop
        Case 1
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Colli)
        Case 2
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!PesoLordo)
        Case 3
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + (fnNotNullN(rs!PesoLordo) - fnNotNullN(rs!Tara))
        Case 4
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Tara)
        Case 5
            GET_RIEPILOGO_QUANTITA_VENDUTO = GET_RIEPILOGO_QUANTITA_VENDUTO + fnNotNullN(rs!Pezzi)
    End Select
End If

rs.Close
Set rs = Nothing

DoEvents

End Function

Private Sub Start_Liquidazione()
On Error GoTo ERR_START_LIQUIDAZIONE
Dim sSQL As String
Dim rsConf As DmtOleDbLib.adoResultset

Cn.Execute "DELETE FROM RV_POTMPLiquidazioneRigheEla"

sSQL = "SELECT * FROM RV_POCaricoMerceRighe WHERE IDRV_POCaricoMerceTesta=" & m_Document(m_Document.PrimaryKey)
Set rsConf = Cn.OpenResultset(sSQL)
While Not rsConf.EOF
Screen.MousePointer = 11
    fncElaborazioneLiquidazione fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
Screen.MousePointer = 0
rsConf.MoveNext
Wend
rsConf.CloseResultset
Set rsConf = Nothing

RegistrazioneLiquidazione

Exit Sub

ERR_START_LIQUIDAZIONE:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "START_LIQUIDAZIONE"
End Sub



'''''''''''''''''''''''''CALCOLO LIQUIDAZIONE'''''''''''''''''''''

Private Sub fncElaborazioneLiquidazione(IDCaricoMerceRighe As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsVend As DmtOleDbLib.adoResultset
Dim ArrayLiquidazione(100, 1) As Long
Dim ArrayVendite(100, 4) As Long
Dim RConf As Integer
Dim CConf As Integer
Dim RVend As Integer
Dim CVend As Integer

Cn.Execute "DELETE FROM RV_POTMPLiquidazione"
Cn.Execute "DELETE FROM RV_POTMPLiquidazioneRigheEla"
Cn.Execute "DELETE FROM RV_POTMPLiquidazioneRighe"



sSQL = "SELECT IDRV_POLiquidazione, IDRV_POLiquidazionePeriodo "
sSQL = sSQL & "FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE (IDRV_POCaricoMerceRighe = " & IDCaricoMerceRighe & ") "
sSQL = sSQL & "GROUP BY IDRV_POLiquidazione, IDRV_POLiquidazionePeriodo"
    
    
'Eliminazione array Liquidazione
RConf = 0
CConf = 0

For RConf = 0 To 100
    ArrayLiquidazione(RConf, 0) = 0
    ArrayLiquidazione(RConf, 1) = 0
Next
    


Set rs = Cn.OpenResultset(sSQL)


RConf = 0
While Not rs.EOF
    CConf = 0
        
        ArrayLiquidazione(RConf, 0) = fnNotNullN(rs!IDRV_POLiquidazione)
        ArrayLiquidazione(RConf, 1) = fnNotNullN(rs!IDRV_POLiquidazionePeriodo)
    
    RConf = RConf + 1
    rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
    
    RConf = 0
    CConf = 0

    For RConf = 0 To 100
        If ArrayLiquidazione(RConf, 0) > 0 Then
            sSQL = "SELECT RV_POLiquidazionePeriodo.IDRV_POLiquidazionePeriodo, RV_POLiquidazionePeriodo.Periodo, RV_POLiquidazionePeriodo.NumeroLiquidazione, "
            sSQL = sSQL & "RV_POLiquidazionePeriodo.DataInizio, RV_POLiquidazionePeriodo.DataFine, RV_POLiquidazionePeriodo.IDTipoImportoArticolo,"
            sSQL = sSQL & "RV_POLiquidazionePeriodo.IDSocio , Anagrafica.Anagrafica, Anagrafica.Nome, "
            sSQL = sSQL & "RV_POLiquidazionePeriodo.IDTipoImportoDocumento, RV_POLiquidazionePeriodo.IDTipoQuantita, "
            sSQL = sSQL & "RV_POLiquidazionePeriodo.ArticoliDiQuadratura, RV_POLiquidazionePeriodo.IDTipoLiquidazione,RV_POLiquidazionePeriodo.IDTipoPrezzoMedio "
            sSQL = sSQL & "FROM RV_POLiquidazionePeriodo LEFT OUTER JOIN "
            sSQL = sSQL & "Anagrafica ON RV_POLiquidazionePeriodo.IDSocio = Anagrafica.IDAnagrafica "
            sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & ArrayLiquidazione(RConf, 1)
        
            Set rs = Cn.OpenResultset(sSQL)
            
            If rs.EOF = False Then
                DATA_INIZIO = fnNotNull(rs!DataInizio)
                DATA_FINE = fnNotNull(rs!DataFine)
                LINK_SOCIO = fnNotNullN(rs!IDSocio)
                TIPO_IMPORTO_ARTICOLO = fnNotNullN(rs!IDTipoImportoArticolo)
                TIPO_IMPORTO_DOCUMENTO = fnNotNullN(rs!IDTipoImportoDocumento)
                TIPO_QUANTITA = fnNotNullN(rs!IDTipoQuantita)
                ARTICOLI_DI_QUAD = fnNormBoolean(fnNotNullN(rs!ArticoliDiQuadratura))
                TIPO_LIQUIDAZIONE = fnNotNullN(rs!IDTipoLiquidazione)
                TIPO_CALCOLO_PREZZO_MEDIO = fnNotNullN(rs!IDTipoPrezzoMedio)
                LINK_PERIODO = ArrayLiquidazione(RConf, 1)
                
                EsecuzioneElaborazione IDCaricoMerceRighe, ArrayLiquidazione(RConf, 0), ArrayLiquidazione(RConf, 1)
                
            End If
        End If

    Next

End Sub
Public Sub RegistrazioneLiquidazione()
    EliminazioneDatiLiquidazioneRegistrata
    
    Registrazione
    
    AggiornamentoTestataLiquidazione
    
End Sub

Private Sub EliminazioneDatiLiquidazioneRegistrata()
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT IDRV_POLiquidazionePeriodo, IDRV_POTMPLiquidazione, IDRV_POCaricoMerceRighe "
sSQL = sSQL & "FROM  RV_POTMPLiquidazioneRigheEla "
sSQL = sSQL & "GROUP BY IDRV_POLiquidazionePeriodo, IDRV_POTMPLiquidazione, IDRV_POCaricoMerceRighe"

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF
    
    sSQL = "DELETE FROM RV_POLiquidazioneRigheEla "
    sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POTMPLiquidazione)
    sSQL = sSQL & " AND IDRV_POLiquidazionePeriodo=" & fnNotNullN(rs!IDRV_POLiquidazionePeriodo)
    sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    
    Cn.Execute sSQL

rs.MoveNext
Wend

rs.Close
Set rs = Nothing



End Sub
Private Sub Registrazione()
Dim sSQL As String
Dim rsEla As ADODB.Recordset
Dim rsUPD As ADODB.Recordset
Dim I As Integer

sSQL = "SELECT * FROM RV_POTMPLiquidazioneRigheEla "

Set rsEla = New ADODB.Recordset
Set rsUPD = New ADODB.Recordset

rsEla.Open sSQL, Cn.InternalConnection
rsUPD.Open "RV_POLiquidazioneRigheEla", Cn.InternalConnection, adOpenDynamic, adLockPessimistic
While Not rsEla.EOF
    rsUPD.AddNew
        rsUPD!IDRV_POLiquidazioneRigheEla = fnGetNewKey("RV_POLiquidazioneRigheEla", "IDRV_POLiquidazioneRigheEla")
        rsUPD!IDRV_POLiquidazione = rsEla!IDRV_POTMPLiquidazione
        rsUPD!IDRV_POLiquidazionePeriodo = rsEla!IDRV_POLiquidazionePeriodo
        For I = 2 To rsEla.Fields.Count - 1
            rsUPD.Fields(rsEla.Fields(I).Name) = rsEla.Fields(I).Value
        Next
    rsUPD.Update
rsEla.MoveNext
Wend
rsUPD.Close
Set rsUPD = Nothing

rsEla.Close
Set rsEla = Nothing
End Sub
Public Sub AggiornamentoTestataLiquidazione()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAgg As DmtOleDbLib.adoResultset
Dim NettoLiquidazione As Double

sSQL = "SELECT IDRV_POTMPLiquidazione, IDRV_POLiquidazionePeriodo "
sSQL = sSQL & "FROM  RV_POTMPLiquidazioneRigheEla "
sSQL = sSQL & "GROUP BY IDRV_POTMPLiquidazione, IDRV_POLiquidazionePeriodo "


Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "SELECT SUM(ImponibileDaReg) AS TotaleDocumento, SUM(ImpostaDaReg) AS ImpostaDocumento, SUM(ImportoLordoDaReg) AS TotaleDocumentoLordoIva, "
    sSQL = sSQL & "SUM(TrattenutePerLavorazione) AS TrattenutaPerLavorazione, SUM(TrattenuteGenerali) AS TrattenutaGenerale, SUM(TrattenuteTotali) "
    sSQL = sSQL & "AS TrattenutaTotale "
    sSQL = sSQL & "FROM RV_POLiquidazioneRigheEla "
    sSQL = sSQL & "WHERE (IDRV_POLiquidazione = " & fnNotNullN(rs!IDRV_POTMPLiquidazione) & ")"
    
    Set rsAgg = Cn.OpenResultset(sSQL)
    
    If rsAgg.EOF = False Then
        TIPO_IMPORTO_DOCUMENTO = GET_PARAMETRO_TIPO_IMPORTO_DOCUMENTO(rs!IDRV_POLiquidazionePeriodo)
        If TIPO_IMPORTO_DOCUMENTO = 1 Then
            NettoLiquidazione = rsAgg!TotaleDocumento - rsAgg!TrattenutaTotale
        Else
            NettoLiquidazione = rsAgg!TotaleDocumentoLordoIva - rsAgg!TrattenutaTotale
        End If
            
        sSQL = "UPDATE RV_POLiquidazione SET "
        sSQL = sSQL & "TotaleDocumento = " & fnNormNumber(rsAgg!TotaleDocumento) & ", "
        sSQL = sSQL & "TotaleIva = " & fnNormNumber(rsAgg!ImpostaDocumento) & ", "
        sSQL = sSQL & "TotaleDocumentoLordoIva = " & fnNormNumber(rsAgg!TotaleDocumentoLordoIva) & ", "
        sSQL = sSQL & "TotaleTrattenutaPerLavorazione = " & fnNormNumber(rsAgg!TrattenutaPerLavorazione) & ", "
        sSQL = sSQL & "TotaleTrattenutaGenerale = " & fnNormNumber(rsAgg!TrattenutaGenerale) & ", "
        sSQL = sSQL & "TotaleTrattenuta = " & fnNormNumber(rsAgg!TrattenutaTotale) & ", "
        sSQL = sSQL & "NettoLiquidazione = " & fnNormNumber(NettoLiquidazione) & " "
        sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & fnNotNullN(rs!IDRV_POTMPLiquidazione)
        
        Cn.Execute sSQL
    End If
    
    
    rsAgg.CloseResultset
    Set rsAgg = Nothing
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

    
End Sub
Private Function GET_PARAMETRO_TIPO_IMPORTO_DOCUMENTO(IDPeriodoLiquidazione As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoImportoDocumento FROM RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & IDPeriodoLiquidazione


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PARAMETRO_TIPO_IMPORTO_DOCUMENTO = 1
Else
    GET_PARAMETRO_TIPO_IMPORTO_DOCUMENTO = fnNotNullN(rs!IDTipoImportoDocumento)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_FUNZIONE(Gestore As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Funzione.IDFunzione "
sSQL = sSQL & "FROM Funzione INNER JOIN "
sSQL = sSQL & "TipoOggetto ON Funzione.IDTipoOggetto = TipoOggetto.IDTipoOggetto INNER JOIN "
sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(Gestore)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_FUNZIONE = 0
Else
    GET_FUNZIONE = fnNotNullN(rs!IDFunzione)
End If


rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLA_ESISTENZA_LIQUIDAZIONE(IDRigaConferimento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'DOCUMENTO DI TRASPORTO
sSQL = "SELECT * FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLA_ESISTENZA_LIQUIDAZIONE = False
Else
    GET_CONTROLLA_ESISTENZA_LIQUIDAZIONE = True
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub RefreshArray()
Dim I As Integer

For I = 0 To 150
    ArrayDelete(I) = 0
    ArrayDeleteConferimento(I) = 0
Next
End Sub
Private Sub StampaDocumentoModalitaTabellare()
Dim sSQL As String
Set oReport = New dmtReportLib.dmtReport
    Set oReport.Connection = Cn
    If MenuOptions.DBType = 1 Then
        'parametri di accesso al database ACCESS
        oReport.Password = "dmt192981046"
        oReport.User = "admin"
    Else
        'parametri di accesso al database SQL Server
        oReport.Password = TheApp.Password
        oReport.User = TheApp.User
    End If
        
        'Imposta l'idfiliale di appartenenza del documento da stampare
            oReport.BranchID = TheApp.Branch
        
        'Imposta l'identificativo del tipo di documento
            'IDTipoOggettoPrg = fncIDTipoOggettoPrg
            oReport.DocTypeID = m_DocType.ID
            
 
            sSQL = fnFillDocTypeCondition
    
            oReport.Where = Trim(sSQL)
            
            'IDReport = fncTrovaReport(Me.cboReportDisponibili.Text, IDTipoOggettoPrg)
            
            'If IDReport > 0 Then
                'fncImpostaDefaultReport (IDReport)
                'oReport.Preview 0, 0, 0
            'Else
             '   MsgBox "ATTENZIONE!!!!" & vbCrLf & "Il report non è stato trovato!", vbCritical, "Impossibile stampare"
            'E 'nd If
        
End Sub
Private Function GET_NUMERO_LOTTO() As Long
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT NumerazioneLottoConferimento FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    GET_NUMERO_LOTTO = 0
Else
    If fnNotNullN(rs!NumerazioneLottoConferimento) + 1 < 10000000 Then
        GET_NUMERO_LOTTO = fnNotNullN(rs!NumerazioneLottoConferimento) + 1
        rs!NumerazioneLottoConferimento = GET_NUMERO_LOTTO
    Else
        GET_NUMERO_LOTTO = 1
        rs!NumerazioneLottoConferimento = 1
    
    End If
    rs.Update
End If
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

Private Function GET_CERTIFICAZIONE_LOTTO_CAMPAGNA(IDLottoDiCampagna As Long, NomeCampo As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_PO01_LottoCampagna INNER JOIN "
sSQL = sSQL & "RV_PO01_Certificazione ON "
sSQL = sSQL & "RV_PO01_LottoCampagna.IDRV_PO01_CertificazioneSocio = RV_PO01_Certificazione.IDRV_PO01_Certificazione " 'INNER JOIN "
'sSQL = sSQL & "RV_PO01_Certificazione ON RV_PO01_CertificazioneSocio.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione "
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

Private Sub GENERA_FILTRO_PER_TIPO_OGGETTO()
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim I As Long
Dim Filtro As String

sSQL = "DELETE FROM RV_POFiltro "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoOggetto=" & fnGetTipoOggetto

Cn.Execute sSQL
    
Filtro = ""

For I = 1 To Me.BrwMain.Conditions.Count
    If Len(Me.BrwMain.Conditions(I).FromValue) > 0 Then
        If Len(Me.BrwMain.Conditions(I).ToValue) > 0 Then
            Filtro = Filtro & Me.BrwMain.Conditions(I).Name & ": dal " & Me.BrwMain.Conditions(I).FromValue
            Filtro = Filtro & " Al " & Me.BrwMain.Conditions(I).ToValue
        Else
            Filtro = Filtro & Me.BrwMain.Conditions(I).Name & ": " & Me.BrwMain.Conditions(I).FromValue
        End If
        Filtro = Filtro & " - "
    End If
Next

sSQL = "SELECT * FROM RV_POFiltro"
Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenDynamic, adLockPessimistic

rs.AddNew
    rs!IDUtente = TheApp.IDUser
    rs!IDAzienda = TheApp.IDFirm
    rs!IDTipoOggetto = fnGetTipoOggetto
    rs!Filtro = Filtro
rs.Update

rs.Close
Set rs = Nothing
End Sub
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

sSQL = "SELECT Count(IDRV_POCaricoMerceTesta) As NumeroInserimenti "
sSQL = sSQL & "FROM RV_POCaricoMerceTesta "
sSQL = sSQL & "WHERE IDTipoDocumentoCoop=1"
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
Private Sub ParametroLottoObbligatorio()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT LottoCampagnaObbligatorio FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LOTTO_CAMPAGNA_OBBLIGATORIO = fnNotNullN(rs!LottoCampagnaObbligatorio)
Else
    LOTTO_CAMPAGNA_OBBLIGATORIO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroTipoArrotondamento()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoArrotondamentoConferimento FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_TIPO_ARROTONDAMENTO = fnNotNullN(rs!IDTipoArrotondamentoConferimento)
Else
    LINK_TIPO_ARROTONDAMENTO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub


Private Sub ParametroTipoSceltaArticoloConferito()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoSceltaArticoloLottoCampagna FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    LINK_TIPO_ARTICOLO_CONFERITO = fnNotNullN(rs!IDRV_POTipoSceltaArticoloLottoCampagna)
Else
    LINK_TIPO_ARTICOLO_CONFERITO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GET_LINK_ORDINAMENTO(Tabella As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(Link_Ordinamento) AS Link_Ordinamento "
sSQL = sSQL & "FROM " & Tabella
sSQL = sSQL & " WHERE " & m_Document.PrimaryKey & "=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ORDINAMENTO = 1
Else
    GET_LINK_ORDINAMENTO = fnNotNullN(rs!LINK_ORDINAMENTO) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function RecuperaSegnoPerDisponibilita(IDFunzione) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


    sSQL = "SELECT Funzione.Funzione, ProcessoPerFunzione.IDProcesso, ProcessoPerFunzione.IDFunzione,"
    sSQL = sSQL & "ContatoreArticoloPerMagazzino.PartecipaGiacenza, ContatoreArticoloPerMagazzino.PartecipaDisponibilita, Processo.Processo,"
    sSQL = sSQL & "ContatorePerProcesso.IDProcessoPerFunzione, ContatorePerProcesso.Numero,ContatorePerProcesso.Quantita,"
    sSQL = sSQL & "ContatorePerProcesso.Valore, ContatorePerProcesso.DataVariazione, ContatoreArticolo.ContatoreArticolo, Processo.IDTipoProcesso,"
    sSQL = sSQL & "ContatoreArticoloPerMagazzino.IDMagazzino "
    sSQL = sSQL & "FROM ContatoreArticolo LEFT OUTER JOIN "
    sSQL = sSQL & "ContatoreArticoloPerMagazzino ON "
    sSQL = sSQL & "ContatoreArticolo.IDContatoreArticolo = ContatoreArticoloPerMagazzino.IDContatoreArticolo RIGHT OUTER JOIN "
    sSQL = sSQL & "Funzione INNER JOIN "
    sSQL = sSQL & "ProcessoPerFunzione ON Funzione.IDFunzione = ProcessoPerFunzione.IDFunzione INNER JOIN "
    sSQL = sSQL & "Processo ON ProcessoPerFunzione.IDProcesso = Processo.IDProcesso LEFT OUTER JOIN "
    sSQL = sSQL & "ContatorePerProcesso ON ProcessoPerFunzione.IDProcessoPerFunzione = ContatorePerProcesso.IDProcessoPerFunzione ON "
    sSQL = sSQL & "ContatoreArticolo.IDContatoreArticolo = ContatorePerProcesso.IDContatoreArticolo "
    sSQL = sSQL & "WHERE (Processo.IDTipoProcesso = 2) And (ProcessoPerFunzione.IDFunzione =" & IDFunzione & ") And (ContatoreArticoloPerMagazzino.IDMagazzino = " & Me.cboMagazzinoConf.CurrentID & ")"


Set rs = Cn.OpenResultset(sSQL)
    
If rs.EOF Then
    RecuperaSegnoPerDisponibilita = ""
Else
    Select Case fnNotNull(rs!PartecipaDisponibilita)
        Case "-"
            RecuperaSegnoPerDisponibilita = "-"
        Case "+"
            RecuperaSegnoPerDisponibilita = "+"
        Case ""
            RecuperaSegnoPerDisponibilita = "-"
    End Select
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
            GET_FUNZIONE_MAGAZZINO = Link_CausaleCarico
        Case 2 'Scarico
            GET_FUNZIONE_MAGAZZINO = Link_CausaleScarico
    End Select
Else
    If fnNotNullN(rs!IDFunzione) = 0 Then
        Select Case IDTipoProcesso
            Case 1 'Carico
                GET_FUNZIONE_MAGAZZINO = Link_CausaleCarico
            Case 2 'Scarico
                GET_FUNZIONE_MAGAZZINO = Link_CausaleScarico
        End Select
    Else
        GET_FUNZIONE_MAGAZZINO = fnNotNullN(rs!IDFunzione)
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_TIPO_PRODOTTO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoProdotto FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPO_PRODOTTO = 0
Else
    GET_TIPO_PRODOTTO = fnNotNullN(rs!IDTipoProdotto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_PARAMETRI_SOCIO(IDEsercizio As Long, IDAnagrafica As Long, NomeCampo As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT " & NomeCampo
sSQL = sSQL & " FROM RV_PONumerazionePerSocio "
sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizio
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    If rs.adoColumns(NomeCampo).DataType = 4 Then
        GET_PARAMETRI_SOCIO = 1
    Else
        GET_PARAMETRI_SOCIO = ""
    End If

Else
    If rs.adoColumns(NomeCampo).DataType = 4 Then
        GET_PARAMETRI_SOCIO = fnNotNullN(rs.adoColumns(NomeCampo).Value)
        If GET_PARAMETRI_SOCIO = 0 Then
            GET_PARAMETRI_SOCIO = 1
        End If
    Else
        GET_PARAMETRI_SOCIO = fnNotNull(rs.adoColumns(NomeCampo).Value)
    End If
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Sub AGGIORNA_PROGRESSIVO_CONFERIMENTO(IDAnagrafica As Long, IDEsercizio As Long, NumeroConferimento As Long, PrefissoNumeroConferimento As String)
On Error GoTo ERR_AGGIORNA_PROGRESSIVO_CONFERIMENTO
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_PONumerazionePerSocio "
sSQL = sSQL & "WHERE IDAzienda=" & m_App.IDFirm
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizio
sSQL = sSQL & " AND IDAnagrafica=" & IDAnagrafica

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rs.EOF Then
    If NumeroConferimento >= fnNotNullN(rs!NumeroConferimento) Then
        rs!NumeroConferimento = NumeroConferimento + 1
        rs.Update
    End If
Else
    rs.AddNew
        rs!IDRV_PONumerazionePerSocio = fnGetNewKey("RV_PONumerazionePerSocio", "IDRV_PONumerazionePerSocio")
        rs!IDAnagrafica = IDAnagrafica
        rs!IDEsercizio = IDEsercizio
        rs!IDAzienda = m_App.IDFirm
        rs!Numero = 1
        rs!Prefisso = GET_PREFISSO_SOCIO_ESERCIZIO_PREC(IDEsercizio, IDAnagrafica)
        rs!NumeroConferimento = NumeroConferimento + 1
        rs!PrefissoNumeroConferimento = PrefissoNumeroConferimento
    rs.Update
End If

rs.Close
Set rs = Nothing

Exit Sub
ERR_AGGIORNA_PROGRESSIVO_CONFERIMENTO:
    MsgBox Err.Description, vbCritical, "AGGIORNA_PROGRESSIVO_CONFERIMENTO"
End Sub
Private Function GET_CONTROLLA_ESISTENZA_COLLEGAMENTO(IDRigaConferimento As Long, IDTipoOggetto As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False

If ATTIVAZIONE_NUOVO_CALCOLO = True Then
    sSQL = "SELECT IDMovimento FROM Movimento "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND RV_POIDCaricoMerceRighe=" & IDRigaConferimento
    sSQL = sSQL & " AND IDTipoOggetto<>" & IDTipoOggetto
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
    End If
    
    rs.CloseResultset
    Set rs = Nothing
Else
    ''''''''''''''''''CONTROLLO ASSEGNAZIONE
    sSQL = "SELECT IDRV_POAssegnazioneMerce FROM RV_POAssegnazioneMerce "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
        rs.CloseResultset
        Set rs = Nothing
        Exit Function
    End If
    
    rs.CloseResultset
    Set rs = Nothing

    
    ''''''''''''''''''''''''''''''''''''''''
     ''''''''''''''''''CONTROLLO SCARTI
    sSQL = "SELECT IDRV_POLavorazione FROM RV_POLavorazione "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
        rs.CloseResultset
        Set rs = Nothing
        Exit Function
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''CONTROLLO DDT''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT IDValoriOggettoDettaglio FROM ValoriOggettoDettaglio0004 "
    sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento

    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
        rs.CloseResultset
        Set rs = Nothing
        Exit Function
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''CONTROLLO FATTURA ACCOMPAGNATORIA
    sSQL = "SELECT IDValoriOggettoDettaglio FROM ValoriOggettoDettaglio0001 "
    sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento

    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
        rs.CloseResultset
        Set rs = Nothing
        Exit Function
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''CONTROLLO CORRISPETTIVI
    sSQL = "SELECT IDValoriOggettoDettaglio FROM ValoriOggettoDettaglio0034 "
    sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento

    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
        rs.CloseResultset
        Set rs = Nothing
        Exit Function
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''CONTROLLO NOTA DI CREDITO
    sSQL = "SELECT IDValoriOggettoDettaglio FROM ValoriOggettoDettaglio0016 "
    sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento

    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
        rs.CloseResultset
        Set rs = Nothing
        Exit Function
    End If
    
    rs.CloseResultset
    Set rs = Nothing
        
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''CONTROLLO NOTA DI DEBITO
     sSQL = "SELECT IDValoriOggettoDettaglio FROM ValoriOggettoDettaglio0007 "
    sSQL = sSQL & "WHERE RV_POIDConferimentoRighe=" & IDRigaConferimento

    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = False
    Else
        GET_CONTROLLA_ESISTENZA_COLLEGAMENTO = True
        rs.CloseResultset
        Set rs = Nothing
        Exit Function
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    ''''''''''''''''''''''''''''''''''''''''

End If



End Function
Public Function GET_NOMECOMPUTER() As String
Dim dwLen As Long
Dim strString As String
Const MAX_COMPUTERNAME_LENGTH As Long = 31
    
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name
    GetComputerName strString, dwLen
    'get only the actual data
    strString = Left(strString, dwLen)
    'Show the computer name
    GET_NOMECOMPUTER = strString
End Function

Function GET_NOMEUTENTE() As String
    Dim strString As String
    Dim lunghezzaStringa As Long
    lunghezzaStringa = 32
    strString = String(lunghezzaStringa, " ")
    GetUserName strString, lunghezzaStringa
    strString = Left(strString, lunghezzaStringa)
    GET_NOMEUTENTE = strString
    GET_NOMEUTENTE = Mid(GET_NOMEUTENTE, 1, Len(GET_NOMEUTENTE) - 1)
End Function

Private Sub cboUtente_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Descrizione "
sSQL = sSQL & "FROM Utente "
sSQL = sSQL & "WHERE IDUtente=" & Me.cboUtente.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtCodiceUtente.Text = ""
Else
    Me.txtCodiceUtente.Text = fnNotNull(rs!Descrizione)
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Function GET_DATI_LOTTO_CAMPAGNA(IDLottoDiCampagna As Long) As String
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_PO01_LottoCampagna.IDRV_PO01_TipoProduzione, RV_PO01_LottoCampagna.IDRV_PO01_Varieta, "
sSQL = sSQL & "RV_PO01_LottoCampagna.IDRV_PO01_FamigliaProdotti, RV_PO01_FamigliaProdotti.FamigliaProdotti, RV_PO01_Varieta.Varieta, "
sSQL = sSQL & "RV_PO01_TipoProduzione.TipoProduzione, RV_PO01_LottoCampagna.DataSbloccoLotto "
sSQL = sSQL & "FROM RV_PO01_LottoCampagna LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_FamigliaProdotti ON "
sSQL = sSQL & "RV_PO01_LottoCampagna.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_Varieta ON RV_PO01_LottoCampagna.IDRV_PO01_Varieta = RV_PO01_Varieta.IDRV_PO01_Varieta LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_TipoProduzione ON RV_PO01_LottoCampagna.IDRV_PO01_TipoProduzione = RV_PO01_TipoProduzione.IDRV_PO01_TipoProduzione "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoDiCampagna

Set rs = Cn.OpenResultset(sSQL)

LINK_VARIETA_LOTTO_CAMPAGNA = 0
LINK_FAMIGLIA_LOTTO_CAMPAGNA = 0
LINK_TIPO_PRODUZIONE_LOTTO_CAMPAGNA = 0

VARIETA_LOTTO_CAMPAGNA = ""
FAMIGLIA_LOTTO_CAMPAGNA = ""
TIPO_PRODUZIONE_LOTTO_CAMPAGNA = ""


If rs.EOF Then
    LINK_VARIETA_LOTTO_CAMPAGNA = 0
    LINK_FAMIGLIA_LOTTO_CAMPAGNA = 0
    LINK_TIPO_PRODUZIONE_LOTTO_CAMPAGNA = 0
    VARIETA_LOTTO_CAMPAGNA = ""
    FAMIGLIA_LOTTO_CAMPAGNA = ""
    TIPO_PRODUZIONE_LOTTO_CAMPAGNA = ""
    DATA_SBLOCCO_LOTTO_CAMPAGNA = ""
Else
    LINK_VARIETA_LOTTO_CAMPAGNA = fnNotNullN(rs!IDRV_PO01_Varieta)
    LINK_FAMIGLIA_LOTTO_CAMPAGNA = fnNotNullN(rs!IDRV_PO01_FamigliaProdotti)
    LINK_TIPO_PRODUZIONE_LOTTO_CAMPAGNA = fnNotNullN(rs!IDRV_PO01_TipoProduzione)
    VARIETA_LOTTO_CAMPAGNA = fnNotNull(rs!Varieta)
    FAMIGLIA_LOTTO_CAMPAGNA = fnNotNull(rs!FamigliaProdotti)
    TIPO_PRODUZIONE_LOTTO_CAMPAGNA = fnNotNull(rs!TipoProduzione)
    DATA_SBLOCCO_LOTTO_CAMPAGNA = fnNotNull(rs!DataSbloccoLotto)
End If

rs.CloseResultset
Set rs = Nothing


Me.txtVarietà.Text = VARIETA_LOTTO_CAMPAGNA
Me.txtFamiglia.Text = FAMIGLIA_LOTTO_CAMPAGNA
Me.txtTipoProduzione.Text = TIPO_PRODUZIONE_LOTTO_CAMPAGNA
Me.txtDataSbloccoLotto.Text = DATA_SBLOCCO_LOTTO_CAMPAGNA
End Function
Private Function GET_UM_ARTICOLO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDUnitaDiMisuraAcquisto "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_UM_ARTICOLO = 0
Else
    GET_UM_ARTICOLO = fnNotNullN(rs!IDUnitaDiMisuraAcquisto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_DESCRIZIONE_LOTTO_CAMPAGNA(IDLottoCampagna As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'If GET_ESISTENZA_BIO = False Then
'    GET_DESCRIZIONE_LOTTO_CAMPAGNA = ""
'End If

sSQL = "SELECT DescrizioneLotto "
sSQL = sSQL & "FROM RV_PO01_LottoCampagna "
sSQL = sSQL & "WHERE IDRV_PO01_LottoCampagna=" & IDLottoCampagna

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_LOTTO_CAMPAGNA = ""
Else
    GET_DESCRIZIONE_LOTTO_CAMPAGNA = fnNotNull(rs!DescrizioneLotto)
End If


rs.CloseResultset
Set rs = Nothing


End Function
Private Function GET_CONTROLLO_COLLEGAMENTI_CONFERIMENTO(IDCaricoMerceTesta As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsConf As DmtOleDbLib.adoResultset
Dim rsnew As ADODB.Recordset
Dim Controllo_Esistenza As Boolean

Controllo_Esistenza = False


'ELIMINAZIONE DATI TEMPORANEI''''''''''''''''''''''''''''''''''''''''''
sSQL = "DELETE FROM RV_POTMPCollegamentiConferimento "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDCaricoMerceTesta
Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDCaricoMerceTesta

Set rsConf = Cn.OpenResultset(sSQL)

Set rsnew = New ADODB.Recordset

rsnew.Open "RV_POTMPCollegamentiConferimento", Cn.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rsConf.EOF
    ''''''''''''''''''''''''''DOCUMENTO DI TRASPORTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "SELECT ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe, ValoriOggettoPerTipo0002.Doc_data, ValoriOggettoPerTipo0002.Doc_numero, "
    sSQL = sSQL & "ValoriOggettoPerTipo0002.IDOggetto "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0004 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0002 ON ValoriOggettoDettaglio0004.IDOggetto = ValoriOggettoPerTipo0002.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0004.RV_POIDConferimentoRighe=" & fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0004.RV_POTipoRiga=1"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If GET_ESISTENZA_SEMAFORO(fnNotNullN(rs!IDOggetto), fnGetTipoOggetto("RV_PODDTL"), "Documento di trasporto", fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), IDCaricoMerceTesta, rsnew, fnNotNull(rs!Doc_Numero), fnNotNull(rs!Doc_Data)) = True Then
            Controllo_Esistenza = True
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    ''''FATTURA ACCOMPAGNATORIA
    sSQL = "SELECT ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe, ValoriOggettoPerTipo0072.Doc_data, ValoriOggettoPerTipo0072.Doc_numero, "
    sSQL = sSQL & "ValoriOggettoPerTipo0072.IDOggetto "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0001 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0072 ON ValoriOggettoDettaglio0001.IDOggetto = ValoriOggettoPerTipo0072.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0001.RV_POIDConferimentoRighe=" & fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0001.RV_POTipoRiga=1"
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If GET_ESISTENZA_SEMAFORO(fnNotNullN(rs!IDOggetto), fnGetTipoOggetto("RV_POFAL"), "Fattura accompagnatoria", fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), IDCaricoMerceTesta, rsnew, fnNotNull(rs!Doc_Numero), fnNotNull(rs!Doc_Data)) = True Then
            Controllo_Esistenza = True
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    
    ''''SCONTRINO NON FISCALE
    sSQL = "SELECT ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe, ValoriOggettoPerTipo0008.Doc_data, ValoriOggettoPerTipo0008.Doc_numero, "
    sSQL = sSQL & "ValoriOggettoPerTipo0008.IDOggetto "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0034 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo0008 ON ValoriOggettoDettaglio0034.IDOggetto = ValoriOggettoPerTipo0008.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0034.RV_POIDConferimentoRighe=" & fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
    sSQL = sSQL & " AND ValoriOggettoDettaglio0034.RV_POTipoRiga=1"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If GET_ESISTENZA_SEMAFORO(fnNotNullN(rs!IDOggetto), fnGetTipoOggetto("RV_POSNFL"), "Corrispettivi", fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), IDCaricoMerceTesta, rsnew, fnNotNull(rs!Doc_Numero), fnNotNull(rs!Doc_Data)) = True Then
            Controllo_Esistenza = True
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing


    
    ''''NOTA DI CREDITO
    sSQL = "SELECT ValoriOggettoDettaglio0016.RV_POIDConferimentoRighe, ValoriOggettoPerTipo000B.Doc_data, ValoriOggettoPerTipo000B.Doc_numero, "
    sSQL = sSQL & "ValoriOggettoPerTipo000B.IDOggetto "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0016 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo000B ON ValoriOggettoDettaglio0016.IDOggetto = ValoriOggettoPerTipo000B.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0016.RV_POIDConferimentoRighe=" & fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If GET_ESISTENZA_SEMAFORO(fnNotNullN(rs!IDOggetto), fnGetTipoOggetto("RV_PONCL"), "Nota di credito", fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), IDCaricoMerceTesta, rsnew, fnNotNull(rs!Doc_Numero), fnNotNull(rs!Doc_Data)) = True Then
            Controllo_Esistenza = True
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing


    ''''NOTA DI DEBITO
    sSQL = "SELECT ValoriOggettoDettaglio0007.RV_POIDConferimentoRighe, ValoriOggettoPerTipo006B.Doc_data, ValoriOggettoPerTipo006B.Doc_numero, "
    sSQL = sSQL & "ValoriOggettoPerTipo006B.IDOggetto "
    sSQL = sSQL & "FROM ValoriOggettoDettaglio0007 INNER JOIN "
    sSQL = sSQL & "ValoriOggettoPerTipo006B ON ValoriOggettoDettaglio0007.IDOggetto = ValoriOggettoPerTipo006B.IDOggetto "
    sSQL = sSQL & "WHERE ValoriOggettoDettaglio0007.RV_POIDConferimentoRighe=" & fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If GET_ESISTENZA_SEMAFORO(fnNotNullN(rs!IDOggetto), fnGetTipoOggetto("RV_POFIL"), "Nota di debito", fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), IDCaricoMerceTesta, rsnew, fnNotNull(rs!Doc_Numero), fnNotNull(rs!Doc_Data)) = True Then
            Controllo_Esistenza = True
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    ''''ASSEGAZIONE MERCE
    sSQL = "SELECT IDRV_POCaricoMerceRighe "
    sSQL = sSQL & "FROM RV_PORepLavorazione "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If GET_ESISTENZA_SEMAFORO(fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), fnGetTipoOggetto("RV_POAssegnazioneMerce"), "Lavorazione", fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), IDCaricoMerceTesta, rsnew, "", "") = True Then
            Controllo_Esistenza = True
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    ''''SCARTI DI LAVORAZIONE
    sSQL = "SELECT IDRV_POCaricoMerceRighe "
    sSQL = sSQL & "FROM RV_POLavorazione "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & fnNotNullN(rsConf!IDRV_POCaricoMerceRighe)
    
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If GET_ESISTENZA_SEMAFORO(fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), fnGetTipoOggetto("RV_POLavorazioneL"), "Scarti di lavorazione", fnNotNullN(rsConf!IDRV_POCaricoMerceRighe), IDCaricoMerceTesta, rsnew, "", "") = True Then
            Controllo_Esistenza = True
        End If
    rs.MoveNext
    Wend
    
    rs.CloseResultset
    Set rs = Nothing
    
    
rsConf.MoveNext
Wend

rsnew.Close
Set rsnew = Nothing

rsConf.CloseResultset
Set rsConf = Nothing

GET_CONTROLLO_COLLEGAMENTI_CONFERIMENTO = Controllo_Esistenza
End Function
Private Function GET_ESISTENZA_SEMAFORO(IDOggetto As Long, IDTipoOggetto As Long, Oggetto As String, IDCaricoMerceRighe As Long, IDCaricoMerceTesta As Long, rsnew As ADODB.Recordset, NumeroDocumento As String, DataDocumento As String) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Semaforo.*, Utente.Utente "
sSQL = sSQL & "FROM Semaforo INNER JOIN "
sSQL = sSQL & "Utente ON Semaforo.IDUtente = Utente.IDUtente "
sSQL = sSQL & "WHERE IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDTipoOggetto=" & IDTipoOggetto

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_ESISTENZA_SEMAFORO = False
Else
    GET_ESISTENZA_SEMAFORO = True
    
    While Not rs.EOF
        rsnew.AddNew
            rsnew!IDRV_POTMPCollegamentiConferimento = fnGetNewKey("RV_POTMPCollegamentiConferimento", "IDRV_POTMPCollegamentiConferimento")
            rsnew!IDRV_POCaricoMerceRighe = IDCaricoMerceRighe
            rsnew!IDRV_POcaricoMercetesta = IDCaricoMerceTesta
            rsnew!IDOggetto = IDOggetto
            rsnew!IDTipoOggetto = IDTipoOggetto
            rsnew!IDUtente = fnNotNullN(rs!IDUtente)
            rsnew!Utente = fnNotNull(rs!Utente)
            rsnew!Oggetto = Oggetto
            rsnew!NumeroDocumento = NumeroDocumento
            If DataDocumento <> "" Then
                rsnew!DataDocumento = DataDocumento
            End If
            rsnew!MacchinaPC = GET_NOME_MACCHINA(rs!IDUtente)
        rsnew.Update
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing



End Function
Private Function GET_NOME_MACCHINA(IDUtente As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MachineID FROM Semaforo "
sSQL = sSQL & "WHERE IDUtente=" & IDUtente
sSQL = sSQL & " AND IDTipoOggetto=120"

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_NOME_MACCHINA = ""
Else
    GET_NOME_MACCHINA = fnNotNull(rs!MachineID)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub AGGIORNA_DATA_DI_LIQUIDAZIONE(IDCaricoMerceTesta As Long, DataLiquidazione As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * "
sSQL = sSQL & "FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDCaricoMerceTesta

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    ''''''AGGIORNA DOCUMENTI DI TRASPORTO''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE ValoriOggettoDettaglio0004 SET "
    sSQL = sSQL & "RV_PODataConferimento=" & fnNormDate(DataLiquidazione) & ", "
    sSQL = sSQL & "RV_POIDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
    If AGGIORNA_PREZZO_MEDIO = 1 Then
        sSQL = sSQL & ", RV_POPrezzoMedioInLiq=" & Abs(fnNotNullN(rs!PrezzoMedio))
    End If
    If AGGIORNA_TIPO_LAVORAZIONE = 1 Then
        If fnNotNullN(rs!IDRV_POTipoLavorazione) > 0 Then
            sSQL = sSQL & ", RV_POIDTipoLavorazione=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione))
        End If
    End If
    sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Cn.Execute sSQL
    
    ''''''AGGIORNA FATTURA ACCOMPAGNATORIA'''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE ValoriOggettoDettaglio0001 SET "
    sSQL = sSQL & "RV_PODataConferimento=" & fnNormDate(DataLiquidazione) & ", "
    sSQL = sSQL & "RV_POIDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
    If AGGIORNA_PREZZO_MEDIO = 1 Then
        sSQL = sSQL & ", RV_POPrezzoMedioInLiq=" & Abs(fnNotNullN(rs!PrezzoMedio))
    End If
    If fnNotNullN(rs!IDRV_POTipoLavorazione) > 0 Then
        sSQL = sSQL & ", RV_POIDTipoLavorazione=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione))
    End If
    sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Cn.Execute sSQL
    
    ''''''AGGIORNA CORRISPETTIVI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE ValoriOggettoDettaglio0034 SET "
    sSQL = sSQL & "RV_PODataConferimento=" & fnNormDate(DataLiquidazione) & ", "
    sSQL = sSQL & "RV_POIDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
    If AGGIORNA_PREZZO_MEDIO = 1 Then
        sSQL = sSQL & ", RV_POPrezzoMedioInLiq=" & Abs(fnNotNullN(rs!PrezzoMedio))
    End If
    If fnNotNullN(rs!IDRV_POTipoLavorazione) > 0 Then
        sSQL = sSQL & ", RV_POIDTipoLavorazione=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione))
    End If
    sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Cn.Execute sSQL
        
    ''''''AGGIORNA NOTA DI CREDITO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE ValoriOggettoDettaglio0016 SET "
    sSQL = sSQL & "RV_PODataConferimento=" & fnNormDate(DataLiquidazione) & ", "
    sSQL = sSQL & "RV_POIDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
    If AGGIORNA_PREZZO_MEDIO = 1 Then
        sSQL = sSQL & ", RV_POPrezzoMedioInLiq=" & Abs(fnNotNullN(rs!PrezzoMedio))
    End If
    If fnNotNullN(rs!IDRV_POTipoLavorazione) > 0 Then
        sSQL = sSQL & ", RV_POIDTipoLavorazione=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione))
    End If
    sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Cn.Execute sSQL
            
    ''''''AGGIORNA NOTA DI DEBITO'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE ValoriOggettoDettaglio0007 SET "
    sSQL = sSQL & "RV_PODataConferimento=" & fnNormDate(DataLiquidazione) & ", "
    sSQL = sSQL & "RV_POIDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
    If AGGIORNA_PREZZO_MEDIO = 1 Then
        sSQL = sSQL & ", RV_POPrezzoMedioInLiq=" & Abs(fnNotNullN(rs!PrezzoMedio))
    End If
    If fnNotNullN(rs!IDRV_POTipoLavorazione) > 0 Then
        sSQL = sSQL & ", RV_POIDTipoLavorazione=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione))
    End If
    sSQL = sSQL & " WHERE RV_POIDConferimentoRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Cn.Execute sSQL
    
    ''''''ASSEGNAZIONE MERCE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE RV_POAssegnazioneMerce SET "
    sSQL = sSQL & "DataConferimento=" & fnNormDate(DataLiquidazione)
    If AGGIORNA_TIPO_LAVORAZIONE = 1 Then
        If fnNotNullN(rs!IDRV_POTipoLavorazione) > 0 Then
            sSQL = sSQL & ", IDTipoLavorazione=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione))
        End If
    End If
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Cn.Execute sSQL

    
    ''''''MOVIMENTI''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sSQL = "UPDATE Movimento SET "
    sSQL = sSQL & "RV_PODataConferimento=" & fnNormDate(DataLiquidazione) & ", "
    sSQL = sSQL & "RV_POIDAnagraficaFatturazione=" & Me.CDSocioFatt.KeyFieldID
    If AGGIORNA_PREZZO_MEDIO = 1 Then
        sSQL = sSQL & ", RV_POPrezzoMedioInLiq=" & Abs(fnNotNullN(rs!PrezzoMedio))
    End If
    If AGGIORNA_TIPO_LAVORAZIONE = 1 Then
        If fnNotNullN(rs!IDRV_POTipoLavorazione) > 0 Then
            sSQL = sSQL & ", RV_POIDTipoLavorazioneConf=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione)) & ", "
            sSQL = sSQL & "RV_POIDTipoLavorazione=" & Abs(fnNotNullN(rs!IDRV_POTipoLavorazione))
        End If
    End If
    sSQL = sSQL & " WHERE RV_POIDCaricoMerceRighe=" & fnNotNullN(rs!IDRV_POCaricoMerceRighe)
    Cn.Execute sSQL
    
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub SCRIVI_CODA(IDOggetto As Long)
Dim rs As ADODB.Recordset
Dim sSQL As String

'''''''''''''''''ELIMINAZIONE DATI UTENTE PER IL TIPO OGGETTO'''''''''''''''''''

sSQL = "DELETE FROM RV_POTMP "
sSQL = sSQL & "WHERE IDUtente=" & m_App.IDUser
'sSQL = sSQL & " AND IDTipoOggetto=" & m_DocType.ID

Cn.Execute sSQL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set rs = New ADODB.Recordset

rs.Open "RV_POTMP", Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDSessione = fnGetNewKey("RV_POTMP", "IDSessione")
    rs!IDUtente = m_App.IDUser
    rs!IDTipoOggetto = m_DocType.ID
    rs!IDOggetto = IDOggetto
    rs!Utente = m_App.User
rs.Update

rs.Close
Set rs = Nothing

End Sub
Private Function GET_NUMERO_DOCUMENTO(NuovoDocumento As Boolean) As Long
On Error GoTo ERR_GET_NUMERO_DOCUMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim X_FRM As Form
Dim OLD_CURSOR As Long
Dim NumeroConferimentoDB As Long
Dim PrefissoConferimentoDB As String

GET_NUMERO_DOCUMENTO = 0

sSQL = "SELECT * FROM RV_POTMP "
sSQL = sSQL & "WHERE IDTipoOggetto=" & m_DocType.ID
sSQL = sSQL & " ORDER BY IDSessione, IDUtente"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    If fnNotNullN(rs!IDUtente) = m_App.IDUser Then
        Me.Caption = "SALVATAGGIO IN CORSO.........."
        
        DoEvents
        
        If NuovoDocumento = True Then
            NumeroConferimentoDB = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "NumeroConferimento")
            PrefissoConferimentoDB = GET_PARAMETRI_SOCIO(Link_Esercizio, Me.CDSocio.KeyFieldID, "PrefissoNumeroConferimento")
            If NumeroConferimentoDB = Me.txtNumeroConferimento.Value Then
                Me.txtNumeroConferimento.Value = NumeroConferimentoDB
            End If
            Me.txtPrefissoConferimento.Text = PrefissoConferimentoDB
            
            AGGIORNA_PROGRESSIVO_CONFERIMENTO Me.CDSocio.KeyFieldID, Me.cboEsercizio.CurrentID, Me.txtNumeroConferimento.Value, Me.txtPrefissoConferimento.Text
            
            m_Document("NumeroDocumentoSocio").Value = Me.txtNumeroConferimento.Value
            m_Document("PrefissoNumeroConferimento").Value = Me.txtPrefissoConferimento.Text
        End If
        If GET_ESISTENZA_NUMERO_DOCUMENTO = True Then
            Me.txtNumeroDocumento.Value = fnGetNumeroDocumento
            m_Document("NumeroDocumento").Value = Me.txtNumeroDocumento.Value
            OLD_CURSOR = Cn.CursorLocation
            Cn.CursorLocation = adUseClient
            m_Document.SaveDocument
            AggiornamentoProgressivoSezionale
            Cn.CursorLocation = OLD_CURSOR
        Else
            fnGetNumeroDocumento
            If NumeroDocumentoDisponibile <= Me.txtNumeroDocumento.Value Then
                AggiornamentoProgressivoSezionale
            End If
        End If
        DoEvents
        
        If APERTURA_FORM_CODA = True Then
            Unload frmCoda
            APERTURA_FORM_CODA = False
        End If
        
        GET_NUMERO_DOCUMENTO = 1
        
        rs.CloseResultset
        Set rs = Nothing
    Else
        rs.CloseResultset
        Set rs = Nothing
    
        If APERTURA_FORM_CODA = False Then
            APERTURA_FORM_CODA = True
            Me.Enabled = False
            frmCoda.Show
        End If
        
        Me.Caption = "ATTENDERE......."
        DoEvents
        'GET_NUMERO_DOCUMENTO NuovoDocumento
        
    End If
End If

Set rs = Nothing

Exit Function

ERR_GET_NUMERO_DOCUMENTO:
    MsgBox Err.Description, vbCritical, "Errore coda"
    GET_NUMERO_DOCUMENTO = -1
    Unload frmCoda
End Function
Private Function GET_ORARIO(StringaData As String) As String
Dim Ora As String
Dim Minuti As String
Dim Secondi As String

If Len(DatePart("h", StringaData)) = 1 Then
    Ora = "0" & DatePart("h", StringaData)
Else
    Ora = DatePart("h", StringaData)
End If
If Len(DatePart("n", StringaData)) = 1 Then
    Minuti = "0" & DatePart("n", StringaData)
Else
    Minuti = DatePart("n", StringaData)
End If
If Len(DatePart("s", StringaData)) = 1 Then
    Secondi = "0" & DatePart("s", StringaData)
Else
    Secondi = DatePart("s", StringaData)
End If

GET_ORARIO = Ora & "." & Minuti & "." & Secondi


End Function
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
Private Function CHECK_ABILITAZIONE_DIAMANTE() As Boolean
Dim I As Integer
Dim swChk As DmtSwChk.SwCheck

Set swChk = New DmtSwChk.SwCheck

Set swChk.Connection = Cn.InternalConnection
swChk.SwComponentName = "MODBASE_"
I = swChk.CheckSwComponent

Select Case I
    Case 0
        CHECK_ABILITAZIONE_DIAMANTE = False
        MsgBox swChk.LastError, vbCritical, "Abilitazione programma"
    Case -1
        CHECK_ABILITAZIONE_DIAMANTE = True
    Case -2
        CHECK_ABILITAZIONE_DIAMANTE = False
        MsgBox swChk.LastError, vbCritical, "Abilitazione programma"
End Select

Set swChk = Nothing

End Function

Public Function GeneraMovimentoImballoGestione(IDTipoOggetto As Long, IDOggetto As Long, IDRigaGestioneImballo As Long, IDArticolo As Long, Articolo As String, Quantita As Double, IDTipoProcesso As Long, IDUnitaDiMisura As Long, IDMagazzino As Long, IDLottoImballo As Long, TracciabilitaImballo As Long, PrezzoUnitario As Double, PrezzoImponibile As Double) As Boolean
Dim QuantitaRimasta As Double
Dim QuantitaUtilizzata As Double
Dim sSQL As String

Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection
'''''ELIMINAZIONE DEL MOVIMENTO'''''''''''''''''''''''''''''''
Mov.IDTipoOggetto = IDTipoOggetto
Mov.IDOggetto = IDOggetto
Mov.Delete
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If IDTipoProcesso = 1 Then

    Cn.BeginTrans
    
    Mov.DataMovimento = Me.txtDataDocumento.Text
    Mov.FattoreDiConversione = Null
    
    Mov.GestioneMatricole = False
    Mov.IDEsercizio = Link_Esercizio
    Mov.IDTipoOggetto = IDTipoOggetto
    Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(20, IDTipoProcesso)
    Mov.IDOggetto = IDOggetto
    Mov.IDUtente = TheApp.IDUser
    Mov.IDMagazzinoEntrata = IDMagazzino
    Mov.IDMagazzinoUscita = IDMagazzino
    Mov.Cessione = 0
    Mov.Field "IDValoriOggettoDettaglio", IDOggetto
    Mov.Field "IDAzienda", TheApp.IDFirm
    Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
    Mov.Field "IDTipoAnagrafica", 3
    Mov.Field "IDArticolo", IDArticolo
    Mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
    Mov.Field "IDcambio", Null
    Mov.Field "DescrizioneArticolo", Articolo
    Mov.Field "QuantitaTotale", Quantita
    Mov.Field "Importo", PrezzoImponibile
    Mov.Field "PrezzoUnitario", PrezzoUnitario
    Mov.Field "DataDocumento", Me.txtDataDocumento.Text
    Mov.Field "Oggetto", Mid("Gestione imballi da " & TheApp.FunctionName & "  numero " & Me.txtNumeroDocumento & " del " & Me.txtDataDocumento, 1, 100)
    Mov.Field "IDTipoMovimento", 1
    Mov.Field "RV_POIDLottoImballo", IDLottoImballo
    Mov.Field "LottoImballo", GET_DESCRIZIONE_LOTTO_IMBALLO(IDLottoImballo)
    
    Mov.Field "TipoRiga", trcNessuno
    
    GeneraMovimentoImballoGestione = Mov.Insert
    
    If GeneraMovimentoImballoGestione = False Then
        Cn.RollbackTrans
    Else
        Cn.CommitTrans
    End If
End If

If IDTipoProcesso = 2 Then
    QuantitaRimasta = Quantita
    
    rsLottoImballo.Filter = "Giacenza>0"
    rsLottoImballo.Sort = "Registra DESC, NumeroProgressivo"
    
    If Not ((rsLottoImballo.EOF) And (rsLottoImballo.BOF)) Then
        rsLottoImballo.MoveFirst
        While Not rsLottoImballo.EOF
            If QuantitaRimasta > 0 Then
                Cn.BeginTrans
                
                Mov.DataMovimento = Me.txtDataDocumento.Text
                Mov.FattoreDiConversione = Null
                
                Mov.GestioneMatricole = False
                Mov.IDEsercizio = Link_Esercizio
                Mov.IDTipoOggetto = IDTipoOggetto
                Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(20, IDTipoProcesso)
                Mov.IDOggetto = IDOggetto
                Mov.IDUtente = TheApp.IDUser
                Mov.IDMagazzinoEntrata = IDMagazzino
                Mov.IDMagazzinoUscita = IDMagazzino
                Mov.Cessione = 0
                Mov.Field "IDValoriOggettoDettaglio", IDOggetto
                Mov.Field "IDAzienda", TheApp.IDFirm
                Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
                Mov.Field "IDTipoAnagrafica", 3
                Mov.Field "IDArticolo", IDArticolo
                Mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
                Mov.Field "IDcambio", Null
                Mov.Field "DescrizioneArticolo", Articolo
                If QuantitaRimasta > 0 Then
                    If rsLottoImballo!Registra = 1 Then
                        If Me.chkConfermaDaUtente.Value = Unchecked Then
                            If (QuantitaRimasta - (fnNotNullN(rsLottoImballo!QuantitaSelezionata) + (fnNotNullN(rsLottoImballo!Giacenza) - fnNotNullN(rsLottoImballo!QuantitaSelezionata)))) <= 0 Then
                                QuantitaUtilizzata = QuantitaRimasta
                            Else
                                QuantitaUtilizzata = fnNotNullN(rsLottoImballo!Giacenza) 'fnNotNullN(rsLottoImballo!QuantitaSelezionata)
                            End If
                        Else
                            If (QuantitaRimasta - fnNotNullN(rsLottoImballo!QuantitaSelezionata)) <= 0 Then
                                QuantitaUtilizzata = QuantitaRimasta
                            Else
                                QuantitaUtilizzata = fnNotNullN(rsLottoImballo!QuantitaSelezionata) 'fnNotNullN(rsLottoImballo!QuantitaSelezionata)
                            End If
                        End If
                    Else
                        If (QuantitaRimasta - fnNotNullN(rsLottoImballo!Giacenza)) <= 0 Then
                            QuantitaUtilizzata = QuantitaRimasta
                        Else
                            QuantitaUtilizzata = fnNotNullN(rsLottoImballo!Giacenza)
                        End If
                    End If

                End If
                Mov.Field "QuantitaTotale", QuantitaUtilizzata
                Mov.Field "Importo", PrezzoImponibile
                Mov.Field "PrezzoUnitario", PrezzoUnitario
                Mov.Field "DataDocumento", Me.txtDataDocumento.Text
                Mov.Field "Oggetto", Mid("Gestione imballi da " & TheApp.FunctionName & "  numero " & Me.txtNumeroDocumento & " del " & Me.txtDataDocumento, 1, 100)
                Mov.Field "IDTipoMovimento", 1
                Mov.Field "RV_POIDLottoImballo", fnNotNullN(rsLottoImballo!IDRV_POLottoImballo)
                Mov.Field "LottoImballo", GET_DESCRIZIONE_LOTTO_IMBALLO(fnNotNullN(rsLottoImballo!IDRV_POLottoImballo))
                
                Mov.Field "TipoRiga", trcNessuno
                
                GeneraMovimentoImballoGestione = Mov.Insert
                
                If GeneraMovimentoImballoGestione = False Then
                    Cn.RollbackTrans
                Else
                    Cn.CommitTrans
                End If
            
            
                QuantitaRimasta = QuantitaRimasta - QuantitaUtilizzata
                
                rsLottoImballo!QuantitaMovimentata = QuantitaUtilizzata
                rsLottoImballo.Update
                
                
                sSQL = "UPDATE RV_POLottoImballo SET "
                sSQL = sSQL & " Giacenza=" & fnNormNumber(rsLottoImballo!Giacenza - QuantitaUtilizzata)
                sSQL = sSQL & " WHERE IDRV_POLottoImballo=" & fnNotNullN(rsLottoImballo!IDRV_POLottoImballo)
                Cn.Execute sSQL
            Else
                If fnNotNullN(rsLottoImballo!QuantitaSelezionata) > 0 Then
                    sSQL = "UPDATE RV_POLottoImballo SET "
                    sSQL = sSQL & " Giacenza=" & fnNormNumber(rsLottoImballo!Giacenza)
                    sSQL = sSQL & " WHERE IDRV_POLottoImballo=" & fnNotNullN(rsLottoImballo!IDRV_POLottoImballo)
                    Cn.Execute sSQL
                End If
            End If
            
            rsLottoImballo.MoveNext
            
        Wend
    End If

    If QuantitaRimasta > 0 Then
        Cn.BeginTrans
        
        Mov.DataMovimento = Me.txtDataDocumento.Text
        Mov.FattoreDiConversione = Null
        
        Mov.GestioneMatricole = False
        Mov.IDEsercizio = Link_Esercizio
        Mov.IDTipoOggetto = IDTipoOggetto
        Mov.IDFunzione = GET_FUNZIONE_MAGAZZINO(20, IDTipoProcesso)
        Mov.IDOggetto = IDOggetto
        Mov.IDUtente = TheApp.IDUser
        Mov.IDMagazzinoEntrata = IDMagazzino
        Mov.IDMagazzinoUscita = IDMagazzino
        Mov.Cessione = 0
        Mov.Field "IDValoriOggettoDettaglio", IDOggetto
        Mov.Field "IDAzienda", TheApp.IDFirm
        Mov.Field "IDAnagrafica", Me.CDSocio.KeyFieldID
        Mov.Field "IDTipoAnagrafica", 3
        Mov.Field "IDArticolo", IDArticolo
        Mov.Field "IDUnitaDiMisura", IDUnitaDiMisura
        Mov.Field "IDcambio", Null
        Mov.Field "DescrizioneArticolo", Articolo
        Mov.Field "QuantitaTotale", QuantitaRimasta
        Mov.Field "Importo", PrezzoImponibile
        Mov.Field "PrezzoUnitario", PrezzoUnitario
        Mov.Field "DataDocumento", Me.txtDataDocumento.Text
        Mov.Field "Oggetto", Mid("Gestione imballi da " & TheApp.FunctionName & "  numero " & Me.txtNumeroDocumento & " del " & Me.txtDataDocumento, 1, 100)
        Mov.Field "IDTipoMovimento", 1
        If TracciabilitaImballo = 1 Then
            Mov.Field "RV_POIDLottoImballo", LINK_LOTTO_IMBALLO_PRED
            Mov.Field "LottoImballo", GET_DESCRIZIONE_LOTTO_IMBALLO(LINK_LOTTO_IMBALLO_PRED)
        Else
            Mov.Field "RV_POIDLottoImballo", 0
            Mov.Field "LottoImballo", ""
        End If
       
        
        Mov.Field "TipoRiga", trcNessuno
        
        GeneraMovimentoImballoGestione = Mov.Insert
        
        If GeneraMovimentoImballoGestione = False Then
            Cn.RollbackTrans
        Else
            Cn.CommitTrans
        End If
    End If
    
End If

Set Mov = Nothing

End Function
Private Function GET_LINK_MAGAZZINO_PER_PROCESSO(IDProcesso As Long, IDDocumento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POProcessiDocumentoCoop.IDRV_POProcessiDocumentoCoop, RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop, "
sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDDocumentoCoop, RV_POProcessiDocumentoCoop.IDTipoProcessoCoop,"
sSQL = sSQL & "RV_POProcessiDocumentoCoop.IDMagazzino, RV_POProcessiDocumentoCoop.IDTipoMagazzino, RV_POProcessiDocumentoCoop.IDFunzione, "
sSQL = sSQL & "RV_POSchemaCoop.IDAzienda , RV_POSchemaCoop.IDUtente "
sSQL = sSQL & "FROM RV_POProcessiDocumentoCoop INNER JOIN "
sSQL = sSQL & "RV_POSchemaCoop ON RV_POProcessiDocumentoCoop.IDRV_POSchemaCoop = RV_POSchemaCoop.IDRV_POSchemaCoop "
sSQL = sSQL & "WHERE RV_POSchemaCoop.IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND RV_POSchemaCoop.IDUtente=0"
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDDocumentoCoop=" & IDDocumento
sSQL = sSQL & " AND RV_POProcessiDocumentoCoop.IDTipoProcessoCoop=" & IDProcesso

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_MAGAZZINO_PER_PROCESSO = 0
Else
    GET_LINK_MAGAZZINO_PER_PROCESSO = fnNotNullN(rs!IDTipoMagazzino)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub MOVIMENTAZIONE_GESTIONE_IMBALLI(IDOggetto As Long, IDProcesso As Long, IDDocumento As Long)
On Error GoTo ERR_MOVIMENTAZIONE_GESTIONE_IMBALLI
Dim IDTipoOggetto As Long
Dim IDMagazzino As Long
Dim ris As Boolean

IDTipoOggetto = fnGetTipoOggetto("RV_POGestImbConf")

IDMagazzino = GET_LINK_MAGAZZINO_PER_PROCESSO(IDProcesso, IDDocumento)

Select Case IDMagazzino
    Case 0
        MsgBox "Nessuna configurazione per questo processo per la gestione imballi da conferimento", vbCritical, "Movimentazione gestione imballi"
        Exit Sub
    Case 1
        IDMagazzino = Me.cboMagazzinoConf.CurrentID
    Case 2
        IDMagazzino = Me.CboMagazzinoVend.CurrentID
    Case Else
        MsgBox "Nessuna configurazione per questo processo per la gestione imballi da conferimento", vbCritical, "Movimentazione gestione imballi"
        Exit Sub
End Select

ris = GeneraMovimentoImballoGestione(IDTipoOggetto, IDOggetto, 0, Me.CDImballoGestione.KeyFieldID, Me.CDImballoGestione.Description, Me.txtQuantitaImballo.Value, IDProcesso, Me.cboUMImbGest.CurrentID, IDMagazzino, Me.txtIDLottoImballoGest.Value, Me.chkTracciaImballoGest.Value, Me.txtImpUniAltriImb.Value, Me.txtImpUniAltriImb.Value * Me.txtQuantitaImballo.Value)

If ris = False Then
    MsgBox "Impossibile creare il movimento", vbCritical, "Movimentazione imballi"
End If

Exit Sub
ERR_MOVIMENTAZIONE_GESTIONE_IMBALLI:
    MsgBox Err.Description, vbCritical, "Movimentazione gestione imballi"
End Sub
Private Function GET_LISTINO_DEFAULT(IDAnagraficaCliente As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsAzienda As DmtOleDbLib.adoResultset
Dim Link_Listino_Imballo As Long

GET_LISTINO_DEFAULT = 0

''''''''''''''''''''LISTINO CLIENTE''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If GET_LISTINO_DEFAULT > 0 Then Exit Function


'''''''''''''''''''''''LISTINO AZIENDA'''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT IDListinoDiBase "
sSQL = sSQL & "FROM ConfigurazioneVendite "
sSQL = sSQL & " WHERE IDAzienda=" & m_App.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LISTINO_DEFAULT = 0
Else
    GET_LISTINO_DEFAULT = fnNotNullN(rs!IDListinoDiBase)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function

Private Function GET_PREZZO_IMBALLO(IDArticolo As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDListino As Long


IDListino = GET_LINK_LISTINO

sSQL = "SELECT PrezzoNettoIVA "
sSQL = sSQL & "FROM ListinoPerArticolo "
sSQL = sSQL & "WHERE IDListino=" & IDListino
sSQL = sSQL & " AND IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREZZO_IMBALLO = 0
Else
    GET_PREZZO_IMBALLO = fnNotNullN(rs!PrezzoNettoIva)
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_LINK_LISTINO() As Long
Dim IDListinoDefault As Long

    IDListinoDefault = GET_LISTINO_DEFAULT(Me.CDSocio.KeyFieldID)
    

    GET_LINK_LISTINO = IDListinoDefault
    
End Function
Private Sub GET_TOTALI_RIGA_ALTRI_ADDEBITI()
    Me.txtImponibileAddebiti.Value = Me.txtQtaAddebiti.Value * Me.txtImpUniAddebiti.Value
    Me.txtImpostaAddebiti.Value = (Me.txtImponibileAddebiti.Value / 100) * Me.txtAliquotaIvaAddebiti.Value
    Me.txtTotaleRigaAddebiti.Value = Me.txtImponibileAddebiti.Value + Me.txtImpostaAddebiti.Value
    
End Sub
Private Function GET_LINK_ANA_FATT(IDAnagraficaSocio) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDAnagraficaFatturazione "
sSQL = sSQL & "FROM RV_PO01_ConfigurazioneSocio "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaSocio


Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_ANA_FATT = IDAnagraficaSocio
Else
    If fnNotNullN(rs!IDAnagraficaFatturazione) = 0 Then
        GET_LINK_ANA_FATT = IDAnagraficaSocio
    Else
        GET_LINK_ANA_FATT = fnNotNullN(rs!IDAnagraficaFatturazione)
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CONTROLLO_LAVORAZIONE(IDRigaConferimento As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POAssegnazioneMerce FROM RV_POAssegnazioneMerce "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_LAVORAZIONE = False
Else
    GET_CONTROLLO_LAVORAZIONE = True
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Sub RICALCOLA_CAMPIONATURA(IDRigaConferimento As Long, QuantitaConferita As Double)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim QuantitaCamp As Double


QuantitaCamp = GET_QUANTITA_CAMPIONATA(IDRigaConferimento)

If QuantitaCamp = 0 Then Exit Sub

sSQL = "SELECT * FROM RV_POCampionaturaRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    
    rs!QuantitaDefinitiva = (fnNotNullN(rs!QuantitaCampionata) / QuantitaCamp) * QuantitaConferita
    rs("ImportoNettoRiga").Value = fnNotNullN(rs("ImportoUnitario").Value) * fnNotNullN(rs("QuantitaDefinitiva").Value)
    rs("ImportoImpostaRiga").Value = (fnNotNullN(rs("ImportoNettoRiga").Value) / 100) * GET_ALIQUOTA_IVA(fnNotNullN(rs!IDIva))
    rs("ImportoLordoRiga").Value = fnNotNullN(rs("ImportoNettoRiga").Value) + fnNotNullN(rs("ImportoImpostaRiga").Value)
    
    rs.Update
rs.MoveNext
Wend


rs.Close
Set rs = Nothing
End Sub
Private Function GET_QUANTITA_CAMPIONATA(IDRigaConferimento As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT QuantitaCampionata FROM RV_POCampionatura "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_QUANTITA_CAMPIONATA = 0
Else
    GET_QUANTITA_CAMPIONATA = fnNotNullN(rs!QuantitaCampionata)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ALIQUOTA_IVA(IDIva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
    
sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & IDIva

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    GET_ALIQUOTA_IVA = fnNotNullN(rs!AliquotaIva)
Else
    GET_ALIQUOTA_IVA = 0
End If

rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_CONTROLLO_QUANTITA_CAMPIONATA(IDRigaConferimento As Long, QuantitaConferita As Double) As Boolean
Dim QuantitaCamp As Double

QuantitaCamp = GET_QUANTITA_CAMPIONATA(IDRigaConferimento)

If QuantitaCamp <= QuantitaConferita Then
    GET_CONTROLLO_QUANTITA_CAMPIONATA = True
Else
    GET_CONTROLLO_QUANTITA_CAMPIONATA = False
End If
End Function
Private Function GET_CERTIFICAZIONE_FAMIGLIA_PRODOTTO(IDSocio As Long, NomeCampo As String, IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDFamiglia As Long

GET_CERTIFICAZIONE_FAMIGLIA_PRODOTTO = ""

''''''RECUPERO DELLA FAMIGLIA DELL'ARTICOLO VENDUTO''''''''''''''''''''''''''''''''''''''''''''''''''''
sSQL = "SELECT RV_PO01_IDFamigliaProdotti "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    IDFamiglia = 0
Else
    IDFamiglia = fnNotNullN(rs!RV_PO01_IDFamigliaProdotti)
End If

rs.CloseResultset
Set rs = Nothing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If IDFamiglia = 0 Then Exit Function

sSQL = "SELECT " & NomeCampo & " "
sSQL = sSQL & "FROM RV_PO01_CertificazioneSocioFamiglia INNER JOIN "
sSQL = sSQL & "RV_PO01_Certificazione ON RV_PO01_CertificazioneSocioFamiglia.IDRV_PO01_Certificazione = RV_PO01_Certificazione.IDRV_PO01_Certificazione "
sSQL = sSQL & "WHERE IDRV_PO01_FamigliaProdotti=" & IDFamiglia
sSQL = sSQL & " AND IDAnagrafica=" & IDSocio
sSQL = sSQL & " AND Predefinito=" & fnNormBoolean(1)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CERTIFICAZIONE_FAMIGLIA_PRODOTTO = ""
Else
    GET_CERTIFICAZIONE_FAMIGLIA_PRODOTTO = fnNotNull(rs.adoColumns(NomeCampo).Value)
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
Private Function GET_LINK_OGGETTO()
Dim IDOggetto As Long
Dim sSQL As String

    IDOggetto = fnGetNewKey("Oggetto", "IDOggetto")
    
    sSQL = "INSERT INTO Oggetto ("
    sSQL = sSQL & "IDOggetto, IDTipoOggetto, IDAzienda, IDAttivitaAzienda, IDSezionale, "
    sSQL = sSQL & "Oggetto, DataEmissione, Numero, DataUltimaVariazione, IDUtenteUltimaVariazione, VirtualDelete, IDFunzione)"
    sSQL = sSQL & " VALUES ("
    sSQL = sSQL & IDOggetto & ", "
    sSQL = sSQL & fnGetTipoOggetto & ", "
    sSQL = sSQL & TheApp.IDFirm & ", "
    sSQL = sSQL & GetAttivitaAzienda(TheApp.IDFirm, TheApp.Branch) & ", "
    sSQL = sSQL & Me.cboSezionale.CurrentID & ", "
    sSQL = sSQL & fnNormString(TheApp.FunctionName) & ", "
    sSQL = sSQL & fnNormDate(Me.txtDataDocumento.Text) & ", "
    sSQL = sSQL & fnNormNumber(Me.txtNumeroDocumento.Value) & ", "
    sSQL = sSQL & fnNormDate(Date) & ", "
    sSQL = sSQL & TheApp.IDUser & ", "
    sSQL = sSQL & 0 & ", "
    sSQL = sSQL & TheApp.FunctionID & ")"
    
    Cn.Execute sSQL
    GET_LINK_OGGETTO = IDOggetto

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
Private Function GET_LINK_OGGETTO_ACQ(IDTipoOggettoAcq As Long, IDConferimento As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim NomeTabella As String

If IDTipoOggettoAcq = 0 Then
    GET_LINK_OGGETTO_ACQ = 0
    Exit Function
Else
    If IDTipoOggettoAcq = 1 Then
        NomeTabella = "ValoriOggettoPerTipo012D"
    End If
    If IDTipoOggettoAcq = 2 Then
        NomeTabella = "ValoriOggettoPerTipo012E"
    End If
    
End If

sSQL = "SELECT IDOggetto FROM " & NomeTabella
sSQL = sSQL & " WHERE RV_POIDCaricoMerceTesta=" & IDConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_OGGETTO_ACQ = 0
Else
    GET_LINK_OGGETTO_ACQ = fnNotNullN(rs!IDOggetto)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_TIPO_LIQUIDAZIONE_PER_CONFERIMENTO(IDFiliale As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoLiqConf,IDRV_POTipoConfLiquidazioneChiuso,IDRV_POTipoConfLiquidazioneNuovo "
sSQL = sSQL & "FROM RV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE IDFiliale=" & IDFiliale
sSQL = sSQL & " AND IDSocio=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    LINK_TIPO_LIQ_CONF = 1
Else
    If fnNotNullN(rs!IDRV_POTipoLiqConf) = 0 Then
        LINK_TIPO_LIQ_CONF = 1
        LINK_STATO_LIQ_CHIUSO = 0
        LINK_STATO_LIQ_NUOVO = 0
    Else
        LINK_TIPO_LIQ_CONF = fnNotNullN(rs!IDRV_POTipoLiqConf)
        LINK_STATO_LIQ_CHIUSO = fnNotNullN(rs!IDRV_POTipoConfLiquidazioneChiuso)
        LINK_STATO_LIQ_NUOVO = fnNotNullN(rs!IDRV_POTipoConfLiquidazioneNuovo)
    End If
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

Private Function GET_LINK_IVA_FORNITORE(IDAnagrafica As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDIva "
sSQL = sSQL & "FROM Fornitore "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagrafica
sSQL = sSQL & " AND IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_IVA_FORNITORE = 0
Else

    GET_LINK_IVA_FORNITORE = fnNotNullN(rs!IDIva)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function

End Function
Private Function GET_LINK_LETTERA_INTENTO_PRED(IDAnagrafica As Long, IDTipoAnagrafica As Long, DataDocumento As String, IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_LETTERA_INTENTO_PRED
Dim sSQL As String
Dim IDLetteraIntento As Integer
Dim cmd As ADODB.Command

    'Caricamento delle lettere d'intento del fornitore/Socio
    sSQL = "EXEC SP_LetteraIntentoDefault "
    sSQL = sSQL & TheApp.IDFirm & ", "
    sSQL = sSQL & 3 & ", "
    sSQL = sSQL & Me.CDSocio.KeyFieldID & ", "
    sSQL = sSQL & fnNormDate(Me.txtDataConferimento.Text)
    
    IDLetteraIntento = 0
    
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SP_LetteraIntentoDefault"
    cmd.ActiveConnection = Cn.InternalConnection
    
    
    cmd.Parameters.Append cmd.CreateParameter("IDAzienda", adInteger, adParamInput, , IDAzienda)
    cmd.Parameters.Append cmd.CreateParameter("IDTipoAnagrafica", adInteger, adParamInput, , IDTipoAnagrafica)
    cmd.Parameters.Append cmd.CreateParameter("IDAnagrafica", adInteger, adParamInput, , IDAnagrafica)
    cmd.Parameters.Append cmd.CreateParameter("DataDocumento", adDate, adParamInput, , DataDocumento)
    cmd.Parameters.Append cmd.CreateParameter("IDLetteraIntento", adInteger, adParamOutput, , IDLetteraIntento)
    cmd.Execute
    
    IDLetteraIntento = cmd.Parameters(4).Value
    
    GET_LINK_LETTERA_INTENTO_PRED = IDLetteraIntento

Exit Function
ERR_GET_LINK_LETTERA_INTENTO_PRED:
    MsgBox Err.Description, vbCritical, "Recupero lettera d'intento"
    
End Function

Private Sub EliminaRigheDaDocumentoImballi(IDTestataConferimento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDTipoOggetto As Long

IDTipoOggetto = fnGetTipoOggetto("RV_POGestImbConf")


Set Mov = New DmtMovim.cMovimentazione

Set Mov.Connection = TheApp.Database.Connection


sSQL = "SELECT * FROM RV_POCaricoMerceImballi "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDTestataConferimento

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF

    Mov.IDTipoOggetto = IDTipoOggetto
    Mov.IDOggetto = fnNotNullN(rs!IDRV_POCaricoMerceImballi)
    Mov.Delete

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

Set Mov = Nothing

End Sub

Private Function ControlloRigheMovimentateCampionatura() As Boolean
'RESTITUISCE TRUE SE NON SONO STATE MOVIMENTATE DALLA CAMPIONATURA LE RIGHE DA ALTRI GESTORI
'ALTRIMENTI RESTITUISCE FALSE

ControlloRigheMovimentateCampionatura = True

Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceRighe FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & m_Document("IDRV_POCaricoMerceTesta").Value

Set rs = Cn.OpenResultset(sSQL, adOpenKeyset)

If rs.EOF = True Then
    ControlloRigheMovimentateCampionatura = True
Else
    While Not rs.EOF
        If IsNull(rs!IDRV_POCaricoMerceRighe) Then
            ControlloRigheMovimentateCampionatura = True
        Else
            If ControlloMovimentazioneCampionatura(fnNotNullN(rs!IDRV_POCaricoMerceRighe)) = True Then
                ControlloRigheMovimentateCampionatura = True
            Else
                ControlloRigheMovimentateCampionatura = False
                rs.MoveLast
            End If
        End If
    rs.MoveNext
    Wend
End If

rs.CloseResultset
Set rs = Nothing




End Function

Private Function ControlloMovimentazioneCampionatura(IDRigaConferimento) As Boolean
'RESTITUSCE TRUE SE NON E STATA SE IL RECORDSET E' VUOTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCampionatura FROM RV_POCampionatura "
sSQL = sSQL & " WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = True Then
    ControlloMovimentazioneCampionatura = True
Else
    ControlloMovimentazioneCampionatura = False
End If

rs.CloseResultset
Set rs = Nothing

End Function
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
Private Function CREA_STRINGA_LOTTO_IMBALLO(DataDocumento As String, NumeroDocumento As String, CodiceSocio As String, NumeroProgressivo As String) As String
On Error GoTo ERR_CREA_STRINGA_LOTTO_IMBALLO
Dim Anno As String
Dim Mese As String
Dim Giorno As String
Dim CodiceSocioLocal As String
Dim NumeroDocumentoLocal As String
Dim I As Long
Dim Zeri As String

Anno = Year(DataDocumento)
Mese = Month(DataDocumento)
Giorno = Day(DataDocumento)

If Len(Mese) = 1 Then
    Mese = "0" & Mese
End If
If Len(Giorno) = 1 Then
    Giorno = "0" & Giorno
End If

Zeri = ""
For I = Len(CodiceSocio) To 4
    Zeri = Zeri & "0"
Next
CodiceSocioLocal = Zeri & CodiceSocio

Zeri = ""
For I = Len(NumeroDocumento) To 5
    Zeri = Zeri & "0"
Next
NumeroDocumentoLocal = Zeri & NumeroDocumento


CREA_STRINGA_LOTTO_IMBALLO = Anno
CREA_STRINGA_LOTTO_IMBALLO = CREA_STRINGA_LOTTO_IMBALLO & Mese
CREA_STRINGA_LOTTO_IMBALLO = CREA_STRINGA_LOTTO_IMBALLO & Giorno
CREA_STRINGA_LOTTO_IMBALLO = CREA_STRINGA_LOTTO_IMBALLO & NumeroDocumentoLocal
CREA_STRINGA_LOTTO_IMBALLO = CREA_STRINGA_LOTTO_IMBALLO & "-" & CodiceSocioLocal
CREA_STRINGA_LOTTO_IMBALLO = CREA_STRINGA_LOTTO_IMBALLO & "-" & NumeroProgressivo

Exit Function
ERR_CREA_STRINGA_LOTTO_IMBALLO:
    MsgBox Err.Description, vbCritical, "CREA_STRINGA_LOTTO_IMBALLO"
    CREA_STRINGA_LOTTO_IMBALLO = ""

End Function
Private Function GET_TRACCIA_IMBALLO(IDArticolo As Long) As Long
On Error GoTo ERR_GET_TRACCIA_IMBALLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT RV_POTracciabilitaImballo "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TRACCIA_IMBALLO = 0
Else
    GET_TRACCIA_IMBALLO = Abs(fnNotNullN(rs!RV_POTracciabilitaImballo))
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_TRACCIA_IMBALLO:
    MsgBox Err.Description, vbCritical, "GET_TRACCIA_IMBALLO"
    GET_TRACCIA_IMBALLO = 0
End Function
Private Function GET_DESCRIZIONE_LOTTO_IMBALLO(IDLottoImballo As Long) As String
On Error GoTo ERR_GET_DESCRIZIONE_LOTTO_IMBALLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT LottoImballo "
sSQL = sSQL & "FROM RV_POLottoImballo "
sSQL = sSQL & "WHERE IDRV_POLottoImballo=" & IDLottoImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_DESCRIZIONE_LOTTO_IMBALLO = ""
Else
    GET_DESCRIZIONE_LOTTO_IMBALLO = fnNotNull(rs!LottoImballo)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_DESCRIZIONE_LOTTO_IMBALLO:
    MsgBox Err.Description, vbCritical, "GET_DESCRIZIONE_LOTTO_IMBALLO"
End Function

Private Function CREA_MODIFICA_LOTTO_IMBALLO(IDLottoImballo As Long, IDRigaConferimento As Long, IDCaricoMerceImballi As Long, IDTestaConferimento, Quantita As Double, IDTipoOggetto As Long, IDAnagrafica As Long, IDArticoloImballo As Long, DataDocumento As String, NumeroDocumento As String, CodiceSocio As String, Optional UtilizzoEsclusivo As Long = 0, Optional RifInterno As String = "") As Long
On Error GoTo ERR_CREA_MODIFICA_LOTTO_IMBALLO
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IDLottoReturn As Long
Dim NumeroProgressivo As Long
Dim NumeroProgressivoString As String
Dim I As Long
Dim QuantitaModificata As Double

IDLottoReturn = 0
NumeroProgressivo = GET_NUMERO_LOTTO_IMBALLO(TheApp.IDFirm)

NumeroProgressivoString = ""
For I = Len(CStr(NumeroProgressivo)) To 6
    NumeroProgressivoString = NumeroProgressivoString & "0"
Next
NumeroProgressivoString = NumeroProgressivoString & CStr(NumeroProgressivo)

sSQL = "SELECT * FROM RV_POLottoImballo "
sSQL = sSQL & "WHERE IDRV_POLottoImballo=" & IDLottoImballo

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    
    rs!IDRV_POLottoImballo = fnGetNewKey("RV_POLottoImballo", "IDRV_POLottoImballo")
    rs!IDAzienda = TheApp.IDFirm
    rs!NumeroProgressivo = NumeroProgressivo
    rs!IDTipoOggetto = IDTipoOggetto
    rs!IDRV_POcaricoMercetesta = IDTestaConferimento
    If IDRigaConferimento > 0 Then
        rs!IDRV_POCaricoMerceRighe = IDRigaConferimento
    End If
    rs!IDArticoloImballo = IDArticoloImballo
    rs!LottoImballo = CREA_STRINGA_LOTTO_IMBALLO(DataDocumento, NumeroDocumento, CodiceSocio, NumeroProgressivoString)
    rs!IDAnagrafica = IDAnagrafica
    rs!QuantitaCaricata = Quantita
    rs!Giacenza = Quantita
    rs!DataCreazione = Date
    rs!UtilizzoEsclusivo = UtilizzoEsclusivo
    If Len(RifInterno) > 0 Then rs!RiferimentoEsterno = RifInterno
Else
    QuantitaModificata = Quantita - fnNotNullN(rs!QuantitaCaricata)
    rs!UtilizzoEsclusivo = UtilizzoEsclusivo
    If Len(RifInterno) > 0 Then rs!RiferimentoEsterno = RifInterno
    If QuantitaModificata <> 0 Then
        rs!QuantitaCaricata = Quantita
        rs!Giacenza = rs!Giacenza + QuantitaModificata
        rs!IDArticoloImballo = IDArticoloImballo
        rs!IDAnagrafica = IDAnagrafica
    End If
End If
    
rs.Update

IDLottoReturn = rs!IDRV_POLottoImballo

rs.Close
Set rs = Nothing

If IDLottoReturn > 0 Then
    If IDRigaConferimento > 0 Then
        sSQL = "UPDATE RV_POCaricoMerceRighe SET "
        sSQL = sSQL & "IDRV_POLottoImballo=" & IDLottoReturn
        sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento
        Cn.Execute sSQL
    End If
    
    'If IDCaricoMerceImballi > 0 Then
    '    sSQL = "UPDATE RV_POCaricoMerceImballi SET "
    '    sSQL = sSQL & "IDRV_POLottoImballo=" & IDLottoReturn
    '    sSQL = sSQL & "WHERE IDRV_POCaricoMerceImballi=" & IDCaricoMerceImballi
    '    Cn.Execute sSQL
    'End If
    
End If

CREA_MODIFICA_LOTTO_IMBALLO = IDLottoReturn

Exit Function
ERR_CREA_MODIFICA_LOTTO_IMBALLO:
    MsgBox Err.Description, vbCritical, "ERR_CREA_MODIFICA_LOTTO_IMBALLO"
    
End Function
Private Function GET_NUMERO_LOTTO_IMBALLO(IDAzienda As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(NumeroProgressivo) AS NumeroProgressivo "
sSQL = sSQL & "FROM RV_POLottoImballo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_NUMERO_LOTTO_IMBALLO = 1
Else
    GET_NUMERO_LOTTO_IMBALLO = fnNotNullN(rs!NumeroProgressivo) + 1
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub CREA_RIGHE_TRACCIABILITA_IMBALLO(IDTestaConferimento As Long, IDTipoOggetto As Long, IDAnagrafica As Long, DataDocumento As String, NumeroDocumento As String, CodiceSocio As String)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POCaricoMerceRighe, IDRV_POCaricoMerceTesta, IDRV_POLottoImballo, TracciaImballo, Colli, IDImballo, TracciaImballo "
sSQL = sSQL & "FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDTestaConferimento

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If ((fnNotNullN(rs!TracciaImballo) = 1)) Then 'And (fnNotNullN(rs!IDRV_POLottoImballo) = 0)) Then
        CREA_MODIFICA_LOTTO_IMBALLO fnNotNullN(rs!IDRV_POLottoImballo), fnNotNullN(rs!IDRV_POCaricoMerceRighe), 0, IDTestaConferimento, _
        fnNotNullN(rs!Colli), IDTipoOggetto, IDAnagrafica, fnNotNullN(rs!IDImballo), DataDocumento, NumeroDocumento, CodiceSocio
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub CREA_RECORDSET_LOTTI_IMBALLI(IDArticoloImballo As Long, IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long)
On Error GoTo ERR_CREA_RECORDSET_LOTTI_IMBALLI
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsVista As ADODB.Recordset
Dim I As Long
Dim QuantitaLottoProcesso As Double

    sSQL = "SELECT * FROM RV_POLottoImballo "
    sSQL = sSQL & "WHERE IDRV_POLottoImballo=0"
    
    Set rsVista = New ADODB.Recordset
    rsVista.Open sSQL, Cn.InternalConnection
    
    If Not (rsLottoImballo Is Nothing) Then
        If rsLottoImballo.State > 0 Then
            rsLottoImballo.Close
        End If
        Set rsLottoImballo = Nothing
    End If
    
    Set rsLottoImballo = New ADODB.Recordset
    rsLottoImballo.CursorLocation = adUseClient
    
    With rsVista
        For I = 0 To rsVista.Fields.Count - 1
            Select Case rsVista.Fields(I).Type
                Case adChar, adVarChar, adVarWChar, adWChar, 201
                    rsLottoImballo.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
                Case adNumeric, adBigInt, adCurrency, adDecimal, adDouble, adInteger, adLongVarBinary, adSingle
                    rsLottoImballo.Fields.Append .Fields(I).Name, adDouble, , adFldIsNullable
                Case adDate, adDBTimeStamp, adDBDate
                    rsLottoImballo.Fields.Append .Fields(I).Name, adDBDate, , adFldIsNullable
                Case adSmallInt, adBoolean
                    rsLottoImballo.Fields.Append .Fields(I).Name, adSmallInt, , adFldIsNullable
                Case Else
                    rsLottoImballo.Fields.Append .Fields(I).Name, .Fields(I).Type, .Fields(I).DefinedSize, adFldIsNullable
            End Select
        Next
        rsLottoImballo.Fields.Append "QuantitaSelezionata", adDouble, , adFldIsNullable
        rsLottoImballo.Fields.Append "Rimanenza", adDouble, , adFldIsNullable
        rsLottoImballo.Fields.Append "Registra", adSmallInt, , adFldIsNullable
        rsLottoImballo.Fields.Append "RegistraDaUtente", adSmallInt, , adFldIsNullable
        
        rsLottoImballo.Fields.Append "QuantitaMovimentata", adDouble, , adFldIsNullable
    End With
    rsVista.Close
    Set rsVista = Nothing
    
    rsLottoImballo.Open , , adOpenKeyset, adLockBatchOptimistic

    sSQL = "SELECT * FROM RV_POLottoImballo "
    sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
    sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo
    
    'sSQL = sSQL & " AND Giacenza>0"
    
    sSQL = sSQL & "ORDER BY DataCreazione, NumeroProgressivo"
    
    Set rs = New ADODB.Recordset
    rs.Open sSQL, Cn.InternalConnection
    
    While Not rs.EOF
        rsLottoImballo.AddNew
            For I = 0 To rs.Fields.Count - 1
                rsLottoImballo.Fields(rs.Fields(I).Name).Value = rs.Fields(I).Value
            Next
            
            QuantitaLottoProcesso = GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(fnNotNullN(rs!IDRV_POLottoImballo), IDTipoOggetto, IDOggetto, IDValoriOggettoDettaglio)
            rsLottoImballo!Giacenza = rsLottoImballo!Giacenza + QuantitaLottoProcesso
            rsLottoImballo!QuantitaSelezionata = QuantitaLottoProcesso
            rsLottoImballo!Rimanenza = rsLottoImballo!Giacenza - rsLottoImballo!QuantitaSelezionata
            rsLottoImballo!QuantitaMovimentata = 0
            rsLottoImballo!RiferimentoEsterno = fnNotNull(rs!RiferimentoEsterno)
            If QuantitaLottoProcesso = 0 Then
                rsLottoImballo!Registra = 0
            Else
                rsLottoImballo!Registra = 1
            End If
            rsLottoImballo!RegistraDaUtente = 0
        rsLottoImballo.Update
    rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
Exit Sub
ERR_CREA_RECORDSET_LOTTI_IMBALLI:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_LOTTI_IMBALLI"
End Sub
Private Function GET_GIACENZA_LOTTO_IMBALLO(IDArticoloImballo As Long) As Double
On Error GoTo ERR_GET_GIACENZA_LOTTO_IMBALLO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT SUM(Giacenza) AS TotaleGiacenza "
sSQL = sSQL & "FROM RV_POLottoImballo "
sSQL = sSQL & " WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & " AND IDArticoloImballo=" & IDArticoloImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_GIACENZA_LOTTO_IMBALLO = 0
Else
    GET_GIACENZA_LOTTO_IMBALLO = fnNotNullN(rs!TotaleGiacenza)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_GIACENZA_LOTTO_IMBALLO:
    MsgBox Err.Description, vbCritical, "GET_GIACENZA_LOTTO_IMBALLO"
End Function
Private Function GET_LINK_LOTTO_IMBALLO_PRED(IDAzienda As Long) As Long
On Error GoTo ERR_GET_LINK_LOTTO_IMBALLO_PRED
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim IDReturn As Long
Dim IDLottoReturn As Long
Dim NumeroProgressivo As Long
Dim NumeroProgressivoString As String
Dim I As Long
Dim QuantitaModificata As Double

NumeroProgressivo = GET_NUMERO_LOTTO_IMBALLO(IDAzienda)

NumeroProgressivoString = ""
For I = Len(CStr(NumeroProgressivo)) To 6
    NumeroProgressivoString = NumeroProgressivoString & "0"
Next
NumeroProgressivoString = NumeroProgressivoString & CStr(NumeroProgressivo)


sSQL = "SELECT * FROM RV_POLottoImballo "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND Sistema=1"

Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew

        IDReturn = fnGetNewKey("RV_POLottoImballo", "IDRV_POLottoImballo")
        rs!IDRV_POLottoImballo = IDReturn
        rs!IDAzienda = IDAzienda
        rs!NumeroProgressivo = NumeroProgressivo
        rs!IDTipoOggetto = 0
        rs!IDRV_POcaricoMercetesta = 0
        rs!IDRV_POCaricoMerceRighe = 0
        rs!IDArticoloImballo = 0
        rs!LottoImballo = CREA_STRINGA_LOTTO_IMBALLO(Date, "", "", NumeroProgressivoString)
        rs!IDAnagrafica = 0
        rs!QuantitaCaricata = 0
        rs!Giacenza = 0
        rs!DataCreazione = Date
        rs!Sistema = 1
        
    rs.Update
Else
    IDReturn = fnNotNullN(rs!IDRV_POLottoImballo)
End If

rs.Close
Set rs = Nothing
GET_LINK_LOTTO_IMBALLO_PRED = IDReturn
Exit Function
ERR_GET_LINK_LOTTO_IMBALLO_PRED:
    MsgBox Err.Description, vbCritical, "GET_LINK_LOTTO_IMBALLO_PRED"
End Function

Private Function GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO(IDLottoImballo As Long, IDTipoOggetto As Long, IDOggetto As Long, IDValoriOggettoDettaglio As Long) As Double
On Error GoTo ERR_GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT  SUM(QuantitaTotale) AS QuantitaTotale "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
If IDLottoImballo > 0 Then
    sSQL = sSQL & " AND RV_POIDLottoImballo=" & IDLottoImballo
End If
Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO = 0
Else
    GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO = fnNotNullN(rs!QuantitaTotale)
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO:
    MsgBox Err.Description, vbCritical, "GET_QUANTITA_LOTTO_UTILIZZATA_PROCESSO"
End Function


Private Sub RIPRISTINO_LOTTI_DA_MOVIMENTO(IDOggetto As Long, IDTipoOggetto As Long, IDValoriOggettoDettaglio As Long)
On Error GoTo ERR_RIPRISTINO_LOTTI_DA_MOVIMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM Movimento "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND IDOggetto=" & IDOggetto
sSQL = sSQL & " AND IDValoriOggettoDettaglio=" & IDValoriOggettoDettaglio
sSQL = sSQL & " AND RV_POIDLottoImballo>0"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    sSQL = "UPDATE RV_POLottoImballo SET "
    sSQL = sSQL & "Giacenza=Giacenza+" & fnNormNumber(rs!QuantitaTotale)
    sSQL = sSQL & "WHERE IDRV_POLottoImballo=" & fnNotNullN(rs!RV_POIDLottoImballo)
    Cn.Execute sSQL
    
rs.MoveNext
Wend


rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_RIPRISTINO_LOTTI_DA_MOVIMENTO:
    MsgBox Err.Description, vbCritical, "RIPRISTINO_LOTTI_DA_MOVIMENTO"
    
End Sub

Private Sub ELIMINAZIONE_LOTTO_IMBALLO(IDLottoImballo As Long)
Dim sSQL As String

sSQL = "DELETE FROM RV_POLottoImballo "
sSQL = sSQL & "WHERE IDRV_POLottoImballo=" & IDLottoImballo
Cn.Execute sSQL

End Sub
Private Function GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO(IDLottoImballo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ris As Boolean
Dim NumeroRecord As Long
GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO = False

If IDLottoImballo = 0 Then Exit Function



sSQL = "SELECT COUNT(IDMovimento) as NumeroRecord "
sSQL = sSQL & "FROM Movimento "
sSQL = sSQL & "WHERE RV_POIDLottoImballo=" & IDLottoImballo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
    
End If

rs.CloseResultset
Set rs = Nothing

If NumeroRecord <= 1 Then
    GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO = False
Else
    GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO = True
End If

End Function
Private Function ControlloRigheMovimentateLottiImballo() As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim ris As Boolean

ris = False

sSQL = "SELECT IDRV_POLottoImballo FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
sSQL = sSQL & " AND IDRV_POLottoImballo>0"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If ris = False Then
        If GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO(fnNotNullN(rs!IDRV_POLottoImballo)) = True Then
            ris = True
        End If
    End If
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

If ris = False Then
    sSQL = "SELECT IDRV_POLottoImballo FROM RV_POCaricoMerceImballi "
    sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    sSQL = sSQL & " AND IDRV_POLottoImballo>0"
    Set rs = Cn.OpenResultset(sSQL)
    
    While Not rs.EOF
        If ris = False Then
            If GET_CONTROLLO_MOVIMENTAZIONE_LOTTO_IMBALLO(fnNotNullN(rs!IDRV_POLottoImballo)) = True Then
                ris = True
            End If
        End If
    rs.MoveNext
    Wend
    rs.CloseResultset
    Set rs = Nothing
End If

ControlloRigheMovimentateLottiImballo = ris


End Function
Private Sub CREA_RECORDSET_LOTTO_IMBALLO_DELETE()
On Error GoTo ERR_CREA_RECORDSET_LOTTO_IMBALLO_DELETE

If Not (rslottoImballoDelete Is Nothing) Then
    If rslottoImballoDelete.State > 0 Then
        rslottoImballoDelete.Close
    End If
    Set rslottoImballoDelete = Nothing
End If

Set rslottoImballoDelete = New ADODB.Recordset
rslottoImballoDelete.CursorLocation = adUseClient

rslottoImballoDelete.Fields.Append "IDRV_POLottoImballo", adInteger, , adFldIsNullable

rslottoImballoDelete.Open , , adOpenKeyset, adLockPessimistic

Exit Sub
ERR_CREA_RECORDSET_LOTTO_IMBALLO_DELETE:
    MsgBox Err.Description, vbCritical, "CREA_RECORDSET_LOTTO_IMBALLO_DELETE"
    
End Sub
Private Sub DELETE_TMP_LOTTO_IMBALLO(IDLottoImballo As Long)
On Error GoTo ERR_DELETE_TMP_LOTTO_IMBALLO
rslottoImballoDelete.Filter = "IDRV_POLottoImballo=" & IDLottoImballo

If rslottoImballoDelete.EOF Then
    rslottoImballoDelete.AddNew
        rslottoImballoDelete!IDRV_POLottoImballo = IDLottoImballo
    rslottoImballoDelete.Update
End If


rslottoImballoDelete.Filter = vbNullString
Exit Sub
ERR_DELETE_TMP_LOTTO_IMBALLO:
    MsgBox Err.Description, vbCritical, "DELETE_TMP_LOTTO_IMBALLO"
    
End Sub
Private Sub DELETE_LOTTO_IMBALLO()
On Error GoTo ERR_DELETE_LOTTO_IMBALLO
If ((rslottoImballoDelete.EOF) And (rslottoImballoDelete.BOF)) Then Exit Sub

rslottoImballoDelete.MoveFirst

While Not rslottoImballoDelete.EOF
    ELIMINAZIONE_LOTTO_IMBALLO fnNotNullN(rslottoImballoDelete!IDRV_POLottoImballo)
rslottoImballoDelete.MoveNext
Wend

CREA_RECORDSET_LOTTO_IMBALLO_DELETE
Exit Sub
ERR_DELETE_LOTTO_IMBALLO:
    MsgBox Err.Description, vbCritical, "DELETE_LOTTO_IMBALLO"
End Sub
Private Sub LOTTI_IMBALLI_PER_ELIMINAZIONE(IDCaricoMerceTesta As Long)
On Error GoTo ERR_LOTTI_IMBALLI_PER_ELIMINAZIONE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

'RIGHE DEL CONFERIMENTO
sSQL = "SELECT IDRV_POLottoImballo FROM RV_POCaricoMerceRighe "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDCaricoMerceTesta
sSQL = sSQL & " AND IDRV_POLottoImballo>0"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    DELETE_TMP_LOTTO_IMBALLO fnNotNullN(rs!IDRV_POLottoImballo)
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

'RIGHE ALTRI IMBALLI
sSQL = "SELECT IDRV_POLottoImballo FROM RV_POCaricoMerceImballi "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDCaricoMerceTesta
sSQL = sSQL & " AND IDRV_POLottoImballo>0"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    DELETE_TMP_LOTTO_IMBALLO fnNotNullN(rs!IDRV_POLottoImballo)
rs.MoveNext
Wend
rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_LOTTI_IMBALLI_PER_ELIMINAZIONE:
    MsgBox Err.Description, vbCritical, "LOTTI_IMBALLI_PER_ELIMINAZIONE"
End Sub

Private Sub ParametroStampaSituazioneImballi()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT CalcolaGestioneImballi FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    CALCOLA_RIEP_IMBALLI_STAMPA = fnNotNullN(rs!CalcolaGestioneImballi)
Else
    CALCOLA_RIEP_IMBALLI_STAMPA = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ADD_PESATURE_CONFERIMENTO(IDConferimento As Long)
Dim sSQL As String
Dim rs As ADODB.Recordset
Dim rsnew As ADODB.Recordset


DELETE_PESATURA_UTENTE


'PREPARAZIONE RECORDSET
sSQL = "SELECT * FROM RV_POTMPPesaturaConf "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

Set rsnew = New ADODB.Recordset
rsnew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

'REPERIMENTO DATI RIGA CONFERIMENTO
sSQL = "SELECT * FROM RV_POCaricoMerceRighePes "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & IDConferimento
Set rs = New ADODB.Recordset
rs.Open sSQL, Cn.InternalConnection

While Not rs.EOF
    rsnew.AddNew
        rsnew!IDRV_POCaricoMerceRighePes = rs!IDRV_POCaricoMerceRighePes
        rsnew!IDUtente = TheApp.IDUser
        rsnew!DataInserimento = rs!DataInserimento
        rsnew!DataUltimaModifica = rs!DataUltimaModifica
        rsnew!Colli = fnNotNullN(rs!Colli)
        rsnew!PesoLordo = fnNotNullN(rs!PesoLordo)
        rsnew!Pezzi = fnNotNullN(rs!Pezzi)
        rsnew!NumeroPedane = fnNotNullN(rs!NumeroPedane)
        rsnew!Eliminato = False
        rsnew!Modificato = False
        rsnew!Confermato = False
        rsnew!IDOrdinamento = rs!IDOrdinamento
        rsnew!Annotazioni = rs!Annotazioni
        rsnew!NumeroDocumentoConsegna = rs!NumeroDocumentoConsegna
        rsnew!DataDocumentoConsegna = rs!DataDocumentoConsegna
        rsnew!IDTipoDocumentoCoop = 2
    rsnew.Update
rs.MoveNext
Wend

rsnew.Close
Set rsnew = Nothing

rs.Close
Set rs = Nothing

End Sub
Private Sub DELETE_PESATURA_UTENTE()

Dim sSQL As String

'ELIMINAZIONE DATI
sSQL = "DELETE FROM RV_POTMPPesaturaConf "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"
Cn.Execute sSQL


End Sub
Private Sub AGGIORNA_PESATURA(IDRigaConferimento As Long, IDConferimento As Long, IDOrdinamento As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsPes As ADODB.Recordset


sSQL = "SELECT * FROM RV_POCaricoMerceRighePes "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceRighe=" & IDRigaConferimento

Set rsPes = New ADODB.Recordset

rsPes.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POTMPPesaturaConf "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDOrdinamento=" & IDOrdinamento
sSQL = sSQL & " AND Modificato=" & fnNormBoolean(1)
sSQL = sSQL & " AND Confermato=" & fnNormBoolean(1)
sSQL = sSQL & " AND Eliminato=" & fnNormBoolean(0)
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"
Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If fnNotNullN(rs!IDRV_POCaricoMerceRighePes) = 0 Then
        rsPes.AddNew
            rsPes!IDRV_POCaricoMerceRighe = IDRigaConferimento
            rsPes!IDRV_POcaricoMercetesta = IDConferimento
            rsPes!IDOrdinamento = IDOrdinamento
            rsPes!IDUtenteInserimento = TheApp.IDUser
            rsPes!DataInserimento = rs!DataInserimento
            rsPes!IDUtenteUltimaModifica = TheApp.IDUser
            rsPes!DataUltimaModifica = rs!DataUltimaModifica
            rsPes!Colli = rs!Colli
            rsPes!PesoLordo = rs!PesoLordo
            rsPes!Pezzi = rs!Pezzi
            rsPes!NumeroPedane = rs!NumeroPedane
            rsPes!Annotazioni = rs!Annotazioni
            rsPes!NumeroDocumentoConsegna = rs!NumeroDocumentoConsegna
            rsPes!DataDocumentoConsegna = rs!DataDocumentoConsegna
        rsPes.Update
    Else
        rsPes.Filter = "IDRV_POCaricoMerceRighePes=" & fnNotNullN(rs!IDRV_POCaricoMerceRighePes)
        
        
        If Not (rsPes.EOF) Then
            If (rs!Eliminato = False) Then
                rsPes!IDUtenteUltimaModifica = TheApp.IDUser
                rsPes!DataUltimaModifica = rs!DataUltimaModifica
                rsPes!Colli = rs!Colli
                rsPes!PesoLordo = rs!PesoLordo
                rsPes!Pezzi = rs!Pezzi
                rsPes!NumeroPedane = rs!NumeroPedane
                rsPes!Annotazioni = rs!Annotazioni
                rsPes!NumeroDocumentoConsegna = rs!NumeroDocumentoConsegna
                rsPes!DataDocumentoConsegna = rs!DataDocumentoConsegna
                rsPes.Update
            End If
        End If
        
        rsPes.Filter = vbNullString
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsPes.Close
Set rs = Nothing

End Sub

Private Sub DELETE_PESATURE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsPes As ADODB.Recordset

sSQL = "SELECT * FROM RV_POCaricoMerceRighePes "
sSQL = sSQL & "WHERE IDRV_POCaricoMerceTesta=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Set rsPes = New ADODB.Recordset

rsPes.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

sSQL = "SELECT * FROM RV_POTMPPesaturaConf "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND Eliminato=" & fnNormBoolean(1)
sSQL = sSQL & " AND Confermato=" & fnNormBoolean(1)
sSQL = sSQL & " AND IDRV_POCaricoMerceRighePes>0"
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF

    rsPes.Filter = "IDRV_POCaricoMerceRighePes=" & fnNotNullN(rs!IDRV_POCaricoMerceRighePes)

    If Not (rsPes.EOF) Then
        rsPes.Delete
    End If
    
    rsPes.Filter = vbNullString
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing

rsPes.Close
Set rs = Nothing

End Sub
Private Sub ParametroAttivaCalcoloPesoLordoConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AttivaCalcoloPesoLordoConferimento FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    ATTIVA_CALCOLO_PESO_LORDO = fnNotNullN(rs!AttivaCalcoloPesoLordoConferimento)
Else
    ATTIVA_CALCOLO_PESO_LORDO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub GET_RIEPILOGO_PESI()
On Error Resume Next

If FORM_PESI_MEDI_SHOW = True Then

    frmPesiMedi.RefreshCalcolo
    
End If

'Me.SetFocus
End Sub
Private Sub ParametroAttivaImportoRiepConf()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AbilitaImportoRiepilogoConferimento FROM RV_POSchemaCoop "
sSQL = sSQL & "WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    VISUALIZZA_IMPORTO_F4 = fnNotNullN(rs!AbilitaImportoRiepilogoConferimento)
Else
    VISUALIZZA_IMPORTO_F4 = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub ParametroPesoArticolo()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDRV_POTipoPesoArticolo FROM RV_POSchemaCoop WHERE ("
sSQL = sSQL & "(IDFiliale=" & m_App.Branch & ") "
sSQL = sSQL & "AND (IDUtente=0))"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    TIPO_PESO_ARTICOLO = fnNotNullN(rs!IDRV_POTipoPesoArticolo)
Else
    TIPO_PESO_ARTICOLO = 0
End If

rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub CONTROLLO_MOVIMENTAZIONE_CONFERIMENTO(IDConferimento As Long)
On Error GoTo ERR_CONTROLLO_MOVIMENTAZIONE_CONFERIMENTO
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim Testo As String

sSQL = "SELECT * FROM RV_POIEControlloMovConf "
sSQL = sSQL & "WHERE IDValoriOggettoDettaglio IS NULL "
sSQL = sSQL & " AND IDRV_POCaricoMerceTesta = " & IDConferimento

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Testo = "ATTENZIONE!!!" & vbCrLf
    Testo = Testo & "Alcuni movimenti del conferimento non sono stati eseguiti." & vbCrLf
    Testo = Testo & "Salvare di nuovo il documento e se il problema persiste contattare l'assistenza"
    
    MsgBox Testo, vbCritical, "CONTROLLO MOVIMENTAZIONE CONFERIMENTO"
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_CONTROLLO_MOVIMENTAZIONE_CONFERIMENTO:
    MsgBox Err.Description, vbCritical, "CONTROLLO_MOVIMENTAZIONE_CONFERIMENTO"
End Sub


Private Sub GET_PARAMATRI_FILIALE()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POSchemaCoop "
sSQL = sSQL & " WHERE IDFiliale=" & m_App.Branch
sSQL = sSQL & " AND IDUtente=0"

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    Link_TipoSocio = fnNotNullN(rs!IDCategoriaAnagrafica)
    Link_TipoImballo = fnNotNullN(rs!IDTipoImballo)
    Link_TipoGrezzo = fnNotNullN(rs!IDTipoGrezzo)
    Link_TipoLavorato = fnNotNullN(rs!IDTipoLavorato)
    Link_TipoScarto = fnNotNullN(rs!IDTipoScarto)
    Link_TipoCaloPeso = fnNotNullN(rs!IDTipoCaloPeso)
    Link_TipoAumentoPeso = fnNotNullN(rs!IDTipoAumentoPeso)
    ATTIVAZIONE_NUOVO_CALCOLO = fnNotNullN(rs!AttivazioneNuovoMetodoCalcolo)
    LINK_TIPO_ARTICOLO_CONFERITO = fnNotNullN(rs!IDRV_POTipoSceltaArticoloLottoCampagna)
    LOTTO_CAMPAGNA_OBBLIGATORIO = fnNotNullN(rs!LottoCampagnaObbligatorio)
    LINK_TIPO_ARROTONDAMENTO = fnNotNullN(rs!IDTipoArrotondamentoConferimento)
    PREZZO_MEDIO_AUT = fnNotNullN(rs!PrezzoMedioDaConf)
    AGGIORNA_PREZZO_MEDIO = fnNotNullN(rs!AggiornaPrezzoMedioDaConf)
    AGGIORNA_TIPO_LAVORAZIONE = fnNotNullN(rs!AggiornaTipoLavDaConf)
    CALCOLA_RIEP_IMBALLI_STAMPA = fnNotNullN(rs!CalcolaGestioneImballi)
    ATTIVA_CALCOLO_PESO_LORDO = fnNotNullN(rs!AttivaCalcoloPesoLordoConferimento)
    VISUALIZZA_IMPORTO_F4 = fnNotNullN(rs!AbilitaImportoRiepilogoConferimento)
    'ABILITA_TAB_CONF = fnNotNullN(rs!AbilitaImportiConferimento)
    IDSOCIO_PRE_CONF = fnNotNullN(rs!IDAnagraficaSocioMovPreConferimento)
    ATTIVA_OBBLIGO_N_DOC_SOCIO = fnNotNullN(rs!ObbligatorioNDocInPesConf)
Else
    Link_TipoSocio = 0
    Link_TipoImballo = 0
    Link_TipoGrezzo = 0
    Link_TipoLavorato = 0
    Link_TipoScarto = 0
    Link_TipoCaloPeso = 0
    Link_TipoAumentoPeso = 0
    ATTIVAZIONE_NUOVO_CALCOLO = 0
    LINK_TIPO_ARTICOLO_CONFERITO = 0
    LOTTO_CAMPAGNA_OBBLIGATORIO = 0
    LINK_TIPO_ARROTONDAMENTO = 0
    PREZZO_MEDIO_AUT = 0
    AGGIORNA_PREZZO_MEDIO = 0
    AGGIORNA_TIPO_LAVORAZIONE = 0
    CALCOLA_RIEP_IMBALLI_STAMPA = 0
    ATTIVA_CALCOLO_PESO_LORDO = 0
    VISUALIZZA_IMPORTO_F4 = 0
    'ABILITA_TAB_CONF = 0
    IDSOCIO_PRE_CONF = 0
    ATTIVA_OBBLIGO_N_DOC_SOCIO = fnNotNullN(rs!ObbligatorioNDocInPesConf)
End If


rs.CloseResultset
Set rs = Nothing


'If ABILITA_TAB_CONF = 1 Then ABILITA_TAB


End Sub
Private Sub ABILITA_TAB()

    cboIVA.TabStop = True
    txtImportoUnitario.TabStop = True
    cboTipoLavorazione.TabStop = True
    cboLiquidato.TabStop = True
    chkPrezzoMedio.TabStop = True
    txtNoteAgg.TabStop = True
    

End Sub
Private Sub txtIDOrdineCliente_Change()
On Error GoTo ERR_txtIDOrdineCliente_Change
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

If Me.txtIDOrdineCliente.Value = 0 Then
    Me.cdCliente.Load 0
    Me.lblNomeCliente.Caption = ""
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
    Me.txtNListaPrelievo.Value = 0
    
    Me.cboDestinazione.WriteOn 0
    
    Me.txtDataPartenza.Value = 0

    
    Exit Sub
End If

sSQL = "SELECT Link_Vet_Vettore, Link_nom_ult_sito, Doc_ordine_chiuso, Doc_data_prevista_evasione, Nom_nome, "
sSQL = sSQL & "Doc_data_presso_nom, Doc_numero_presso_nom, Doc_annotazioni_variazio, RV_POAnnotazioniInterna, "
sSQL = sSQL & "RV_PODescrizioneCorpoDocEv, RV_POIDLuogoPresaMerce, RV_POIDTrasportatoreSuccessivo, RV_POIDUtenteBlocco, "
sSQL = sSQL & "RV_PODataArrivoMerceLuogo, RV_POOraArrivoMerceLuogo, RV_PODataArrivoMerce, RV_POOraArrivoMerce, "
sSQL = sSQL & "Link_Doc_spedizione, Link_Nom_Lettera_intento, Link_Nom_IVA, Doc_data, Doc_numero, Link_nom_Anagrafica, "
sSQL = sSQL & "RV_PODataOrdinePadre, RV_PONumeroOrdinePadre, RV_POIDOrdinePadre, RV_PONumeroListaPrelievo, Link_Doc_sezionale, "
sSQL = sSQL & "RV_POIDTipoOrdine, RV_POTargaAutomezzo, RV_POIstruzioniMittente, Doc_annotazioni_variazio, RV_PODescrizioneCorpoDocEv, RV_POAnnotazioniInterna "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & " WHERE IDOggetto=" & Me.txtIDOrdineCliente.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.lblNomeCliente.Caption = ""
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
    Me.txtNListaPrelievo.Value = 0
    
    Me.cboDestinazione.WriteOn 0

    Me.txtDataPartenza.Value = 0
    
    Me.txtIDOrdinePadre = 0

Else
    Me.cdCliente.Load fnNotNullN(rs!Link_nom_anagrafica)
    Me.lblNomeCliente.Caption = fnNotNull(rs!Nom_nome)
    If Len(fnNotNull(rs!Doc_numero_presso_nom)) > 0 Then
        Me.lblNomeCliente.Caption = Me.lblNomeCliente.Caption & " (" & fnNotNull(rs!Doc_numero_presso_nom) & ")"
    End If
    
    
    Me.txtDataOrdine.Value = fnNotNullN(rs!RV_PODataOrdinePadre) ' fnNotNullN(rs!Doc_Data)
    Me.txtNumeroOrdine.Value = fnNotNullN(rs!RV_PONumeroOrdinePadre) 'fnNotNullN(rs!Doc_Numero)
    Me.txtNListaPrelievo.Value = fnNotNullN(rs!RV_PONumeroListaPrelievo)
    
    Me.cboDestinazione.WriteOn fnNotNullN(rs!Link_Nom_ult_sito)

    Me.txtDataPartenza.Value = fnNotNullN(rs!Doc_data_prevista_evasione)
    Me.txtIDOrdinePadre = fnNotNullN(rs!RV_POIDOrdinePadre)

End If

rs.CloseResultset
Set rs = Nothing

If Not (BrwMain.Visible) Then Change

Exit Sub
ERR_txtIDOrdineCliente_Change:
    MsgBox Err.Description, vbCritical, "txtIDOrdineCliente_Change"
End Sub
Private Sub cmdEliminaRifOrdine_Click()
On Error GoTo ERR_cmdEliminaRifOrdine_Click
Dim Testo As String

Testo = "Sei sicuro di voler eliminare il riferimento all'ordine?"

If MsgBox(Testo, vbQuestion + vbYesNo, "Elimazione riferimento ordine") = vbNo Then Exit Sub

Me.txtIDOrdineCliente.Value = 0

Exit Sub
ERR_cmdEliminaRifOrdine_Click:
    MsgBox Err.Description, vbCritical, "cmdEliminaRifOrdine_Click"
End Sub
Private Function GET_PREFISSO_SOCIO_ESERCIZIO_PREC(IDEsercizio As Long, IDAnagraficaSocio As Long) As String
On Error GoTo ERR_GET_PREFISSO_SOCIO_ESERCIZIO_PREC
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim IDEsercizioPrec As Long

GET_PREFISSO_SOCIO_ESERCIZIO_PREC = ""

sSQL = "SELECT IDEsercizio, IDEsercizioRiferimento "
sSQL = sSQL & "FROM Esercizio "
sSQL = sSQL & "WHERE IDEsercizio=" & IDEsercizio

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    IDEsercizioPrec = 0
Else
    IDEsercizioPrec = fnNotNullN(rs!IDEsercizioRiferimento)
End If

rs.CloseResultset
Set rs = Nothing

If IDEsercizioPrec = 0 Then Exit Function

sSQL = "SELECT IDRV_PONumerazionePerSocio, Prefisso "
sSQL = sSQL & "FROM RV_PONumerazionePerSocio "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaSocio
sSQL = sSQL & " AND IDEsercizio=" & IDEsercizioPrec

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PREFISSO_SOCIO_ESERCIZIO_PREC = ""
Else
    GET_PREFISSO_SOCIO_ESERCIZIO_PREC = Trim(fnNotNull(rs!Prefisso))
End If

rs.CloseResultset
Set rs = Nothing

Exit Function
ERR_GET_PREFISSO_SOCIO_ESERCIZIO_PREC:
    GET_PREFISSO_SOCIO_ESERCIZIO_PREC = ""
End Function

Private Sub txtTaraUnitaria_LostFocus()
    CalcoloPesoNetto 1
End Sub

Private Sub txtTargaAutomezzo_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Function GET_CODICE_VARIETA_ART(IDArticolo As Long, NomeCampoVarieta As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CODICE_VARIETA_ART = ""

sSQL = "SELECT Articolo.IDArticolo, RV_PO01_Varieta." & NomeCampoVarieta
sSQL = sSQL & " FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & " RV_PO01_Varieta ON Articolo.RV_PO01_IDVarieta = RV_PO01_Varieta.IDRV_PO01_Varieta"
sSQL = sSQL & " WHERE Articolo.IDArticolo=" & IDArticolo
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CODICE_VARIETA_ART = fnNotNull(rs.adoColumns(NomeCampoVarieta))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CODICE_FAMIGLIA_ART(IDArticolo As Long, NomeCampoFamiglia As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CODICE_FAMIGLIA_ART = ""

sSQL = "SELECT Articolo.IDArticolo, RV_PO01_FamigliaProdotti." & NomeCampoFamiglia
sSQL = sSQL & " FROM Articolo LEFT OUTER JOIN "
sSQL = sSQL & " RV_PO01_FamigliaProdotti ON Articolo.RV_PO01_IDFamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti"
sSQL = sSQL & " WHERE Articolo.IDArticolo=" & IDArticolo
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CODICE_FAMIGLIA_ART = fnNotNull(rs.adoColumns(NomeCampoFamiglia))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_CODICE_LOTTO_CAMPAGNA(IDLottoCampagna As Long, NomeCampo As String) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CODICE_LOTTO_CAMPAGNA = ""

sSQL = "SELECT RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna, RV_PO01_FamigliaProdotti.FamigliaProdotti, RV_PO01_FamigliaProdotti.CodiceImEx, RV_PO01_Varieta.CodiceImEx AS CodiceImExVarieta, "
sSQL = sSQL & "RV_PO01_Varieta.Varieta "
sSQL = sSQL & "FROM RV_PO01_LottoCampagna LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_FamigliaProdotti ON RV_PO01_LottoCampagna.IDRV_PO01_FamigliaProdotti = RV_PO01_FamigliaProdotti.IDRV_PO01_FamigliaProdotti LEFT OUTER JOIN "
sSQL = sSQL & "RV_PO01_Varieta ON RV_PO01_LottoCampagna.IDRV_PO01_Varieta = RV_PO01_Varieta.IDRV_PO01_Varieta "
sSQL = sSQL & "WHERE RV_PO01_LottoCampagna.IDRV_PO01_LottoCampagna=" & IDLottoCampagna
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_CODICE_LOTTO_CAMPAGNA = fnNotNull(rs.adoColumns(NomeCampo))
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_ATTIVA_LOTTO_PROD_ANA_FATT(IDAnagraficaSocio) As Long
On Error GoTo ERR_GET_ATTIVA_LOTTO_PROD_ANA_FATT
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_ATTIVA_LOTTO_PROD_ANA_FATT = 0

sSQL = "SELECT AttivaSelezioneLottoInAnaFatt "
sSQL = sSQL & "FROM RV_PO01_ConfigurazioneSocio "
sSQL = sSQL & "WHERE IDAnagrafica=" & IDAnagraficaSocio

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_ATTIVA_LOTTO_PROD_ANA_FATT = Abs(fnNotNullN(rs!AttivaSelezioneLottoInAnaFatt))
End If

rs.CloseResultset
Set rs = Nothing
Exit Function
ERR_GET_ATTIVA_LOTTO_PROD_ANA_FATT:
    MsgBox Err.Description, vbCritical, "GET_ATTIVA_LOTTO_PROD_ANA_FATT"
End Function
Private Sub SetRigaConferimento(idord As Long)
On Error GoTo ERR_SetRigaConferimento
Screen.MousePointer = 11
If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
    m_DocumentsLink.MoveFirst
    While Not m_DocumentsLink.EOF
        If idord = fnNotNullN(m_DocumentsLink("Link_Ordinamento").Value) Then
            Screen.MousePointer = 0
            Exit Sub
        End If
    m_DocumentsLink.MoveNext
    Wend
End If
Screen.MousePointer = 0
Exit Sub
ERR_SetRigaConferimento:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "SetRigaConferimento"
End Sub
Private Function GET_SOMMA_COLLI_PES() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_SOMMA_COLLI_PES = 0

sSQL = "SELECT SUM(Colli) AS TotaleColli "
sSQL = sSQL & " FROM RV_POTMPPesaturaConf "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"
sSQL = sSQL & " AND IDOrdinamento=" & Link_Ordinamento_riga_conf
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SOMMA_COLLI_PES = fnNotNullN(rs!TotaleColli)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_SOMMA_PESO_PES() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_SOMMA_PESO_PES = 0

sSQL = "SELECT SUM(PesoLordo) AS TotalePesoLordo "
sSQL = sSQL & " FROM RV_POTMPPesaturaConf "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"
sSQL = sSQL & " AND IDOrdinamento=" & Link_Ordinamento_riga_conf
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SOMMA_PESO_PES = fnNotNullN(rs!TotalePesoLordo)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_SOMMA_PEZZI_PES() As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_SOMMA_PEZZI_PES = 0

sSQL = "SELECT SUM(Pezzi) AS TotalePezzi "
sSQL = sSQL & " FROM RV_POTMPPesaturaConf "
sSQL = sSQL & " WHERE IDUtente=" & TheApp.IDUser
sSQL = sSQL & " AND IDTipoDocumentoCoop=2"
sSQL = sSQL & " AND IDOrdinamento=" & Link_Ordinamento_riga_conf
Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    GET_SOMMA_PEZZI_PES = fnNotNullN(rs!TotalePezzi)
End If

rs.CloseResultset
Set rs = Nothing
End Function
