VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{7A1D73E4-F461-11D0-8F01-004033A00AF2}#1.0#0"; "DmtWheel.ocx"
Object = "{5C67DC8E-40E7-11D3-AF44-00105A2FBE61}#3.0#0"; "DmtPrnDlgCtl.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.10#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{9385BB2E-6637-11D1-850D-002018802E11}#3.1#0"; "Dmtsplit.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   11745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20775
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
   ScaleHeight     =   11745
   ScaleWidth      =   20775
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   11400
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   20775
      _LayoutVersion  =   2
      _ExtentX        =   36645
      _ExtentY        =   20108
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   4440
         TabIndex        =   46
         Top             =   0
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   900
      End
      Begin VB.PictureBox PicForm 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   11115
         Left            =   0
         ScaleHeight     =   11085
         ScaleWidth      =   20145
         TabIndex        =   48
         Top             =   0
         Width           =   20175
         Begin VB.PictureBox PicForm2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   10815
            Left            =   120
            ScaleHeight     =   10785
            ScaleWidth      =   19905
            TabIndex        =   49
            Top             =   120
            Width           =   19935
            Begin VB.CommandButton cmdAlyante 
               Caption         =   "ALYANTE"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   15840
               TabIndex        =   323
               Top             =   0
               Width           =   2535
            End
            Begin VB.CommandButton Command3 
               Caption         =   "CAUSALI E-DATI FATTURA"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   13200
               TabIndex        =   322
               Top             =   0
               Width           =   2535
            End
            Begin VB.CommandButton Command1 
               Caption         =   "ALTRE OPERAZIONI"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   10560
               TabIndex        =   288
               Top             =   0
               Width           =   2535
            End
            Begin VB.CommandButton cmdTracciabilita 
               Caption         =   "CONFIGURAZIONE TRACCIABILITA WEB"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   5280
               TabIndex        =   254
               Top             =   0
               Width           =   2535
            End
            Begin VB.CommandButton cmdArtDerivatiDaOrd 
               Caption         =   "CONFIGURAZIONE ARTICOLI DERIVATI DA ORDINE"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   7920
               TabIndex        =   250
               Top             =   20
               Width           =   2535
            End
            Begin VB.Frame FraIntrastat 
               Caption         =   "INTRASTAT"
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
               Left            =   6600
               TabIndex        =   213
               Top             =   1680
               Width           =   13215
               Begin DmtCodDescCtl.DmtCodDesc CDModoTrasporto 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   214
                  Top             =   240
                  Width           =   4335
                  _ExtentX        =   7646
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":479EA
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":47A39
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":47A95
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
               Begin DmtCodDescCtl.DmtCodDesc CDNaturaTransazione 
                  Height          =   615
                  Left            =   4560
                  TabIndex        =   215
                  Top             =   240
                  Width           =   4335
                  _ExtentX        =   7646
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":47AEF
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":47B3E
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":47B9E
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
               Begin DMTDataCmb.DMTCombo cboIntraProvincia 
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   309
                  TabStop         =   0   'False
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
               Begin VB.Label Label7 
                  Caption         =   "Provincia"
                  Height          =   255
                  Index           =   26
                  Left            =   8880
                  TabIndex        =   310
                  Top             =   240
                  Width           =   2055
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "Magazzino di vendita"
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
               Height          =   1815
               Left            =   3360
               TabIndex        =   62
               Top             =   840
               Width           =   3135
               Begin DMTDataCmb.DMTCombo cboCausaleScarico_Vend 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   6
                  Top             =   1440
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin DMTDataCmb.DMTCombo cboCausaleCarico_Vend 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   5
                  Top             =   840
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin DMTDataCmb.DMTCombo cboMagazzinoVendita 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   4
                  Top             =   240
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin VB.Label Label2 
                  Caption         =   "Causale di scarico"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   64
                  Top             =   1200
                  Width           =   2655
               End
               Begin VB.Label Label2 
                  Caption         =   "Causale di carico"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   63
                  Top             =   600
                  Width           =   2655
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Magazzino di conferimento"
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
               Height          =   1815
               Left            =   120
               TabIndex        =   59
               Top             =   840
               Width           =   3135
               Begin DMTDataCmb.DMTCombo cboCausaleScarico_Car 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   3
                  Top             =   1440
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin DMTDataCmb.DMTCombo cboCausaleCarico_Car 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   2
                  Top             =   840
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin DMTDataCmb.DMTCombo cboMagazzinoCarico 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   1
                  Top             =   240
                  Width           =   2895
                  _ExtentX        =   5106
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
                  Caption         =   "Causale di scarico"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   61
                  Top             =   1200
                  Width           =   2895
               End
               Begin VB.Label Label1 
                  Caption         =   "Causale di carico"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   60
                  Top             =   600
                  Width           =   2895
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Altre impostazioni"
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
               Left            =   6600
               TabIndex        =   50
               Top             =   840
               Width           =   13215
               Begin DMTDataCmb.DMTCombo cboTipoProdottoLavorato 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   51
                  Top             =   480
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
               Begin DMTDataCmb.DMTCombo cboTipoProdottoGrezzo 
                  Height          =   315
                  Left            =   2520
                  TabIndex        =   52
                  Top             =   480
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
               Begin DMTDataCmb.DMTCombo cboCategoriaAnagraficaSocio 
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   53
                  Top             =   480
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTDataCmb.DMTCombo cboTipoProdottoImballo 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   54
                  Top             =   480
                  Width           =   2295
                  _ExtentX        =   4048
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
               Begin DMTDataCmb.DMTCombo cboCategoriaAnagraficaCoop 
                  Height          =   315
                  Left            =   9720
                  TabIndex        =   326
                  Top             =   480
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin VB.Label Label7 
                  Caption         =   "Categoria anagrafica cooperativa"
                  Height          =   255
                  Index           =   23
                  Left            =   9720
                  TabIndex        =   327
                  Top             =   240
                  Width           =   3255
               End
               Begin VB.Label Label7 
                  Caption         =   "Tipo prodotto per imballo"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   58
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.Label Label7 
                  Caption         =   "Categoria anagrafica socio"
                  Height          =   255
                  Index           =   1
                  Left            =   7080
                  TabIndex        =   57
                  Top             =   240
                  Width           =   2535
               End
               Begin VB.Label Label7 
                  Caption         =   "Tipo prodotto grezzo"
                  Height          =   255
                  Index           =   4
                  Left            =   2520
                  TabIndex        =   56
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.Label Label7 
                  Caption         =   "Tipo prodotto lavorato"
                  Height          =   255
                  Index           =   5
                  Left            =   4800
                  TabIndex        =   55
                  Top             =   240
                  Width           =   2055
               End
            End
            Begin DmtGridCtl.DmtGrid BrwMain 
               Height          =   735
               Left            =   6120
               TabIndex        =   65
               Top             =   1200
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
            Begin TabDlg.SSTab SSTab1 
               Height          =   7935
               Left            =   120
               TabIndex        =   66
               Top             =   2760
               Width           =   19695
               _ExtentX        =   34740
               _ExtentY        =   13996
               _Version        =   393216
               Tabs            =   10
               Tab             =   6
               TabsPerRow      =   5
               TabHeight       =   794
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Parametri di assegnazione merce"
               TabPicture(0)   =   "frmMain.frx":47BF8
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "Label15(1)"
               Tab(0).Control(1)=   "Label15(0)"
               Tab(0).Control(2)=   "Label4(2)"
               Tab(0).Control(3)=   "Label4(1)"
               Tab(0).Control(4)=   "Label4(0)"
               Tab(0).Control(5)=   "Label14"
               Tab(0).Control(6)=   "Label15(3)"
               Tab(0).Control(7)=   "Label4(3)"
               Tab(0).Control(8)=   "Label4(5)"
               Tab(0).Control(9)=   "Label4(6)"
               Tab(0).Control(10)=   "Label15(4)"
               Tab(0).Control(11)=   "Label4(7)"
               Tab(0).Control(12)=   "Label4(8)"
               Tab(0).Control(13)=   "Label15(5)"
               Tab(0).Control(14)=   "Label15(6)"
               Tab(0).Control(15)=   "Label23(1)"
               Tab(0).Control(16)=   "Label4(9)"
               Tab(0).Control(17)=   "Label4(10)"
               Tab(0).Control(18)=   "Label4(11)"
               Tab(0).Control(19)=   "Label23(2)"
               Tab(0).Control(20)=   "txtNOrdListaIVGammaLav"
               Tab(0).Control(21)=   "txtNListaOrdIVGamma"
               Tab(0).Control(22)=   "txtNListaOrdinePred"
               Tab(0).Control(23)=   "cboSezListaPrelievo"
               Tab(0).Control(24)=   "txtIDOrdineIVGammaLav"
               Tab(0).Control(25)=   "CdClienteIVGammaLav"
               Tab(0).Control(26)=   "txtNOrdIVGammaLav"
               Tab(0).Control(27)=   "txtDataOrdIVGammaLav"
               Tab(0).Control(28)=   "CDClienteOrdIVGamma"
               Tab(0).Control(29)=   "txtNOrdIVGamma"
               Tab(0).Control(30)=   "txtDataOrdIVGamma"
               Tab(0).Control(31)=   "cboUMRigaOrdine"
               Tab(0).Control(32)=   "CDClienteOrdinePred"
               Tab(0).Control(33)=   "txtNumeroOrdinePred"
               Tab(0).Control(34)=   "txtDataOrdinePred"
               Tab(0).Control(35)=   "CDTipoPedana"
               Tab(0).Control(36)=   "cboListinoImballiDefault"
               Tab(0).Control(37)=   "chkStampaPedana"
               Tab(0).Control(38)=   "txtNumeroEtichettePedana"
               Tab(0).Control(39)=   "chkStampaNuovoMetodo"
               Tab(0).Control(40)=   "txtIDOrdineIVGamma"
               Tab(0).Control(41)=   "cmdOrdIVGamma"
               Tab(0).Control(42)=   "cmdOrdIVGammaLav"
               Tab(0).Control(43)=   "cmdEliminaRifOrdIVGamma"
               Tab(0).Control(44)=   "cmdEliminaRifOrdIVGammaLav"
               Tab(0).Control(45)=   "chkVisEtiLavUtente"
               Tab(0).Control(46)=   "chkVisEtiPedUtente"
               Tab(0).Control(47)=   "chkStampaEtiPDF"
               Tab(0).Control(48)=   "txtPercorsoEtiPDF"
               Tab(0).Control(49)=   "Frame8"
               Tab(0).ControlCount=   50
               TabCaption(1)   =   "Processi per documenti"
               TabPicture(1)   =   "frmMain.frx":47C14
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "cmdElimina_Processo"
               Tab(1).Control(0).Enabled=   0   'False
               Tab(1).Control(1)=   "cmdSalva_Processo"
               Tab(1).Control(2)=   "cmdNuovo_Processo"
               Tab(1).Control(3)=   "cboTipoMagazzino"
               Tab(1).Control(4)=   "cboTipoProcesso"
               Tab(1).Control(5)=   "cboDocumentoCoop"
               Tab(1).Control(6)=   "GrigliaProcessi"
               Tab(1).Control(7)=   "cboCausaliMagazzino"
               Tab(1).Control(8)=   "Label5(3)"
               Tab(1).Control(9)=   "Label5(1)"
               Tab(1).Control(10)=   "Label5(0)"
               Tab(1).Control(11)=   "Label5(2)"
               Tab(1).ControlCount=   12
               TabCaption(2)   =   "Sezionale per documenti"
               TabPicture(2)   =   "frmMain.frx":47C30
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "chkPredefinito_Sezionale"
               Tab(2).Control(1)=   "cmdNuovo_Sezionale"
               Tab(2).Control(2)=   "cmdSalva_Sezionale"
               Tab(2).Control(3)=   "cmdElimina_Sezionale"
               Tab(2).Control(4)=   "cboDocumenti_PerSezionale"
               Tab(2).Control(5)=   "cboSezionale"
               Tab(2).Control(6)=   "GrigliaSezionale"
               Tab(2).Control(7)=   "Label10(0)"
               Tab(2).Control(8)=   "Label10(1)"
               Tab(2).ControlCount=   9
               TabCaption(3)   =   "Protocollo ICE"
               TabPicture(3)   =   "frmMain.frx":47C4C
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "chkUsaProtICEPeriodo"
               Tab(3).Control(1)=   "cmdProtICEPeriodo"
               Tab(3).Control(1).Enabled=   0   'False
               Tab(3).Control(2)=   "cmdNuovoProtocolloICE"
               Tab(3).Control(3)=   "cmdSalvaProtocolloICE"
               Tab(3).Control(4)=   "cmdEliminaProtocolloICE"
               Tab(3).Control(4).Enabled=   0   'False
               Tab(3).Control(5)=   "chkPredefinito_ICE"
               Tab(3).Control(6)=   "cboEsercizioICE"
               Tab(3).Control(7)=   "txtNumeroProgressivo"
               Tab(3).Control(8)=   "cboProtocolloICE"
               Tab(3).Control(9)=   "GrigliaProgressivoICE"
               Tab(3).Control(10)=   "Label8"
               Tab(3).Control(11)=   "Label9(0)"
               Tab(3).Control(12)=   "Label9(1)"
               Tab(3).ControlCount=   13
               TabCaption(4)   =   "Opzioni documenti "
               TabPicture(4)   =   "frmMain.frx":47C68
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "chkGestioneArticoli"
               Tab(4).Control(1)=   "chkCreazioneAutomatica"
               Tab(4).Control(2)=   "cmdNuovo_OperazionePerDoc"
               Tab(4).Control(3)=   "cmdSalva__OperazionePerDoc"
               Tab(4).Control(4)=   "cmdElimina_OperazionePerDoc"
               Tab(4).Control(4).Enabled=   0   'False
               Tab(4).Control(5)=   "GrigliaOperazione"
               Tab(4).Control(6)=   "cboDocCoopPerOpe"
               Tab(4).Control(7)=   "cboTipoLavorazioneAut"
               Tab(4).Control(8)=   "Label6(0)"
               Tab(4).Control(9)=   "Label6(1)"
               Tab(4).ControlCount=   10
               TabCaption(5)   =   "Tipo prodotti di quadratura"
               TabPicture(5)   =   "frmMain.frx":47C84
               Tab(5).ControlEnabled=   0   'False
               Tab(5).Control(0)=   "cboCausaleAumentoPeso"
               Tab(5).Control(1)=   "cboCausaleCaloPeso"
               Tab(5).Control(2)=   "cboCausaleTipoScarto"
               Tab(5).Control(3)=   "cboTipoAumentoPeso"
               Tab(5).Control(4)=   "cboTipoCaloPeso"
               Tab(5).Control(5)=   "cboTipoScarto"
               Tab(5).Control(6)=   "cboCausaleScaricoTipoScarto"
               Tab(5).Control(7)=   "cboCausaleScaricoCaloPeso"
               Tab(5).Control(8)=   "cboCausaleScaricoAumentoPeso"
               Tab(5).Control(9)=   "Label11(8)"
               Tab(5).Control(10)=   "Label11(7)"
               Tab(5).Control(11)=   "Label11(6)"
               Tab(5).Control(12)=   "Label11(0)"
               Tab(5).Control(13)=   "Label11(1)"
               Tab(5).Control(14)=   "Label11(2)"
               Tab(5).Control(15)=   "Label11(3)"
               Tab(5).Control(16)=   "Label11(4)"
               Tab(5).Control(17)=   "Label11(5)"
               Tab(5).ControlCount=   18
               TabCaption(6)   =   "Altri parametri"
               TabPicture(6)   =   "frmMain.frx":47CA0
               Tab(6).ControlEnabled=   -1  'True
               Tab(6).Control(0)=   "Frame4"
               Tab(6).Control(0).Enabled=   0   'False
               Tab(6).Control(1)=   "Frame6"
               Tab(6).Control(1).Enabled=   0   'False
               Tab(6).Control(2)=   "fraContributi"
               Tab(6).Control(2).Enabled=   0   'False
               Tab(6).Control(3)=   "Frame5"
               Tab(6).Control(3).Enabled=   0   'False
               Tab(6).ControlCount=   4
               TabCaption(7)   =   "Fatturazione dei soci"
               TabPicture(7)   =   "frmMain.frx":47CBC
               Tab(7).ControlEnabled=   0   'False
               Tab(7).Control(0)=   "chkRaggrLiqSocio"
               Tab(7).Control(1)=   "chkNonRipAcconti"
               Tab(7).Control(2)=   "chkNumerazioneFACSCoop"
               Tab(7).Control(3)=   "chkRipDescrRifLiq"
               Tab(7).Control(4)=   "chkLiquidazioneLordoIVA"
               Tab(7).Control(5)=   "txtRigaInizialePerFattura"
               Tab(7).Control(6)=   "txtRigaFinalePerFattura"
               Tab(7).Control(7)=   "txtDescrizioneConto"
               Tab(7).Control(7).Enabled=   0   'False
               Tab(7).Control(8)=   "txtDescrizioneConto_RigaNeg"
               Tab(7).Control(8).Enabled=   0   'False
               Tab(7).Control(9)=   "txtDescrizioneConto_RigaPos"
               Tab(7).Control(9).Enabled=   0   'False
               Tab(7).Control(10)=   "txtCodiceConto_RigaNeg"
               Tab(7).Control(11)=   "txtCodiceConto_RigaPos"
               Tab(7).Control(12)=   "txtCodiceConto"
               Tab(7).Control(13)=   "cboCausContPerFatturazione"
               Tab(7).Control(14)=   "cboValutaPerFatturazione"
               Tab(7).Control(15)=   "cboPagamentoPerFatturazione"
               Tab(7).Control(16)=   "cboIvaPerFatturazione"
               Tab(7).Control(17)=   "cboTipoCorpoFattSocio"
               Tab(7).Control(18)=   "cboFunzioneDDTAcq"
               Tab(7).Control(19)=   "cboFunzioneFAAcq"
               Tab(7).Control(20)=   "Label12(8)"
               Tab(7).Control(21)=   "Label12(7)"
               Tab(7).Control(22)=   "Label12(6)"
               Tab(7).Control(23)=   "Label12(0)"
               Tab(7).Control(24)=   "Label12(1)"
               Tab(7).Control(25)=   "lblPianodeiDeiConti(4)"
               Tab(7).Control(26)=   "Label12(2)"
               Tab(7).Control(27)=   "Label12(3)"
               Tab(7).Control(28)=   "Label12(4)"
               Tab(7).Control(29)=   "Label12(5)"
               Tab(7).Control(30)=   "lblPianodeiDeiConti(0)"
               Tab(7).Control(31)=   "lblPianodeiDeiConti(1)"
               Tab(7).ControlCount=   32
               TabCaption(8)   =   "Interessi"
               TabPicture(8)   =   "frmMain.frx":47CD8
               Tab(8).ControlEnabled=   0   'False
               Tab(8).Control(0)=   "txtNumeroGiorniAnno"
               Tab(8).Control(1)=   "txtTassoInteressi"
               Tab(8).Control(2)=   "txtDataInizioInteressi"
               Tab(8).Control(3)=   "cmdNuovo_Interessi"
               Tab(8).Control(4)=   "cmdSalva_Interessi"
               Tab(8).Control(5)=   "cmdElimina_Interessi"
               Tab(8).Control(5).Enabled=   0   'False
               Tab(8).Control(6)=   "GrigliaInteressi"
               Tab(8).Control(7)=   "txtPercentualeGiorno"
               Tab(8).Control(8)=   "Label13(3)"
               Tab(8).Control(9)=   "Label13(2)"
               Tab(8).Control(10)=   "Label13(1)"
               Tab(8).Control(11)=   "Label13(0)"
               Tab(8).ControlCount=   12
               TabCaption(9)   =   "Parametri per evasione ordini"
               TabPicture(9)   =   "frmMain.frx":47CF4
               Tab(9).ControlEnabled=   0   'False
               Tab(9).Control(0)=   "chkPesoArtPed"
               Tab(9).Control(1)=   "chkVisMasAvvioVeloce"
               Tab(9).Control(2)=   "chkRipPedOrdSel"
               Tab(9).Control(3)=   "Frame7"
               Tab(9).Control(4)=   "cboUtenteOrd"
               Tab(9).Control(5)=   "cmdEliminaUtenteEva"
               Tab(9).Control(6)=   "cmdSalvaUtenteEva"
               Tab(9).Control(7)=   "cmdNuovoUtenteEva"
               Tab(9).Control(8)=   "GrigliaUtenteEva"
               Tab(9).Control(9)=   "FraOrdSmist"
               Tab(9).Control(10)=   "cboTipoSpezzatura"
               Tab(9).Control(11)=   "CDArtPesPedNeg"
               Tab(9).Control(12)=   "CDArtPesPedPos"
               Tab(9).Control(13)=   "cboCatCommTrasporto"
               Tab(9).Control(14)=   "Label23(0)"
               Tab(9).Control(15)=   "Label16(2)"
               Tab(9).Control(16)=   "Label22"
               Tab(9).Control(17)=   "Label16(0)"
               Tab(9).Control(18)=   "Label20"
               Tab(9).Control(19)=   "Label16(1)"
               Tab(9).ControlCount=   20
               Begin VB.CheckBox chkRaggrLiqSocio 
                  Caption         =   "Raggruppa liquidazioni per socio/anagrafica di fatturazione"
                  Height          =   375
                  Left            =   -69600
                  TabIndex        =   321
                  Top             =   4920
                  Width           =   7215
               End
               Begin VB.CheckBox chkNonRipAcconti 
                  Caption         =   "Non riportare il totale degli acconti nella casella ""Acconti"" della fattura"
                  Height          =   375
                  Left            =   -69600
                  TabIndex        =   320
                  Top             =   4560
                  Width           =   7215
               End
               Begin VB.Frame Frame8 
                  Caption         =   "Ordine vivaio"
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
                  Height          =   2175
                  Left            =   -65280
                  TabIndex        =   315
                  Top             =   5040
                  Width           =   6495
                  Begin VB.CheckBox chkGestioneVivaio 
                     Caption         =   "Attiva nell'ordine da cliente la gestione del vivaio"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   316
                     Top             =   360
                     Width           =   6255
                  End
                  Begin DMTDataCmb.DMTCombo cboTipoTrattAggVivaio 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   317
                     Top             =   1560
                     Width           =   5415
                     _ExtentX        =   9551
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
                  Begin DmtCodDescCtl.DmtCodDesc CDArticoloCommConf 
                     Height          =   615
                     Left            =   120
                     TabIndex        =   318
                     Top             =   720
                     Width           =   7335
                     _ExtentX        =   12938
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":47D10
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":47D8D
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":47DD9
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
               End
               Begin VB.TextBox txtPercorsoEtiPDF 
                  Height          =   315
                  Left            =   -65760
                  TabIndex        =   313
                  Top             =   3240
                  Width           =   9975
               End
               Begin VB.CheckBox chkStampaEtiPDF 
                  Caption         =   "Archivia etichetta in PDF"
                  Height          =   255
                  Left            =   -65760
                  TabIndex        =   312
                  Top             =   2640
                  Width           =   4095
               End
               Begin VB.CheckBox chkUsaProtICEPeriodo 
                  Caption         =   "Usa protocollo ICE per periodo"
                  Height          =   615
                  Left            =   -74880
                  TabIndex        =   307
                  Top             =   1080
                  Width           =   3135
               End
               Begin VB.CommandButton cmdProtICEPeriodo 
                  Caption         =   "CONFIGURA PROTOCOLLO ICE PER PERIODO"
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
                  Left            =   -71760
                  TabIndex        =   306
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   5295
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Funzionamento"
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
                  Height          =   6855
                  Left            =   120
                  TabIndex        =   136
                  Top             =   960
                  Width           =   15975
                  Begin VB.CommandButton Command4 
                     Caption         =   "PARAMETRI PER GESTIONE CERTIFICATI"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   3360
                     Style           =   1  'Graphical
                     TabIndex        =   325
                     Top             =   6000
                     Width           =   3135
                  End
                  Begin VB.CheckBox chkLottoProdPerSocio 
                     Caption         =   "Controllo codice lotto di produzione duplicato per socio"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   324
                     Top             =   6240
                     Width           =   9255
                  End
                  Begin VB.CommandButton Command2 
                     Caption         =   "ALTRI PARAMETRI DI FUNZIONAMENTO"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   120
                     Style           =   1  'Graphical
                     TabIndex        =   319
                     Top             =   6000
                     Width           =   3135
                  End
                  Begin VB.CheckBox chkVisOrdNewPed 
                     Caption         =   "Visualizza elenco ordini in lavorazione quando è nuova pedana"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   311
                     Top             =   4920
                     Width           =   6255
                  End
                  Begin VB.CheckBox chkStampaDocNonAtt 
                     Caption         =   "Stampa documento in evasione ordine NON ATTIVO"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   308
                     Top             =   5280
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkRipDiffColliDaLavDaOrd 
                     Caption         =   "Riporta differenza dei colli quando si seleziona una riga ordine"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   305
                     Top             =   3960
                     Width           =   6255
                  End
                  Begin VB.CheckBox chkEscludiImbImpOrd 
                     Caption         =   "Escludi imballo per la ricerca del prezzo dall'ordine"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   304
                     Top             =   3720
                     Width           =   6255
                  End
                  Begin VB.CheckBox chkVisListaMerceOrd 
                     Caption         =   "Visualizza elenco merce ordinata quando si seleziona un ordine"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   303
                     Top             =   3480
                     Width           =   6255
                  End
                  Begin VB.CheckBox chkNuovoRecLav 
                     Caption         =   "In lavorazione imposta in modalità inserimento quando si seleziona il conferimento"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   300
                     Top             =   6000
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkAttivaUMDocVend 
                     Caption         =   "Attiva l'unità di misura nei documenti di vendita"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   299
                     Top             =   5760
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkForzaTabulazioneConf 
                     Caption         =   "Forza tabulazione nelle caselle degli importi del conferimento "
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   298
                     Top             =   5520
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkStampaEvOrdAtt 
                     Caption         =   "Stampa numero copie in evasione ordine sempre attivo"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   297
                     Top             =   5040
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkVisImportoF4 
                     Caption         =   "Visualizza importo di acquisto e vendita nel riepilogo conferimento (F4)"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   287
                     Top             =   4800
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkAttivaCalcoloPesoConf 
                     Caption         =   "Attiva calcolo peso nel conferimento/acquisto merce"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   286
                     Top             =   4560
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkCalcoloTraspPerPeso 
                     Caption         =   "Calcola le commissioni per il trasporto pedane con il peso della merce"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   285
                     Top             =   4320
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkDocVendForzaNuovo 
                     Caption         =   "All'apertura del documento di vendita forza a nuovo "
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   284
                     Top             =   4080
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkAbilitaForLotto 
                     Caption         =   "Abilita ricerca fornitori nel lotto di produzione"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   281
                     Top             =   3840
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkRiporaPMConfVend 
                     Caption         =   "Riporta indicazione prezzo medio nel conferimento nei documenti di vendita"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   280
                     Top             =   3840
                     Visible         =   0   'False
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkStampaRiepilogoImballiConf 
                     Caption         =   "Esegui riepilogo imballi quando si stampa il confermento o acquisto merce"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   279
                     Top             =   3600
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkFocusColliIVGamma 
                     Caption         =   "Forza il focus nei colli quando si utilizzano le lavorazioni di I° Gamma nei processi di IV° Gamma"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   273
                     Top             =   3360
                     Width           =   9135
                  End
                  Begin VB.CheckBox MsgConfNeg 
                     Caption         =   "Lancia un messaggio quando la lavorazione di un conferimento va in negativo"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   272
                     Top             =   3120
                     Width           =   9135
                  End
                  Begin VB.CheckBox chkAggTipoLavDaConf 
                     Caption         =   "Aggiorna tipo lavorazione da conferimento se valorizzato"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   263
                     Top             =   4680
                     Width           =   6255
                  End
                  Begin VB.CheckBox chkPrezzoMedioDaConf 
                     Caption         =   "Prezzo medio automatico nel conferimento"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   262
                     Top             =   4200
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkAggPrezzoMedioDaConf 
                     Caption         =   "Aggiorna prezzo medio da conferimento"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   261
                     Top             =   4440
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkZeriRifDoc 
                     Caption         =   "Non inserire gli zeri davanti al numero documento di vendita"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   253
                     Top             =   2880
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkPrezziArtDaOrdine 
                     Caption         =   "Prezzi articoli da ordine"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   252
                     Top             =   3240
                     Width           =   6255
                  End
                  Begin VB.CheckBox chKVisRiepConfVend 
                     Caption         =   "Visualizza riepilogo conferimento nei documenti di vendita"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   249
                     Top             =   2640
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkVisAndamDaOrd 
                     Caption         =   "Visualizza andamento ordine"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   227
                     Top             =   2160
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkVisAndamLav 
                     Caption         =   "Visualizza andamento ordine da lavorazione"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   226
                     Top             =   2400
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkOrdineAutomatico 
                     Caption         =   "Assegnazione ordine da smistare in lavorazione automatica dal conferimento"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   217
                     Top             =   1920
                     Width           =   9255
                  End
                  Begin VB.CheckBox chkPedanaAutomatica 
                     Caption         =   "Assegnazione numero pedana in lavorazione automatica da conferimento"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   216
                     Top             =   1680
                     Width           =   9255
                  End
                  Begin VB.CheckBox chkVisOrdSmist 
                     Caption         =   "Visualizza totali della merce da smistare"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   212
                     Top             =   1440
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkVisOrdPrep 
                     Caption         =   "Visualizza totali degli ordini da preparare"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   211
                     Top             =   1200
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkIvaARendere 
                     Caption         =   "Aliquota IVA articolo per imballi a rendere"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   210
                     Top             =   480
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkLottoCampagnaObbligatorio 
                     Caption         =   "Lotto di campagna obbligatorio"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   155
                     Top             =   960
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkGestioneConferimento 
                     Caption         =   "Chiusura conferimento dalla vendita"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   152
                     Top             =   720
                     Width           =   6015
                  End
                  Begin VB.CheckBox chkIvaBloccata 
                     Caption         =   "Aliquota IVA dell'imballo uguale all'IVA dell'articolo venduto"
                     Height          =   255
                     Left            =   6600
                     TabIndex        =   151
                     Top             =   240
                     Width           =   9255
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtNumerazioneLotto 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   137
                     Top             =   435
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
                  Begin DMTEDITNUMLib.dmtNumber txtNumerazioneLottoConf 
                     Height          =   315
                     Left            =   1440
                     TabIndex        =   138
                     Top             =   435
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
                  Begin DMTDataCmb.DMTCombo cboTipoArrotondamento 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   141
                     Top             =   960
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
                  Begin DMTDataCmb.DMTCombo cboTipoSceltaArticoloConferito 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   153
                     Top             =   1560
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
                  Begin DMTDataCmb.DMTCombo cboTipoComportamentoLavorazione 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   156
                     Top             =   2160
                     Width           =   3735
                     _ExtentX        =   6588
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
                  Begin DMTDataCmb.DMTCombo cboTipoPesoArticolo 
                     Height          =   315
                     Left            =   3360
                     TabIndex        =   158
                     Top             =   1530
                     Width           =   3015
                     _ExtentX        =   5318
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
                  Begin DMTDataCmb.DMTCombo cboTipoArrotondamentoConf 
                     Height          =   315
                     Left            =   3360
                     TabIndex        =   174
                     Top             =   960
                     Width           =   3015
                     _ExtentX        =   5318
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
                  Begin DMTDataCmb.DMTCombo cboListinoCampionatura 
                     Height          =   315
                     Left            =   3960
                     TabIndex        =   218
                     Top             =   2160
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
                  Begin DMTDataCmb.DMTCombo cboDecimaliConf 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   220
                     Top             =   2760
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
                  Begin DMTDataCmb.DMTCombo cboDecimaliLav 
                     Height          =   315
                     Left            =   2280
                     TabIndex        =   222
                     Top             =   2760
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
                  Begin DMTDataCmb.DMTCombo cboDecimaliVend 
                     Height          =   315
                     Left            =   4440
                     TabIndex        =   224
                     Top             =   2760
                     Width           =   1935
                     _ExtentX        =   3413
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
                  Begin DMTDataCmb.DMTCombo cboPorto 
                     Height          =   315
                     Left            =   2760
                     TabIndex        =   282
                     Top             =   440
                     Width           =   3615
                     _ExtentX        =   6376
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
                  Begin VB.Label Label7 
                     Caption         =   "Porto per cui non calcolare il trasporto"
                     Height          =   255
                     Index           =   24
                     Left            =   2760
                     TabIndex        =   283
                     ToolTipText     =   "Indica la modalità di consegna di un documento di vendita per cui non calcolare il trasporto nelle commissioni"
                     Top             =   240
                     Width           =   3615
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Decimali per vend."
                     Height          =   255
                     Index           =   22
                     Left            =   4440
                     TabIndex        =   225
                     Top             =   2520
                     Width           =   1815
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Decimali per lav."
                     Height          =   255
                     Index           =   21
                     Left            =   2280
                     TabIndex        =   223
                     Top             =   2520
                     Width           =   1815
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Decimali per conf."
                     Height          =   255
                     Index           =   20
                     Left            =   120
                     TabIndex        =   221
                     Top             =   2520
                     Width           =   2055
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Listino di acq. per camp."
                     Height          =   255
                     Index           =   19
                     Left            =   3960
                     TabIndex        =   219
                     Top             =   1920
                     Width           =   2415
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Tipo di arrot. conferimento"
                     Height          =   255
                     Index           =   16
                     Left            =   3360
                     TabIndex        =   175
                     Top             =   750
                     Width           =   3015
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Tipo peso articolo"
                     Height          =   255
                     Index           =   15
                     Left            =   3360
                     TabIndex        =   159
                     Top             =   1320
                     Width           =   2175
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Tipo comportamento lavorazione"
                     Height          =   255
                     Index           =   14
                     Left            =   120
                     TabIndex        =   157
                     Top             =   1920
                     Width           =   4935
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Tipo scelta articolo conferito"
                     Height          =   255
                     Index           =   13
                     Left            =   120
                     TabIndex        =   154
                     Top             =   1320
                     Width           =   2895
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Tipo di arrot. lavorazione"
                     Height          =   255
                     Index           =   10
                     Left            =   120
                     TabIndex        =   142
                     Top             =   750
                     Width           =   2535
                  End
                  Begin VB.Label Label7 
                     Caption         =   "N° lotto conf."
                     Height          =   255
                     Index           =   9
                     Left            =   1440
                     TabIndex        =   140
                     Top             =   240
                     Width           =   1215
                  End
                  Begin VB.Label Label7 
                     Caption         =   "N° lotto vend."
                     Height          =   255
                     Index           =   8
                     Left            =   120
                     TabIndex        =   139
                     Top             =   240
                     Width           =   1575
                  End
               End
               Begin VB.Frame fraContributi 
                  Caption         =   "Contributi UE per la tracc."
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
                  Height          =   2775
                  Left            =   16200
                  TabIndex        =   274
                  Top             =   5040
                  Width           =   3255
                  Begin VB.TextBox txtContributi 
                     Height          =   1965
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   276
                     Top             =   720
                     Width           =   3015
                  End
                  Begin VB.CheckBox chkContributi 
                     Caption         =   "Contributi"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   275
                     Top             =   360
                     Width           =   1335
                  End
               End
               Begin VB.Frame Frame6 
                  Caption         =   "Altri parametri"
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
                  Height          =   4455
                  Left            =   16200
                  TabIndex        =   143
                  Top             =   960
                  Width           =   3255
                  Begin VB.TextBox txtIscrizioneAlboCoop 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   185
                     Top             =   600
                     Width           =   3015
                  End
                  Begin VB.TextBox txtCodiceAssociato 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   145
                     Top             =   1800
                     Width           =   3015
                  End
                  Begin VB.TextBox txtParametroBNDOO 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   144
                     Top             =   1200
                     Width           =   3015
                  End
                  Begin DMTDataCmb.DMTCombo cboLingua 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   149
                     Top             =   2400
                     Width           =   3015
                     _ExtentX        =   5318
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
                     TabIndex        =   208
                     Top             =   3000
                     Width           =   3015
                     _ExtentX        =   5318
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
                  Begin DMTDataCmb.DMTCombo cboLuogoPresaMerce 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   301
                     Top             =   3600
                     Width           =   3015
                     _ExtentX        =   5318
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
                  Begin VB.CheckBox chkNuovoCalcolo 
                     Caption         =   "Attivazione del nuovo calcolo"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   148
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2895
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Luogo presa merce"
                     Height          =   255
                     Index           =   25
                     Left            =   120
                     TabIndex        =   302
                     Top             =   3360
                     Width           =   2775
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Sezionale C.M.R."
                     Height          =   255
                     Index           =   18
                     Left            =   120
                     TabIndex        =   209
                     Top             =   2760
                     Width           =   2775
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Iscrizione albo coop."
                     Height          =   255
                     Index           =   17
                     Left            =   120
                     TabIndex        =   186
                     Top             =   360
                     Width           =   1935
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Lingua predefinita"
                     Height          =   255
                     Index           =   12
                     Left            =   120
                     TabIndex        =   150
                     Top             =   2160
                     Width           =   1935
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Codice associato"
                     Height          =   255
                     Index           =   11
                     Left            =   120
                     TabIndex        =   147
                     Top             =   1560
                     Width           =   2055
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Parametro B.N.D.O.O."
                     Height          =   255
                     Index           =   7
                     Left            =   120
                     TabIndex        =   146
                     Top             =   960
                     Width           =   1935
                  End
               End
               Begin VB.CheckBox chkNumerazioneFACSCoop 
                  Caption         =   "Numerazione della fattura per conto dei soci uguale a quella della cooperativa"
                  Height          =   375
                  Left            =   -69600
                  TabIndex        =   271
                  Top             =   4200
                  Width           =   7215
               End
               Begin VB.CheckBox chkPesoArtPed 
                  Caption         =   "Peso dell'articolo della pedana in fattura"
                  Height          =   255
                  Left            =   -64080
                  TabIndex        =   270
                  Top             =   2400
                  Width           =   5295
               End
               Begin VB.CheckBox chkVisMasAvvioVeloce 
                  Caption         =   "Visualizza le maschere di spezzatura del lotto e ripesatura pedana nell'avvio veloce"
                  Height          =   255
                  Left            =   -64080
                  TabIndex        =   269
                  Top             =   2040
                  Width           =   8535
               End
               Begin VB.CheckBox chkRipPedOrdSel 
                  Caption         =   "Riporta le pedane nell'ordine cliente selezionato"
                  Height          =   375
                  Left            =   -63480
                  TabIndex        =   267
                  Top             =   6360
                  Width           =   4695
               End
               Begin VB.CheckBox chkVisEtiPedUtente 
                  Caption         =   "Visualizza etichette pedana per utente"
                  Height          =   255
                  Left            =   -65760
                  TabIndex        =   248
                  Top             =   2280
                  Width           =   3975
               End
               Begin VB.CheckBox chkVisEtiLavUtente 
                  Caption         =   "Visualizza etichette lavorazione per utente"
                  Height          =   255
                  Left            =   -65760
                  TabIndex        =   247
                  Top             =   1920
                  Width           =   3975
               End
               Begin VB.CommandButton cmdEliminaRifOrdIVGammaLav 
                  Height          =   330
                  Left            =   -74640
                  Picture         =   "frmMain.frx":47E33
                  Style           =   1  'Graphical
                  TabIndex        =   245
                  ToolTipText     =   "Elimina riferimento dell'ordine"
                  Top             =   6450
                  Width           =   375
               End
               Begin VB.CommandButton cmdEliminaRifOrdIVGamma 
                  Height          =   330
                  Left            =   -74640
                  Picture         =   "frmMain.frx":483BD
                  Style           =   1  'Graphical
                  TabIndex        =   244
                  ToolTipText     =   "Elimina riferimento dell'ordine"
                  Top             =   5490
                  Width           =   375
               End
               Begin VB.CommandButton cmdOrdIVGammaLav 
                  Height          =   330
                  Left            =   -74280
                  Picture         =   "frmMain.frx":48947
                  Style           =   1  'Graphical
                  TabIndex        =   243
                  ToolTipText     =   "Seleziona ordine per lavorazioni di IV Gamma"
                  Top             =   6450
                  Width           =   375
               End
               Begin VB.CommandButton cmdOrdIVGamma 
                  Height          =   330
                  Left            =   -74280
                  Picture         =   "frmMain.frx":48ED1
                  Style           =   1  'Graphical
                  TabIndex        =   242
                  ToolTipText     =   "Seleziona ordine per lavorazioni di IV Gamma"
                  Top             =   5490
                  Width           =   375
               End
               Begin DMTEDITNUMLib.dmtNumber txtIDOrdineIVGamma 
                  Height          =   255
                  Left            =   -70320
                  TabIndex        =   240
                  Top             =   5280
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
               Begin VB.Frame Frame7 
                  Caption         =   "Ordine di preparazione"
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
                  Height          =   1215
                  Left            =   -72600
                  TabIndex        =   191
                  Top             =   2640
                  Width           =   8295
                  Begin VB.CheckBox chkBloccoOrdPrep 
                     Caption         =   "Blocca ordine"
                     Height          =   195
                     Left            =   120
                     TabIndex        =   193
                     Top             =   960
                     Width           =   1575
                  End
                  Begin VB.CommandButton cmdTrovaOrdPrep 
                     Height          =   340
                     Left            =   1080
                     Picture         =   "frmMain.frx":4945B
                     Style           =   1  'Graphical
                     TabIndex        =   192
                     Tag             =   "Trova ordine per preparazione"
                     Top             =   480
                     Width           =   495
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdine 
                     Height          =   315
                     Left            =   6840
                     TabIndex        =   199
                     TabStop         =   0   'False
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
                  Begin DMTDATETIMELib.dmtDate txtDataOrdine 
                     Height          =   315
                     Left            =   5520
                     TabIndex        =   200
                     TabStop         =   0   'False
                     Top             =   480
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                  End
                  Begin DmtCodDescCtl.DmtCodDesc CDCliente 
                     Height          =   615
                     Left            =   1680
                     TabIndex        =   203
                     Top             =   240
                     Width           =   3855
                     _ExtentX        =   6800
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":499E5
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":49A33
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":49A85
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
                  Begin DMTEDITNUMLib.dmtNumber txtIDOrdine 
                     Height          =   340
                     Left            =   120
                     TabIndex        =   204
                     Top             =   480
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   600
                     _StockProps     =   253
                     Text            =   "0"
                     ForeColor       =   12582912
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label21 
                     Caption         =   "Identificativo"
                     Height          =   255
                     Index           =   0
                     Left            =   120
                     TabIndex        =   205
                     Top             =   240
                     Width           =   1455
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Data ordine"
                     Height          =   255
                     Index           =   4
                     Left            =   5520
                     TabIndex        =   202
                     Top             =   240
                     Width           =   1335
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Numero ordine"
                     Height          =   255
                     Index           =   4
                     Left            =   6840
                     TabIndex        =   201
                     Top             =   240
                     Width           =   1335
                  End
               End
               Begin DMTDataCmb.DMTCombo cboUtenteOrd 
                  Height          =   315
                  Left            =   -74760
                  TabIndex        =   187
                  Top             =   1920
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
               Begin VB.CommandButton cmdEliminaUtenteEva 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -65520
                  TabIndex        =   181
                  Top             =   5280
                  Width           =   1215
               End
               Begin VB.CommandButton cmdSalvaUtenteEva 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -65520
                  TabIndex        =   180
                  Top             =   4680
                  Width           =   1215
               End
               Begin VB.CommandButton cmdNuovoUtenteEva 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -65520
                  TabIndex        =   179
                  Top             =   4080
                  Width           =   1215
               End
               Begin DmtGridCtl.DmtGrid GrigliaUtenteEva 
                  Height          =   1815
                  Left            =   -74760
                  TabIndex        =   178
                  Top             =   3960
                  Width           =   9135
                  _ExtentX        =   16113
                  _ExtentY        =   3201
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
               Begin VB.CheckBox chkStampaNuovoMetodo 
                  Caption         =   "Stampa con le etichette professional"
                  Height          =   255
                  Left            =   -65760
                  TabIndex        =   176
                  Top             =   1560
                  Width           =   3495
               End
               Begin VB.CheckBox chkRipDescrRifLiq 
                  Caption         =   "Riporta descrizione riferimento liquidazione"
                  Height          =   375
                  Left            =   -69600
                  TabIndex        =   173
                  Top             =   3840
                  Width           =   7215
               End
               Begin VB.CheckBox chkLiquidazioneLordoIVA 
                  Caption         =   "Liquidazione lordo I.V.A."
                  Height          =   375
                  Left            =   -69600
                  TabIndex        =   172
                  Top             =   3480
                  Width           =   7215
               End
               Begin VB.CheckBox chkGestioneArticoli 
                  Caption         =   "Gestione articoli derivati"
                  Height          =   375
                  Left            =   -71160
                  TabIndex        =   32
                  Top             =   1800
                  Width           =   3015
               End
               Begin DMTEDITNUMLib.dmtNumber txtNumeroGiorniAnno 
                  Height          =   285
                  Left            =   -71280
                  TabIndex        =   162
                  Top             =   1560
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   503
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtNumber txtTassoInteressi 
                  Height          =   285
                  Left            =   -72840
                  TabIndex        =   161
                  Top             =   1560
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDATETIMELib.dmtDate txtDataInizioInteressi 
                  Height          =   285
                  Left            =   -74880
                  TabIndex        =   160
                  Top             =   1560
                  Width           =   1935
                  _Version        =   65536
                  _ExtentX        =   3413
                  _ExtentY        =   503
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.CommandButton cmdNuovo_Interessi 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   167
                  Top             =   2520
                  Width           =   1335
               End
               Begin VB.CommandButton cmdSalva_Interessi 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   165
                  Top             =   3480
                  Width           =   1335
               End
               Begin VB.CommandButton cmdElimina_Interessi 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   169
                  TabStop         =   0   'False
                  Top             =   4440
                  Width           =   1335
               End
               Begin VB.Frame Frame4 
                  Caption         =   "Quadratura "
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2535
                  Left            =   240
                  TabIndex        =   129
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   4695
                  Begin DMTEDITNUMLib.dmtNumber txtQtaMinimaPerVendita 
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   130
                     Top             =   1200
                     Width           =   615
                     _Version        =   65536
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtQtaMinimaPerConferimento 
                     Height          =   255
                     Left            =   3480
                     TabIndex        =   131
                     Top             =   840
                     Width           =   615
                     _Version        =   65536
                     _ExtentX        =   1085
                     _ExtentY        =   450
                     _StockProps     =   253
                     Text            =   "0"
                     BackColor       =   16777215
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Appearance      =   1
                     UseSeparator    =   -1  'True
                     DecFinalZeros   =   -1  'True
                     AllowEmpty      =   0   'False
                  End
                  Begin DMTDataCmb.DMTCombo cboTipoQuadratura 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   132
                     Top             =   480
                     Width           =   2535
                     _ExtentX        =   4471
                     _ExtentY        =   582
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Tipo quadratura per causali"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   6
                     Left            =   120
                     TabIndex        =   135
                     Top             =   240
                     Width           =   1935
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Q.tà minima per chiusura lotto di conferimento"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   2
                     Left            =   120
                     TabIndex        =   134
                     Top             =   840
                     Width           =   3375
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Q.tà minima per chiusura lotto di vendita"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Index           =   3
                     Left            =   120
                     TabIndex        =   133
                     Top             =   1200
                     Width           =   2895
                  End
               End
               Begin DMTEDITNUMLib.dmtNumber txtNumeroEtichettePedana 
                  Height          =   255
                  Left            =   -62640
                  TabIndex        =   117
                  Top             =   3720
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Appearance      =   1
                  AllowEmpty      =   0   'False
               End
               Begin VB.CheckBox chkStampaPedana 
                  Caption         =   "Stampa pedana"
                  Height          =   255
                  Left            =   -65760
                  TabIndex        =   116
                  Top             =   3720
                  Width           =   1815
               End
               Begin VB.CheckBox chkCreazioneAutomatica 
                  Caption         =   "Creazione automatica lotto di vendita"
                  Height          =   375
                  Left            =   -74760
                  TabIndex        =   33
                  Top             =   1800
                  Width           =   3615
               End
               Begin VB.CommandButton cmdElimina_Processo 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   17
                  TabStop         =   0   'False
                  Top             =   4440
                  Width           =   1335
               End
               Begin VB.CommandButton cmdSalva_Processo 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   15
                  Top             =   3480
                  Width           =   1335
               End
               Begin VB.CommandButton cmdNuovo_Processo 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   16
                  Top             =   2520
                  Width           =   1335
               End
               Begin VB.CommandButton cmdNuovoProtocolloICE 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -56880
                  TabIndex        =   29
                  Top             =   3360
                  Width           =   1335
               End
               Begin VB.CommandButton cmdSalvaProtocolloICE 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -56880
                  TabIndex        =   28
                  Top             =   4320
                  Width           =   1335
               End
               Begin VB.CommandButton cmdEliminaProtocolloICE 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -56880
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   5280
                  Width           =   1335
               End
               Begin VB.CheckBox chkPredefinito_ICE 
                  Caption         =   "Predefinito"
                  Height          =   315
                  Left            =   -58320
                  TabIndex        =   26
                  Top             =   2160
                  Width           =   1215
               End
               Begin VB.CheckBox chkPredefinito_Sezionale 
                  Caption         =   "Predefinito"
                  Height          =   315
                  Left            =   -63120
                  TabIndex        =   20
                  Top             =   1680
                  Width           =   1215
               End
               Begin VB.CommandButton cmdNuovo_Sezionale 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   22
                  Top             =   2520
                  Width           =   1335
               End
               Begin VB.CommandButton cmdSalva_Sezionale 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   21
                  Top             =   3480
                  Width           =   1335
               End
               Begin VB.CommandButton cmdElimina_Sezionale 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -57000
                  TabIndex        =   23
                  Top             =   4440
                  Width           =   1335
               End
               Begin VB.CommandButton cmdNuovo_OperazionePerDoc 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -64680
                  TabIndex        =   36
                  Top             =   2520
                  Width           =   1335
               End
               Begin VB.CommandButton cmdSalva__OperazionePerDoc 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -64680
                  TabIndex        =   35
                  Top             =   3360
                  Width           =   1335
               End
               Begin VB.CommandButton cmdElimina_OperazionePerDoc 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -64680
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   4200
                  Width           =   1335
               End
               Begin VB.TextBox txtRigaInizialePerFattura 
                  Height          =   1155
                  Left            =   -74880
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   78
                  Top             =   3600
                  Width           =   5055
               End
               Begin VB.TextBox txtRigaFinalePerFattura 
                  Height          =   1215
                  Left            =   -74880
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   77
                  Top             =   5040
                  Width           =   5055
               End
               Begin VB.TextBox txtDescrizioneConto 
                  Height          =   315
                  Left            =   -73080
                  TabIndex        =   76
                  TabStop         =   0   'False
                  Top             =   1770
                  Width           =   3255
               End
               Begin VB.TextBox txtDescrizioneConto_RigaNeg 
                  Height          =   315
                  Left            =   -73080
                  TabIndex        =   71
                  TabStop         =   0   'False
                  Top             =   2970
                  Width           =   3255
               End
               Begin VB.TextBox txtDescrizioneConto_RigaPos 
                  Height          =   315
                  Left            =   -73080
                  TabIndex        =   70
                  TabStop         =   0   'False
                  Top             =   2370
                  Width           =   3255
               End
               Begin VB.TextBox txtCodiceConto_RigaNeg 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   69
                  Top             =   2970
                  Width           =   1815
               End
               Begin VB.TextBox txtCodiceConto_RigaPos 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   68
                  Top             =   2370
                  Width           =   1815
               End
               Begin VB.TextBox txtCodiceConto 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   67
                  Top             =   1770
                  Width           =   1815
               End
               Begin DMTDataCmb.DMTCombo cboListinoImballiDefault 
                  Height          =   315
                  Left            =   -74640
                  TabIndex        =   11
                  Top             =   3240
                  Width           =   4215
                  _ExtentX        =   7435
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
               Begin DmtCodDescCtl.DmtCodDesc CDTipoPedana 
                  Height          =   615
                  Left            =   -74640
                  TabIndex        =   10
                  Top             =   2400
                  Width           =   4215
                  _ExtentX        =   7435
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":49ADF
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":49B2F
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":49B86
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
               Begin DMTDATETIMELib.dmtDate txtDataOrdinePred 
                  Height          =   330
                  Left            =   -69720
                  TabIndex        =   8
                  Top             =   1650
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdinePred 
                  Height          =   330
                  Left            =   -68280
                  TabIndex        =   9
                  Top             =   1650
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTDataCmb.DMTCombo cboCausContPerFatturazione 
                  Height          =   315
                  Left            =   -69600
                  TabIndex        =   72
                  Top             =   1800
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
               Begin DMTDataCmb.DMTCombo cboValutaPerFatturazione 
                  Height          =   315
                  Left            =   -69600
                  TabIndex        =   73
                  Top             =   2400
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
               Begin DMTDataCmb.DMTCombo cboPagamentoPerFatturazione 
                  Height          =   315
                  Left            =   -67320
                  TabIndex        =   74
                  Top             =   1800
                  Width           =   4935
                  _ExtentX        =   8705
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
               Begin DMTDataCmb.DMTCombo cboIvaPerFatturazione 
                  Height          =   315
                  Left            =   -67320
                  TabIndex        =   75
                  Top             =   2400
                  Width           =   4935
                  _ExtentX        =   8705
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
               Begin DMTDataCmb.DMTCombo cboCausaleAumentoPeso 
                  Height          =   315
                  Left            =   -70680
                  TabIndex        =   43
                  Top             =   4680
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin DMTDataCmb.DMTCombo cboCausaleCaloPeso 
                  Height          =   315
                  Left            =   -70680
                  TabIndex        =   42
                  Top             =   3000
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin DMTDataCmb.DMTCombo cboCausaleTipoScarto 
                  Height          =   315
                  Left            =   -70680
                  TabIndex        =   41
                  Top             =   1320
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin DMTDataCmb.DMTCombo cboTipoAumentoPeso 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   40
                  Top             =   4920
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin DMTDataCmb.DMTCombo cboTipoCaloPeso 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   39
                  Top             =   3240
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin DMTDataCmb.DMTCombo cboTipoScarto 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   38
                  Top             =   1560
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin DmtGridCtl.DmtGrid GrigliaOperazione 
                  Height          =   5415
                  Left            =   -74760
                  TabIndex        =   79
                  Top             =   2280
                  Width           =   9975
                  _ExtentX        =   17595
                  _ExtentY        =   9551
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
               Begin DMTDataCmb.DMTCombo cboDocCoopPerOpe 
                  Height          =   315
                  Left            =   -74760
                  TabIndex        =   31
                  Top             =   1320
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin DMTDataCmb.DMTCombo cboTipoMagazzino 
                  Height          =   315
                  Left            =   -66120
                  TabIndex        =   14
                  Top             =   1440
                  Width           =   2775
                  _ExtentX        =   4895
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
               Begin DMTDataCmb.DMTCombo cboEsercizioICE 
                  Height          =   315
                  Left            =   -62520
                  TabIndex        =   27
                  Top             =   2160
                  Width           =   2535
                  _ExtentX        =   4471
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
               Begin DMTDataCmb.DMTCombo cboDocumenti_PerSezionale 
                  Height          =   315
                  Left            =   -68160
                  TabIndex        =   19
                  Top             =   1680
                  Width           =   4935
                  _ExtentX        =   8705
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
                  Left            =   -74880
                  TabIndex        =   18
                  Top             =   1680
                  Width           =   6615
                  _ExtentX        =   11668
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
               Begin DMTEDITNUMLib.dmtNumber txtNumeroProgressivo 
                  Height          =   315
                  Left            =   -59880
                  TabIndex        =   25
                  Top             =   2160
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  MinValue        =   "1"
               End
               Begin DMTDataCmb.DMTCombo cboProtocolloICE 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   24
                  Top             =   2160
                  Width           =   12255
                  _ExtentX        =   21616
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
               Begin DMTDataCmb.DMTCombo cboTipoProcesso 
                  Height          =   315
                  Left            =   -68880
                  TabIndex        =   13
                  Top             =   1440
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
               Begin DMTDataCmb.DMTCombo cboDocumentoCoop 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   12
                  Top             =   1440
                  Width           =   5895
                  _ExtentX        =   10398
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
               Begin DmtGridCtl.DmtGrid GrigliaProcessi 
                  Height          =   5775
                  Left            =   -74880
                  TabIndex        =   80
                  Top             =   1950
                  Width           =   17775
                  _ExtentX        =   31353
                  _ExtentY        =   10186
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
               Begin DmtGridCtl.DmtGrid GrigliaProgressivoICE 
                  Height          =   5055
                  Left            =   -74880
                  TabIndex        =   81
                  Top             =   2640
                  Width           =   17895
                  _ExtentX        =   31565
                  _ExtentY        =   8916
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
               Begin DmtGridCtl.DmtGrid GrigliaSezionale 
                  Height          =   5655
                  Left            =   -74880
                  TabIndex        =   82
                  Top             =   2070
                  Width           =   17775
                  _ExtentX        =   31353
                  _ExtentY        =   9975
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
               Begin DMTDataCmb.DMTCombo cboTipoLavorazioneAut 
                  Height          =   315
                  Left            =   -71280
                  TabIndex        =   34
                  Top             =   1320
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin DmtCodDescCtl.DmtCodDesc CDClienteOrdinePred 
                  Height          =   615
                  Left            =   -74640
                  TabIndex        =   7
                  Top             =   1410
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":49BE0
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":49C2E
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":49C80
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
               Begin DMTDataCmb.DMTCombo cboTipoCorpoFattSocio 
                  Height          =   315
                  Left            =   -69600
                  TabIndex        =   120
                  Top             =   3000
                  Width           =   7215
                  _ExtentX        =   12726
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
               Begin DMTDataCmb.DMTCombo cboCausaliMagazzino 
                  Height          =   315
                  Left            =   -63240
                  TabIndex        =   121
                  Top             =   1440
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
               Begin DMTDataCmb.DMTCombo cboCausaleScaricoTipoScarto 
                  Height          =   315
                  Left            =   -70680
                  TabIndex        =   124
                  Top             =   1920
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin DMTDataCmb.DMTCombo cboCausaleScaricoCaloPeso 
                  Height          =   315
                  Left            =   -70680
                  TabIndex        =   125
                  Top             =   3600
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin DMTDataCmb.DMTCombo cboCausaleScaricoAumentoPeso 
                  Height          =   315
                  Left            =   -70680
                  TabIndex        =   128
                  Top             =   5280
                  Width           =   4335
                  _ExtentX        =   7646
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
               Begin DmtGridCtl.DmtGrid GrigliaInteressi 
                  Height          =   5655
                  Left            =   -74880
                  TabIndex        =   171
                  TabStop         =   0   'False
                  Top             =   2040
                  Width           =   17775
                  _ExtentX        =   31353
                  _ExtentY        =   9975
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
               Begin DMTEDITNUMLib.dmtNumber txtPercentualeGiorno 
                  Height          =   285
                  Left            =   -69960
                  TabIndex        =   163
                  Top             =   1560
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   503
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  DecimalPlaces   =   5
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboUMRigaOrdine 
                  Height          =   315
                  Left            =   -70200
                  TabIndex        =   183
                  Top             =   2640
                  Width           =   4095
                  _ExtentX        =   7223
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
               Begin VB.Frame FraOrdSmist 
                  Caption         =   "Ordine di smistamento"
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
                  Left            =   -72600
                  TabIndex        =   189
                  Top             =   1800
                  Width           =   8295
                  Begin VB.CommandButton cmdTrovaOrdSmist 
                     Height          =   340
                     Left            =   1080
                     Picture         =   "frmMain.frx":49CDA
                     Style           =   1  'Graphical
                     TabIndex        =   190
                     Tag             =   "Trova ordine per smistamento"
                     Top             =   480
                     Width           =   495
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtNumeroOrdineSmist 
                     Height          =   315
                     Left            =   6840
                     TabIndex        =   194
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
                  Begin DmtCodDescCtl.DmtCodDesc CDClienteSmist 
                     Height          =   615
                     Left            =   1680
                     TabIndex        =   195
                     Top             =   240
                     Width           =   3855
                     _ExtentX        =   6800
                     _ExtentY        =   1085
                     PropCodice      =   $"frmMain.frx":4A264
                     BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PropDescrizione =   $"frmMain.frx":4A2B2
                     BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     MenuFunctions   =   $"frmMain.frx":4A304
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
                  Begin DMTDATETIMELib.dmtDate txtDataOrdineSmist 
                     Height          =   315
                     Left            =   5520
                     TabIndex        =   196
                     Top             =   480
                     Width           =   1215
                     _Version        =   65536
                     _ExtentX        =   2143
                     _ExtentY        =   556
                     _StockProps     =   253
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                  End
                  Begin DMTEDITNUMLib.dmtNumber txtIDOrdineSmist 
                     Height          =   345
                     Left            =   120
                     TabIndex        =   206
                     Top             =   480
                     Width           =   975
                     _Version        =   65536
                     _ExtentX        =   1720
                     _ExtentY        =   600
                     _StockProps     =   253
                     Text            =   "0"
                     ForeColor       =   12582912
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     Appearance      =   1
                     AllowEmpty      =   0   'False
                  End
                  Begin VB.Label Label21 
                     Caption         =   "Identificativo"
                     Height          =   255
                     Index           =   1
                     Left            =   120
                     TabIndex        =   207
                     Top             =   240
                     Width           =   1455
                  End
                  Begin VB.Label Label2 
                     Caption         =   "Data ordine"
                     Height          =   255
                     Index           =   0
                     Left            =   5520
                     TabIndex        =   198
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Numero ordine"
                     Height          =   255
                     Index           =   1
                     Left            =   6840
                     TabIndex        =   197
                     Top             =   240
                     Width           =   1335
                  End
               End
               Begin DMTDATETIMELib.dmtDate txtDataOrdIVGamma 
                  Height          =   330
                  Left            =   -68880
                  TabIndex        =   228
                  Top             =   5505
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtNOrdIVGamma 
                  Height          =   330
                  Left            =   -67560
                  TabIndex        =   229
                  Top             =   5505
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DmtCodDescCtl.DmtCodDesc CDClienteOrdIVGamma 
                  Height          =   615
                  Left            =   -73800
                  TabIndex        =   230
                  Top             =   5265
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4A35E
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4A3AC
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4A3FE
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
               Begin DMTDATETIMELib.dmtDate txtDataOrdIVGammaLav 
                  Height          =   330
                  Left            =   -68880
                  TabIndex        =   234
                  Top             =   6450
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtNOrdIVGammaLav 
                  Height          =   330
                  Left            =   -67560
                  TabIndex        =   235
                  Top             =   6450
                  Width           =   975
                  _Version        =   65536
                  _ExtentX        =   1720
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
               End
               Begin DmtCodDescCtl.DmtCodDesc CdClienteIVGammaLav 
                  Height          =   615
                  Left            =   -73800
                  TabIndex        =   236
                  Top             =   6195
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4A458
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4A4A6
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4A4F8
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
               Begin DMTEDITNUMLib.dmtNumber txtIDOrdineIVGammaLav 
                  Height          =   255
                  Left            =   -68520
                  TabIndex        =   241
                  Top             =   6000
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
               Begin DMTDataCmb.DMTCombo cboTipoSpezzatura 
                  Height          =   315
                  Left            =   -64080
                  TabIndex        =   256
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   2295
                  _ExtentX        =   4048
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
               Begin DMTDataCmb.DMTCombo cboFunzioneDDTAcq 
                  Height          =   315
                  Left            =   -69720
                  TabIndex        =   257
                  Top             =   6000
                  Width           =   3615
                  _ExtentX        =   6376
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
               Begin DMTDataCmb.DMTCombo cboFunzioneFAAcq 
                  Height          =   315
                  Left            =   -65880
                  TabIndex        =   259
                  Top             =   6000
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
               Begin DmtCodDescCtl.DmtCodDesc CDArtPesPedNeg 
                  Height          =   615
                  Left            =   -74760
                  TabIndex        =   264
                  Top             =   6240
                  Width           =   5535
                  _ExtentX        =   9763
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4A552
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4A5A1
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4A60C
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
               Begin DmtCodDescCtl.DmtCodDesc CDArtPesPedPos 
                  Height          =   615
                  Left            =   -69120
                  TabIndex        =   266
                  Top             =   6240
                  Width           =   5535
                  _ExtentX        =   9763
                  _ExtentY        =   1085
                  PropCodice      =   $"frmMain.frx":4A666
                  BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PropDescrizione =   $"frmMain.frx":4A6B5
                  BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MenuFunctions   =   $"frmMain.frx":4A720
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
               Begin DMTDataCmb.DMTCombo cboCatCommTrasporto 
                  Height          =   315
                  Left            =   -61680
                  TabIndex        =   278
                  Top             =   1320
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
               Begin DMTDataCmb.DMTCombo cboSezListaPrelievo 
                  Height          =   315
                  Left            =   -70200
                  TabIndex        =   289
                  Top             =   3240
                  Width           =   4095
                  _ExtentX        =   7223
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
               Begin DMTEDITNUMLib.dmtNumber txtNListaOrdinePred 
                  Height          =   330
                  Left            =   -66960
                  TabIndex        =   291
                  Top             =   1650
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtNListaOrdIVGamma 
                  Height          =   330
                  Left            =   -66480
                  TabIndex        =   293
                  Top             =   5490
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin DMTEDITNUMLib.dmtNumber txtNOrdListaIVGammaLav 
                  Height          =   330
                  Left            =   -66480
                  TabIndex        =   295
                  Top             =   6450
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   582
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
               End
               Begin VB.Label Label23 
                  Caption         =   "Percorso di archiviazione (Consigliata una cartella condivisa da tutti i PC che utlizzano la lavorazione)"
                  Height          =   255
                  Index           =   2
                  Left            =   -65760
                  TabIndex        =   314
                  Top             =   3000
                  Width           =   9975
               End
               Begin VB.Label Label4 
                  Caption         =   "N° lista"
                  Height          =   255
                  Index           =   11
                  Left            =   -66480
                  TabIndex        =   296
                  Top             =   6240
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Caption         =   "N° lista"
                  Height          =   255
                  Index           =   10
                  Left            =   -66480
                  TabIndex        =   294
                  Top             =   5280
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Caption         =   "N° lista"
                  Height          =   255
                  Index           =   9
                  Left            =   -66960
                  TabIndex        =   292
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.Label Label23 
                  Caption         =   "Sezionale per lista prelievo ordine"
                  Height          =   255
                  Index           =   1
                  Left            =   -70200
                  TabIndex        =   290
                  Top             =   3000
                  Width           =   4095
               End
               Begin VB.Label Label23 
                  Caption         =   "Categoria commissione per trasporto"
                  Height          =   255
                  Index           =   0
                  Left            =   -61680
                  TabIndex        =   277
                  Top             =   1080
                  Width           =   3495
               End
               Begin VB.Label Label16 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FFFF&
                  Caption         =   "Altri parametri"
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
                  Index           =   2
                  Left            =   -64080
                  TabIndex        =   268
                  Top             =   1680
                  Width           =   8535
               End
               Begin VB.Label Label12 
                  Caption         =   "Funzione per F.A. Acquisto"
                  Height          =   255
                  Index           =   8
                  Left            =   -65880
                  TabIndex        =   260
                  Top             =   5760
                  Width           =   3135
               End
               Begin VB.Label Label12 
                  Caption         =   "Funzione per D.D.T. Acquisto"
                  Height          =   255
                  Index           =   7
                  Left            =   -69720
                  TabIndex        =   258
                  Top             =   5760
                  Width           =   3015
               End
               Begin VB.Label Label22 
                  Caption         =   "Tipo spezzatura lotto"
                  Height          =   255
                  Left            =   -64080
                  TabIndex        =   255
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   2055
               End
               Begin VB.Label Label15 
                  Caption         =   "Etichette Professional"
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
                  Left            =   -65760
                  TabIndex        =   246
                  Top             =   1200
                  Width           =   2535
               End
               Begin VB.Label Label15 
                  Caption         =   "Ordine predefinito per merce lavorata e utilizzata in IV Gamma"
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
                  Left            =   -74640
                  TabIndex        =   239
                  Top             =   6000
                  Width           =   6495
               End
               Begin VB.Label Label4 
                  Caption         =   "Numero"
                  Height          =   255
                  Index           =   8
                  Left            =   -67560
                  TabIndex        =   238
                  Top             =   6240
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Caption         =   "Data ordine"
                  Height          =   255
                  Index           =   7
                  Left            =   -68880
                  TabIndex        =   237
                  Top             =   6240
                  Width           =   1095
               End
               Begin VB.Label Label15 
                  Caption         =   "Ordine predefinito per lavorazioni di IV Gamma"
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
                  Left            =   -74640
                  TabIndex        =   233
                  Top             =   5040
                  Width           =   5175
               End
               Begin VB.Label Label4 
                  Caption         =   "Numero"
                  Height          =   255
                  Index           =   6
                  Left            =   -67560
                  TabIndex        =   232
                  Top             =   5280
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Caption         =   "Data ordine "
                  Height          =   255
                  Index           =   5
                  Left            =   -68880
                  TabIndex        =   231
                  Top             =   5280
                  Width           =   1215
               End
               Begin VB.Label Label4 
                  Caption         =   "Unità di misura ordine predefinita"
                  Height          =   255
                  Index           =   3
                  Left            =   -70200
                  TabIndex        =   184
                  Top             =   2400
                  Width           =   2895
               End
               Begin VB.Label Label15 
                  Caption         =   "Ordini da cliente"
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
                  Left            =   -70200
                  TabIndex        =   182
                  Top             =   2160
                  Width           =   1815
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "% Giornaliera"
                  Height          =   255
                  Index           =   3
                  Left            =   -69960
                  TabIndex        =   170
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "N° giorni anno"
                  Height          =   255
                  Index           =   2
                  Left            =   -71280
                  TabIndex        =   168
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "% Annuale"
                  Height          =   255
                  Index           =   1
                  Left            =   -72600
                  TabIndex        =   166
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label Label13 
                  Caption         =   "Data inizio"
                  Height          =   255
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   164
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.Label Label11 
                  Caption         =   "Causale carico del tipo prodotto aumento peso"
                  Height          =   255
                  Index           =   8
                  Left            =   -70680
                  TabIndex        =   127
                  Top             =   5040
                  Width           =   4335
               End
               Begin VB.Label Label11 
                  Caption         =   "Causale carico del tipo prodotto calo peso"
                  Height          =   255
                  Index           =   7
                  Left            =   -70680
                  TabIndex        =   126
                  Top             =   3360
                  Width           =   3975
               End
               Begin VB.Label Label11 
                  Caption         =   "Causale carico del tipo prodotto scarto"
                  Height          =   255
                  Index           =   6
                  Left            =   -70680
                  TabIndex        =   123
                  Top             =   1680
                  Width           =   4335
               End
               Begin VB.Label Label5 
                  Caption         =   "Causali di magazzino"
                  Height          =   255
                  Index           =   3
                  Left            =   -63240
                  TabIndex        =   122
                  Top             =   1200
                  Width           =   2175
               End
               Begin VB.Label Label12 
                  Caption         =   "Tipo corpo fattura socio"
                  Height          =   255
                  Index           =   6
                  Left            =   -69600
                  TabIndex        =   119
                  Top             =   2760
                  Width           =   1935
               End
               Begin VB.Label Label14 
                  Caption         =   "N° etichette"
                  Height          =   255
                  Left            =   -63840
                  TabIndex        =   118
                  Top             =   3720
                  Width           =   1215
               End
               Begin VB.Label Label5 
                  Caption         =   "Tipo processo"
                  Height          =   255
                  Index           =   1
                  Left            =   -68880
                  TabIndex        =   112
                  Top             =   1200
                  Width           =   2055
               End
               Begin VB.Label Label5 
                  Caption         =   "Documento "
                  Height          =   255
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   111
                  Top             =   1200
                  Width           =   4215
               End
               Begin VB.Label Label8 
                  Caption         =   "Descrizione protocollo ICE"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   110
                  Top             =   1920
                  Width           =   5055
               End
               Begin VB.Label Label9 
                  Caption         =   "Progressivo"
                  Height          =   255
                  Index           =   0
                  Left            =   -59880
                  TabIndex        =   109
                  Top             =   1920
                  Width           =   1455
               End
               Begin VB.Label Label10 
                  Caption         =   "Sezionale"
                  Height          =   255
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   108
                  Top             =   1440
                  Width           =   3975
               End
               Begin VB.Label Label10 
                  Caption         =   "Documenti GreenTop"
                  Height          =   255
                  Index           =   1
                  Left            =   -68160
                  TabIndex        =   107
                  Top             =   1440
                  Width           =   3015
               End
               Begin VB.Label Label9 
                  Caption         =   "Esercizio"
                  Height          =   255
                  Index           =   1
                  Left            =   -62520
                  TabIndex        =   106
                  Top             =   1920
                  Width           =   2535
               End
               Begin VB.Label Label5 
                  Caption         =   "Tipo magazzino"
                  Height          =   255
                  Index           =   2
                  Left            =   -66120
                  TabIndex        =   105
                  Top             =   1200
                  Width           =   2055
               End
               Begin VB.Label Label6 
                  Caption         =   "Documento coop."
                  Height          =   255
                  Index           =   0
                  Left            =   -74760
                  TabIndex        =   104
                  Top             =   1080
                  Width           =   3375
               End
               Begin VB.Label Label11 
                  Caption         =   "Tipo prodotto di scarto"
                  Height          =   255
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   103
                  Top             =   1320
                  Width           =   2175
               End
               Begin VB.Label Label11 
                  Caption         =   "Causale scarico del tipo prodotto scarto"
                  Height          =   255
                  Index           =   1
                  Left            =   -70680
                  TabIndex        =   102
                  Top             =   1080
                  Width           =   4215
               End
               Begin VB.Label Label11 
                  Caption         =   "Tipo prodotto calo peso"
                  Height          =   255
                  Index           =   2
                  Left            =   -74880
                  TabIndex        =   101
                  Top             =   3000
                  Width           =   3375
               End
               Begin VB.Label Label11 
                  Caption         =   "Tipo prodotto aumento peso"
                  Height          =   255
                  Index           =   3
                  Left            =   -74880
                  TabIndex        =   100
                  Top             =   4680
                  Width           =   3375
               End
               Begin VB.Label Label11 
                  Caption         =   "Causale scarico del tipo prodotto calo peso"
                  Height          =   255
                  Index           =   4
                  Left            =   -70680
                  TabIndex        =   99
                  Top             =   2760
                  Width           =   4455
               End
               Begin VB.Label Label11 
                  Caption         =   "Causale scarico del tipo prodotto aumento peso"
                  Height          =   255
                  Index           =   5
                  Left            =   -70680
                  TabIndex        =   98
                  Top             =   4440
                  Width           =   4575
               End
               Begin VB.Label Label6 
                  Caption         =   "Tipo lavorazione per automazione lotto"
                  Height          =   255
                  Index           =   1
                  Left            =   -71280
                  TabIndex        =   97
                  Top             =   1080
                  Width           =   3855
               End
               Begin VB.Label Label12 
                  Caption         =   "Riga iniziale per fattura"
                  Height          =   255
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   96
                  Top             =   3360
                  Width           =   2415
               End
               Begin VB.Label Label12 
                  Caption         =   "Riga finale per fattura"
                  Height          =   255
                  Index           =   1
                  Left            =   -74880
                  TabIndex        =   95
                  Top             =   4800
                  Width           =   2415
               End
               Begin VB.Label lblPianodeiDeiConti 
                  Caption         =   "Piano dei conti per la riga di liquidazione"
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
                  Index           =   4
                  Left            =   -74880
                  MouseIcon       =   "frmMain.frx":4A77A
                  MousePointer    =   99  'Custom
                  TabIndex        =   94
                  Top             =   1530
                  Width           =   5055
               End
               Begin VB.Label Label12 
                  Caption         =   "Aliquota Iva "
                  Height          =   255
                  Index           =   2
                  Left            =   -67320
                  TabIndex        =   93
                  Top             =   2160
                  Width           =   2175
               End
               Begin VB.Label Label12 
                  Caption         =   "Tipo pagamento"
                  Height          =   255
                  Index           =   3
                  Left            =   -67320
                  TabIndex        =   92
                  Top             =   1560
                  Width           =   2655
               End
               Begin VB.Label Label12 
                  Caption         =   "Valuta"
                  Height          =   255
                  Index           =   4
                  Left            =   -69600
                  TabIndex        =   91
                  Top             =   2160
                  Width           =   2055
               End
               Begin VB.Label Label12 
                  Caption         =   "Causale contabile"
                  Height          =   255
                  Index           =   5
                  Left            =   -69600
                  TabIndex        =   90
                  Top             =   1560
                  Width           =   2655
               End
               Begin VB.Label lblPianodeiDeiConti 
                  Caption         =   "Piano dei conti per la riga positive"
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
                  Index           =   0
                  Left            =   -74880
                  MouseIcon       =   "frmMain.frx":4AA84
                  MousePointer    =   99  'Custom
                  TabIndex        =   89
                  Top             =   2160
                  Width           =   5055
               End
               Begin VB.Label lblPianodeiDeiConti 
                  Caption         =   "Piano dei conti per la riga negative"
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
                  Index           =   1
                  Left            =   -74880
                  MouseIcon       =   "frmMain.frx":4AD8E
                  MousePointer    =   99  'Custom
                  TabIndex        =   88
                  Top             =   2730
                  Width           =   5055
               End
               Begin VB.Label Label4 
                  Caption         =   "Data"
                  Height          =   255
                  Index           =   0
                  Left            =   -69720
                  TabIndex        =   87
                  Top             =   1410
                  Width           =   1095
               End
               Begin VB.Label Label4 
                  Caption         =   "Numero"
                  Height          =   255
                  Index           =   1
                  Left            =   -68280
                  TabIndex        =   86
                  Top             =   1410
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Caption         =   "Listino imballi di default"
                  Height          =   255
                  Index           =   2
                  Left            =   -74640
                  TabIndex        =   85
                  Top             =   3000
                  Width           =   3135
               End
               Begin VB.Label Label15 
                  Caption         =   "Ordine predefinito per merce in giacenza"
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
                  Left            =   -74640
                  TabIndex        =   84
                  Top             =   1200
                  Width           =   4215
               End
               Begin VB.Label Label15 
                  Caption         =   "Tipo pedana predefinita per lavorazione"
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
                  Left            =   -74640
                  TabIndex        =   83
                  Top             =   2160
                  Width           =   4815
               End
               Begin VB.Label Label16 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FFFF&
                  Caption         =   "Parametrizzazione evasione ordini per utente"
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
                  Index           =   0
                  Left            =   -74760
                  TabIndex        =   177
                  Top             =   1320
                  Width           =   10455
               End
               Begin VB.Label Label20 
                  Caption         =   "Utente"
                  Height          =   255
                  Left            =   -74760
                  TabIndex        =   188
                  Top             =   1680
                  Width           =   2775
               End
               Begin VB.Label Label16 
                  Alignment       =   2  'Center
                  BackColor       =   &H0000FFFF&
                  Caption         =   "Parametrizzazione ripesatura pedana"
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
                  Index           =   1
                  Left            =   -74760
                  TabIndex        =   265
                  Top             =   5880
                  Width           =   19215
               End
            End
            Begin DMTDataCmb.DMTCombo cboFiliale 
               Height          =   315
               Left            =   120
               TabIndex        =   0
               Top             =   360
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   135
               Left            =   7920
               TabIndex        =   251
               Top             =   765
               Visible         =   0   'False
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   238
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin VB.Label lblPercentualeIstat 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   6480
               TabIndex        =   114
               Top             =   2640
               Width           =   2055
            End
            Begin VB.Label Label3 
               Caption         =   "Filiale"
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
               Index           =   0
               Left            =   120
               TabIndex        =   113
               Top             =   120
               Width           =   4935
            End
         End
      End
      Begin DmtPrnDlgCtl.DMTDialog DmtPrnDlg 
         Left            =   720
         Top             =   1170
         _ExtentX        =   661
         _ExtentY        =   661
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
         Left            =   600
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   60
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   7875
         Left            =   0
         TabIndex        =   115
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   13891
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
         Top             =   360
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   1200
         MousePointer    =   9  'Size W E
         Top             =   0
         Width           =   60
      End
      Begin VB.Line Line2 
         X1              =   2040
         X2              =   6960
         Y1              =   3240
         Y2              =   3240
      End
   End
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   44
      Top             =   11400
      Width           =   20775
      _ExtentX        =   36645
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
'L'oggetto per la gestione dei sottodocumenti'''''''''''''''''''''''''''''''''''
'QUADRATURA
Private WithEvents m_DocumentsLink As DmtDocManLib.DocumentsLink                '
Attribute m_DocumentsLink.VB_VarHelpID = -1
'PROCESSI PER DOCUMENTO                                                         '
Private WithEvents m_DocumentsLink1 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink1.VB_VarHelpID = -1
'SEZIONALI PER DOCUMENTO
Private WithEvents m_DocumentsLink2 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink2.VB_VarHelpID = -1
'PROGRESSIVO PROTOCOLLO ICE
Private WithEvents m_DocumentsLink3 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink3.VB_VarHelpID = -1
'OPERAZIONE PER DOCUMENTO
Private WithEvents m_DocumentsLink4 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink4.VB_VarHelpID = -1
'OPERAZIONE PER ANNOTAZIONI PER DOCUMENTO
Private WithEvents m_DocumentsLink5 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink5.VB_VarHelpID = -1
'INTERESSI ANTICIPAZIONI
Private WithEvents m_DocumentsLink6 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink6.VB_VarHelpID = -1
'PARAMETRIZZAZIONE UTENTI PER EVASIONE ORDINI
Private WithEvents m_DocumentsLink7 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink7.VB_VarHelpID = -1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Private Const CAMPO_PER_CAPTION = "Filiale"


'Versione del controllo ActiveBar
Private Const BARMENUVERSION = "3.0"
'Variabile per la gestione degli shortcut del Menu
Private aryShortCut(1) As New ActiveBar3LibraryCtl.ShortCut


Public NuovoDocumento As Integer


'******************************************************************************



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
Dim Testo As String

    '///////////////////////////////////////////////////////////////////
    'Inserire qui il codice di controllo sulla validità e consistenza
    'dei dati da salvare.
    '///////////////////////////////////////////////////////////////////

    PermissionToSave = True
    
    If Len(CodiceAnnullaOperazione) > 0 Then
        If Mid(CodiceAnnullaOperazione, 1, 1) <> "R" Then
            MsgBox "Il codice per il ripristino delle operazioni nell'evasione ordini deve avere il carattere 'R' davanti", vbCritical, TheApp.FunctionName
            PermissionToSave = False
            Exit Function
        End If
    End If
    
    If Len(CodiceConfermaOrdine) > 0 Then
        If Mid(CodiceConfermaOrdine, 1, 1) <> "C" Then
            MsgBox "Il codice per la conferma dell'ordine nell'evasione ordini deve avere il carattere 'C' davanti", vbCritical, TheApp.FunctionName
            PermissionToSave = False
            Exit Function
        End If
    End If
    If Len(CodiceGestioneErrori) > 0 Then
        If Mid(CodiceGestioneErrori, 1, 1) <> "E" Then
            MsgBox "Il codice per la la gestione degli errori nell'evasione ordini deve avere il carattere 'E' davanti", vbCritical, TheApp.FunctionName
            PermissionToSave = False
            Exit Function
        End If
    End If
    
    If Me.cboFunzioneDDTAcq.CurrentID > 0 Then
        If GET_CONTROLLO_PROCESSO_FUNZIONE(Me.cboFunzioneDDTAcq.CurrentID) = False Then
            Testo = "ATTENZIONE!!!!" & vbCrLf
            Testo = Testo & "La funzione " & Me.cboFunzioneDDTAcq.Text & " non è compatibile per il passaggio di un conferimento ad un documento di acquisto"
            MsgBox Testo, vbCritical, "Controllo configurazione"
            PermissionToSave = False
            Exit Function
        End If
    End If
    
    If Me.cboFunzioneFAAcq.CurrentID > 0 Then
        If GET_CONTROLLO_PROCESSO_FUNZIONE(Me.cboFunzioneFAAcq.CurrentID) = False Then
            Testo = "ATTENZIONE!!!!" & vbCrLf
            Testo = Testo & "La funzione " & Me.cboFunzioneFAAcq.Text & " non è compatibile per il passaggio di un conferimento ad un documento di acquisto"
            MsgBox Testo, vbCritical, "Controllo configurazione"
            PermissionToSave = False
            Exit Function
        End If
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


    'Annulla una eventuale operazione precedente.
    If m_Document.TableNew Then
        m_Document.AbortNew
    End If

    'Creazione buffers vuoti
    m_Document.NewDoc
    
    
    'Refresh delle variabili di stato
    m_Search = False
    m_Changed = False
    m_Saved = False
    
    'Refresh della toolbar in modalità inserimento
    SetStatus4Modality Insert
    
    'Ripristina la vista del Form
    BrwMain.Visible = False
    
    
    'Il primo campo del Form riceve l'input focus
    SetFocusTabIndex0
    
        
    
    
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
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDAnagrafica = 0
        ctrControl.Description = ""
        ctrControl.SecondDescription = ""
        
    ElseIf sType = "dmtCurrency" Or sType = "dmtNumber" Then
        ctrControl.Value = 0
        ctrControl.Text = ""
    ElseIf sType = "DmtSearchACS" Then
        ctrControl.IDNode = 0
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
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtNumber"
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtCurrency"
                    
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "dmtTime"
                    Field.Control.Text = fnNormDate(m_Document.Fields(Field.Name).Value)
                Case "DmtSearchACS"
                        Field.Control.Description = fnNotNull(m_Document.Fields("Anagrafica").Value)
                        Field.Control.SecondDescription = fnNotNull(m_Document.Fields("Nome").Value)
                        Field.Control.IDAnagrafica = m_Document.Fields(Field.Name).Value
                Case "CheckBox"
                    Field.Control.Value = Abs(fnNotNullN((m_Document.Fields(Field.Name).Value)))
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
    
    
    'Filiale
    Set Field = New FormField
    Set Field.Control = Me.cboFiliale
    Field.Name = "IDFiliale"
    Field.Visible = True
    Me.cboFiliale.Tag = Field.Name
    m_FormFields.Add Field
    
    'Magazzino di carico
    Set Field = New FormField
    Set Field.Control = Me.cboMagazzinoCarico
    Field.Name = "IDMagazzino_Carico"
    Field.Visible = True
    Me.cboMagazzinoCarico.Tag = Field.Name
    m_FormFields.Add Field
    
    'Causale di carico del magazzino di carico
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleCarico_Car
    Field.Name = "IDCausale_Carico_Mag_Carico"
    Field.Visible = True
    Me.cboCausaleCarico_Car.Tag = Field.Name
    m_FormFields.Add Field
    
    'Causale di scarico del magazzino di carico
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleScarico_Car
    Field.Name = "IDCausale_Scarico_Mag_Carico"
    Field.Visible = True
    Me.cboCausaleScarico_Car.Tag = Field.Name
    m_FormFields.Add Field

    'Magazzino di vendita
    Set Field = New FormField
    Set Field.Control = Me.cboMagazzinoVendita
    Field.Name = "IDMagazzino_Vendita"
    Field.Visible = True
    Me.cboMagazzinoVendita.Tag = Field.Name
    m_FormFields.Add Field
    
    'Causale di carico del magazzino di vendita
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleCarico_Vend
    Field.Name = "IDCausale_Carico_Mag_Vendita"
    Field.Visible = True
    Me.cboCausaleCarico_Vend.Tag = Field.Name
    m_FormFields.Add Field
    
    'Causale di scarico del magazzino di vendita
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleScarico_Vend
    Field.Name = "IDCausale_Scarico_Mag_Vendita"
    Field.Visible = True
    Me.cboCausaleScarico_Vend.Tag = Field.Name
    m_FormFields.Add Field
    
    'Tipo prodotto imballo
    Set Field = New FormField
    Set Field.Control = Me.cboTipoProdottoImballo
    Field.Name = "IDTipoImballo"
    Field.Visible = True
    Me.cboTipoProdottoImballo.Tag = Field.Name
    m_FormFields.Add Field
    
    'Tipo prodotto grezzo
    Set Field = New FormField
    Set Field.Control = Me.cboTipoProdottoGrezzo
    Field.Name = "IDTipoGrezzo"
    Field.Visible = True
    Me.cboTipoProdottoGrezzo.Tag = Field.Name
    m_FormFields.Add Field
    
    'Tipo prodotto lavorato
    Set Field = New FormField
    Set Field.Control = Me.cboTipoProdottoLavorato
    Field.Name = "IDTipoLavorato"
    Field.Visible = True
    Me.cboTipoProdottoLavorato.Tag = Field.Name
    m_FormFields.Add Field
    
    'Categoria anagrafica per identificare un socio
    Set Field = New FormField
    Set Field.Control = Me.cboCategoriaAnagraficaSocio
    Field.Name = "IDCategoriaAnagrafica"
    Field.Visible = True
    Me.cboCategoriaAnagraficaSocio.Tag = Field.Name
    m_FormFields.Add Field

    'Quantita minima richiesta per chiudere un lotto di conferimento
    Set Field = New FormField
    Set Field.Control = Me.txtQtaMinimaPerConferimento
    Field.Name = "QtaMinimaConfPerChiusura"
    Field.Visible = True
    Me.txtQtaMinimaPerConferimento.Tag = Field.Name
    m_FormFields.Add Field
    
    'Quantita minima richiesta per chiudere un lotto di vendita
    Set Field = New FormField
    Set Field.Control = Me.txtQtaMinimaPerVendita
    Field.Name = "QtaMinimaVendPerChiusura"
    Field.Visible = True
    Me.txtQtaMinimaPerVendita.Tag = Field.Name
    m_FormFields.Add Field

    'Indica se l'IVA dell'imballo deve essere uguale all'IVA della riga dell'articolo
    Set Field = New FormField
    Set Field.Control = Me.chkIvaBloccata
    Field.Name = "IvaBloccata"
    Field.Visible = True
    Me.chkIvaBloccata.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica il tipo prodotto per identificare un prodotto di calo peso
    Set Field = New FormField
    Set Field.Control = Me.cboTipoCaloPeso
    Field.Name = "IDTipoCaloPeso"
    Field.Visible = True
    Me.cboTipoCaloPeso.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica il tipo prodotto per identificare un prodotto di scarto
    Set Field = New FormField
    Set Field.Control = Me.cboTipoScarto
    Field.Name = "IDTipoScarto"
    Field.Visible = True
    Me.cboTipoScarto.Tag = Field.Name
    m_FormFields.Add Field

    'Indica il tipo prodotto per identificare un prodotto di aumento di peso
    Set Field = New FormField
    Set Field.Control = Me.cboTipoAumentoPeso
    Field.Name = "IDTipoAumentoPeso"
    Field.Visible = True
    Me.cboTipoAumentoPeso.Tag = Field.Name
    m_FormFields.Add Field

    'Indica la causale del prodotto di calo peso
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleCaloPeso
    Field.Name = "IDCausaleCaloPeso"
    Field.Visible = True
    Me.cboCausaleCaloPeso.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica la causale di carico del prodotto di calo peso
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleScaricoCaloPeso
    Field.Name = "IDCausaleCaloPesoCarico"
    Field.Visible = True
    Me.cboCausaleScaricoCaloPeso.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica la causale del prodotto di scarto
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleTipoScarto
    Field.Name = "IDCausaleScarto"
    Field.Visible = True
    Me.cboCausaleTipoScarto.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica la causale di carico di un prodotto di scarto
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleScaricoTipoScarto
    Field.Name = "IDCausaleScartoCarico"
    Field.Visible = True
    Me.cboCausaleScaricoTipoScarto.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica la causale del prodotto di aumento di peso
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleAumentoPeso
    Field.Name = "IDCausaleAumentoPeso"
    Field.Visible = True
    Me.cboCausaleAumentoPeso.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica la causale di scarico di un prodotto di aumento di peso
    Set Field = New FormField
    Set Field.Control = Me.cboCausaleScaricoAumentoPeso
    Field.Name = "IDCausaleAumentoPesoCarico"
    Field.Visible = True
    Me.cboCausaleScaricoAumentoPeso.Tag = Field.Name
    m_FormFields.Add Field
    
    'Indica il tipo di quadratura
    Set Field = New FormField
    Set Field.Control = Me.cboTipoQuadratura
    Field.Name = "IDTipoQuadratura"
    Field.Visible = True
    Me.cboTipoQuadratura.Tag = Field.Name
    m_FormFields.Add Field

    'BNDOO
    Set Field = New FormField
    Set Field.Control = Me.txtParametroBNDOO
    Field.Name = "BNDOO"
    Field.Visible = True
    Me.txtParametroBNDOO.Tag = Field.Name
    m_FormFields.Add Field

    'Riga fattura iniziale
    Set Field = New FormField
    Set Field.Control = Me.txtRigaInizialePerFattura
    Field.Name = "RigaFatturazioneIniziale"
    Field.Visible = True
    Me.txtRigaInizialePerFattura.Tag = Field.Name
    m_FormFields.Add Field

    'Riga fattura finale
    Set Field = New FormField
    Set Field.Control = Me.txtRigaFinalePerFattura
    Field.Name = "RigaFatturazioneFinale"
    Field.Visible = True
    Me.txtRigaFinalePerFattura.Tag = Field.Name
    m_FormFields.Add Field

    'Codice del piano dei conti della riga di liquidazione
    Set Field = New FormField
    Set Field.Control = Me.txtCodiceConto
    Field.Name = "PDCCodiceFatturazione"
    Field.Visible = True
    Me.txtCodiceConto.Tag = Field.Name
    m_FormFields.Add Field

    'Descrizione del piano dei conti della riga di liquidazione
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneConto
    Field.Name = "PDCDescrizioneFatturazione"
    Field.Visible = True
    Me.txtDescrizioneConto.Tag = Field.Name
    m_FormFields.Add Field

    'Codice del piano dei conti della riga di liquidazione negativa
    Set Field = New FormField
    Set Field.Control = Me.txtCodiceConto_RigaNeg
    Field.Name = "PDCCodiceRigheLiqNegative"
    Field.Visible = True
    Me.txtCodiceConto_RigaNeg.Tag = Field.Name
    m_FormFields.Add Field

    'Descrizione del piano dei conti della riga di liquidazione positiva
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneConto_RigaNeg
    Field.Name = "PDCDescrizioneRigheLiqNegative"
    Field.Visible = True
    Me.txtDescrizioneConto_RigaNeg.Tag = Field.Name
    m_FormFields.Add Field

    'Codice del piano dei conti della riga di liquidazione positiva
    Set Field = New FormField
    Set Field.Control = Me.txtCodiceConto_RigaPos
    Field.Name = "PDCCodiceRigheLiqPositive"
    Field.Visible = True
    Me.txtCodiceConto_RigaPos.Tag = Field.Name
    m_FormFields.Add Field

    'Descrizione del piano dei conti della riga di liquidazione positiva
    Set Field = New FormField
    Set Field.Control = Me.txtDescrizioneConto_RigaPos
    Field.Name = "PDCDescrizioneRigheLiqPositive"
    Field.Visible = True
    Me.txtDescrizioneConto_RigaPos.Tag = Field.Name
    m_FormFields.Add Field

    'Aliquota IVA di fatturazione
    Set Field = New FormField
    Set Field.Control = Me.cboIvaPerFatturazione
    Field.Name = "IDIvaFatturazione"
    Field.Visible = True
    Me.cboIvaPerFatturazione.Tag = Field.Name
    m_FormFields.Add Field

    'Pagamento di fatturazione
    Set Field = New FormField
    Set Field.Control = Me.cboPagamentoPerFatturazione
    Field.Name = "IDPagamentoFatturazione"
    Field.Visible = True
    Me.cboPagamentoPerFatturazione.Tag = Field.Name
    m_FormFields.Add Field

    'Valuta di fatturazione
    Set Field = New FormField
    Set Field.Control = Me.cboValutaPerFatturazione
    Field.Name = "IDValutaFatturazione"
    Field.Visible = True
    Me.cboValutaPerFatturazione.Tag = Field.Name
    m_FormFields.Add Field

    'causale contabile
    Set Field = New FormField
    Set Field.Control = Me.cboCausContPerFatturazione
    Field.Name = "IDCausaleContabileFatturazione"
    Field.Visible = True
    Me.cboCausContPerFatturazione.Tag = Field.Name
    m_FormFields.Add Field

    'Numerazione del lotto
    Set Field = New FormField
    Set Field.Control = Me.txtNumerazioneLotto
    Field.Name = "NumerazioneLotto"
    Field.Visible = True
    Me.txtNumerazioneLotto.Tag = Field.Name
    m_FormFields.Add Field

    'Numerazione del lotto di conferimento
    Set Field = New FormField
    Set Field.Control = Me.txtNumerazioneLottoConf
    Field.Name = "NumerazioneLottoConferimento"
    Field.Visible = True
    Me.txtNumerazioneLottoConf.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo di arrotondamento lavorazione
    Set Field = New FormField
    Set Field.Control = Me.cboTipoArrotondamento
    Field.Name = "IDTipoArrotondamento"
    Field.Visible = True
    Me.cboTipoArrotondamento.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo di arrotondamento conferimento
    Set Field = New FormField
    Set Field.Control = Me.cboTipoArrotondamentoConf
    Field.Name = "IDTipoArrotondamentoConferimento"
    Field.Visible = True
    Me.cboTipoArrotondamentoConf.Tag = Field.Name
    m_FormFields.Add Field

    'Gestione chiusura conferimento dalla vendita
    Set Field = New FormField
    Set Field.Control = Me.chkGestioneConferimento
    Field.Name = "GestioneConferimento"
    Field.Visible = True
    Me.chkGestioneConferimento.Tag = Field.Name
    m_FormFields.Add Field
    
    'Identificativo cliente cooperativa per ordine in giacenza
    Set Field = New FormField
    Set Field.Control = Me.CDClienteOrdinePred
    Field.Name = "IDClienteCoop"
    Field.Visible = True
    Me.CDClienteOrdinePred.Tag = Field.Name
    m_FormFields.Add Field
    
    'Data ordine predefinito
    Set Field = New FormField
    Set Field.Control = Me.txtDataOrdinePred
    Field.Name = "DataOrdineCoop"
    Field.Visible = True
    Me.txtDataOrdinePred.Tag = Field.Name
    m_FormFields.Add Field
    
    'Numero ordine predefinito
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroOrdinePred
    Field.Name = "NumeroOrdineCoop"
    Field.Visible = True
    Me.txtNumeroOrdinePred.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo pedana di default
    Set Field = New FormField
    Set Field.Control = Me.CDTipoPedana
    Field.Name = "IDTipoPedanaDefault"
    Field.Visible = True
    Me.CDTipoPedana.Tag = Field.Name
    m_FormFields.Add Field

    'Listino imballi di default
    Set Field = New FormField
    Set Field.Control = Me.cboListinoImballiDefault
    Field.Name = "IDListinoImballiDefault"
    Field.Visible = True
    Me.cboListinoImballiDefault.Tag = Field.Name
    m_FormFields.Add Field

'    'Tipo gestione articoli vendita
'    Set Field = New FormField
'    Set Field.Control = Me.cboTipoGestioneArticoliVend
'    Field.Name = "IDTipoGestioneArticoliVendita"
'    Field.Visible = True
'    Me.cboTipoGestioneArticoliVend.Tag = Field.Name
'    m_FormFields.Add Field


    'Codice associato
    Set Field = New FormField
    Set Field.Control = Me.txtCodiceAssociato
    Field.Name = "CodiceAssociato"
    Field.Visible = True
    Me.txtCodiceAssociato.Tag = Field.Name
    m_FormFields.Add Field

    'Stampa pedana veloce
    Set Field = New FormField
    Set Field.Control = Me.chkStampaPedana
    Field.Name = "StampaPedana"
    Field.Visible = True
    Me.chkStampaPedana.Tag = Field.Name
    m_FormFields.Add Field

    'Numero etichette pedana veloce
    Set Field = New FormField
    Set Field.Control = Me.txtNumeroEtichettePedana
    Field.Name = "NumeroEtichettePedana"
    Field.Visible = True
    Me.txtNumeroEtichettePedana.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo corpo fattura
    Set Field = New FormField
    Set Field.Control = Me.cboTipoCorpoFattSocio
    Field.Name = "IDTipoCorpoFatturaSocio"
    Field.Visible = True
    Me.cboTipoCorpoFattSocio.Tag = Field.Name
    m_FormFields.Add Field

    'Nuovo calcolo
    Set Field = New FormField
    Set Field.Control = Me.chkNuovoCalcolo
    Field.Name = "AttivazioneNuovoMetodoCalcolo"
    Field.Visible = True
    Me.chkNuovoCalcolo.Tag = Field.Name
    m_FormFields.Add Field

    'Lingua
    Set Field = New FormField
    Set Field.Control = Me.cboLingua
    Field.Name = "IDLinguaPredefinita"
    Field.Visible = True
    Me.cboLingua.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo scelta articolo conferito
    Set Field = New FormField
    Set Field.Control = Me.cboTipoSceltaArticoloConferito
    Field.Name = "IDRV_POTipoSceltaArticoloLottoCampagna"
    Field.Visible = True
    Me.cboTipoSceltaArticoloConferito.Tag = Field.Name
    m_FormFields.Add Field

    'Lotto di campagna obbligatorio
    Set Field = New FormField
    Set Field.Control = Me.chkLottoCampagnaObbligatorio
    Field.Name = "LottoCampagnaObbligatorio"
    Field.Visible = True
    Me.chkLottoCampagnaObbligatorio.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo comportamento lavorazione
    Set Field = New FormField
    Set Field.Control = Me.cboTipoComportamentoLavorazione
    Field.Name = "IDRV_POTipoComportamentoLavorazione"
    Field.Visible = True
    Me.cboTipoComportamentoLavorazione.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo peso articolo
    Set Field = New FormField
    Set Field.Control = Me.cboTipoPesoArticolo
    Field.Name = "IDRV_POTipoPesoArticolo"
    Field.Visible = True
    Me.cboTipoPesoArticolo.Tag = Field.Name
    m_FormFields.Add Field

    'Liquidazione lordo IVA
    Set Field = New FormField
    Set Field.Control = Me.chkLiquidazioneLordoIVA
    Field.Name = "LiquidazioneInclusoIVA"
    Field.Visible = True
    Me.chkLiquidazioneLordoIVA.Tag = Field.Name
    m_FormFields.Add Field

    'Riferimento liquidazione
    Set Field = New FormField
    Set Field.Control = Me.chkRipDescrRifLiq
    Field.Name = "RiportaDescrizioneRifLiq"
    Field.Visible = True
    Me.chkRipDescrRifLiq.Tag = Field.Name
    m_FormFields.Add Field

    'Stampa con etichette nuove
    Set Field = New FormField
    Set Field.Control = Me.chkStampaNuovoMetodo
    Field.Name = "StampaEtichetteNuove"
    Field.Visible = True
    Me.chkStampaNuovoMetodo.Tag = Field.Name
    m_FormFields.Add Field

    'Iscrizione albo cooperativa
    Set Field = New FormField
    Set Field.Control = Me.txtIscrizioneAlboCoop
    Field.Name = "IscrizioneAlboCoop"
    Field.Visible = True
    Me.txtIscrizioneAlboCoop.Tag = Field.Name
    m_FormFields.Add Field

    'IDentificativo dell'unita di misura predefinita
    Set Field = New FormField
    Set Field.Control = Me.cboUMRigaOrdine
    Field.Name = "IDRV_POTipoUMOrdine"
    Field.Visible = True
    Me.cboUMRigaOrdine.Tag = Field.Name
    m_FormFields.Add Field

    'IDentificativo del sezionale predefinito del CMR
    Set Field = New FormField
    Set Field.Control = Me.cboSezionaleCMR
    Field.Name = "IDSezionaleCMRPred"
    Field.Visible = True
    Me.cboSezionaleCMR.Tag = Field.Name
    m_FormFields.Add Field

    'Aliquota IVA imballo a rendere
    Set Field = New FormField
    Set Field.Control = Me.chkIvaARendere
    Field.Name = "IvaImballoARendere"
    Field.Visible = True
    Me.chkIvaARendere.Tag = Field.Name
    m_FormFields.Add Field

    'Visualizza totali ordini da preparare all'avivo della gestione
    Set Field = New FormField
    Set Field.Control = Me.chkVisOrdPrep
    Field.Name = "VisualizzaTotaliOrdinePrep"
    Field.Visible = True
    Me.chkVisOrdPrep.Tag = Field.Name
    m_FormFields.Add Field

    'Visualizza totali ordini da smistare all'avivo della gestione
    Set Field = New FormField
    Set Field.Control = Me.chkVisOrdSmist
    Field.Name = "VisualizzaTotaliOrdineSmist"
    Field.Visible = True
    Me.chkVisOrdSmist.Tag = Field.Name
    m_FormFields.Add Field

    'Modo di trasporto instrastat
    Set Field = New FormField
    Set Field.Control = Me.CDModoTrasporto
    Field.Name = "IDModoTrasportoIntra"
    Field.Visible = True
    Me.CDModoTrasporto.Tag = Field.Name
    m_FormFields.Add Field

    'Natura di transazione instrastat
    Set Field = New FormField
    Set Field.Control = Me.CDNaturaTransazione
    Field.Name = "IDNaturaTransazione"
    Field.Visible = True
    Me.CDNaturaTransazione.Tag = Field.Name
    m_FormFields.Add Field
    
    'Pedana automatica
    Set Field = New FormField
    Set Field.Control = Me.chkPedanaAutomatica
    Field.Name = "PedanaAutomatica"
    Field.Visible = True
    Me.chkPedanaAutomatica.Tag = Field.Name
    m_FormFields.Add Field

    'Ordine automatico
    Set Field = New FormField
    Set Field.Control = Me.chkOrdineAutomatico
    Field.Name = "OrdineAutomatico"
    Field.Visible = True
    Me.chkOrdineAutomatico.Tag = Field.Name
    m_FormFields.Add Field

    'Listino per campionatura
    Set Field = New FormField
    Set Field.Control = Me.cboListinoCampionatura
    Field.Name = "IDListinoCampionatura"
    Field.Visible = True
    Me.cboListinoCampionatura.Tag = Field.Name
    m_FormFields.Add Field

    'Decimali conferimento
    Set Field = New FormField
    Set Field.Control = Me.cboDecimaliConf
    Field.Name = "IDRV_POTipoDecimaliPesiConferimento"
    Field.Visible = True
    Me.cboDecimaliConf.Tag = Field.Name
    m_FormFields.Add Field

    'Decimali lavorazione
    Set Field = New FormField
    Set Field.Control = Me.cboDecimaliLav
    Field.Name = "IDRV_POTipoDecimaliPesiLavorazione"
    Field.Visible = True
    Me.cboDecimaliLav.Tag = Field.Name
    m_FormFields.Add Field

    'Decimali vendita
    Set Field = New FormField
    Set Field.Control = Me.cboDecimaliVend
    Field.Name = "IDRV_POTipoDecimaliPesiVendita"
    Field.Visible = True
    Me.cboDecimaliVend.Tag = Field.Name
    m_FormFields.Add Field

    'Visualizza andamento ordine
    Set Field = New FormField
    Set Field.Control = Me.chkVisAndamDaOrd
    Field.Name = "VisAndamentoOrdDaOrd"
    Field.Visible = True
    Me.chkVisAndamDaOrd.Tag = Field.Name
    m_FormFields.Add Field

    'Visualizza andamento ordine da lavorazione
    Set Field = New FormField
    Set Field.Control = Me.chkVisAndamLav
    Field.Name = "VisAndamentoOrdDaLav"
    Field.Visible = True
    Me.chkVisAndamLav.Tag = Field.Name
    m_FormFields.Add Field

    'IDOrdine IV Gamma
    Set Field = New FormField
    Set Field.Control = Me.txtIDOrdineIVGamma
    Field.Name = "IDOrdineIVGamma"
    Field.Visible = True
    Me.txtIDOrdineIVGamma.Tag = Field.Name
    m_FormFields.Add Field

    'IDOrdine da IV Gamma per lavorazione
    Set Field = New FormField
    Set Field.Control = Me.txtIDOrdineIVGammaLav
    Field.Name = "IDOrdineLavorazioneIVGamma"
    Field.Visible = True
    Me.txtIDOrdineIVGammaLav.Tag = Field.Name
    m_FormFields.Add Field

    'Visualizza etichette Pro lavorazione per utente
    Set Field = New FormField
    Set Field.Control = Me.chkVisEtiLavUtente
    Field.Name = "EtichetteProPerUtenteLav"
    Field.Visible = True
    Me.chkVisEtiLavUtente.Tag = Field.Name
    m_FormFields.Add Field

    'Visualizza etichette Pro pedana per utente
    Set Field = New FormField
    Set Field.Control = Me.chkVisEtiPedUtente
    Field.Name = "EtichetteProPerUtentePed"
    Field.Visible = True
    Me.chkVisEtiPedUtente.Tag = Field.Name
    m_FormFields.Add Field

    'Visualizza riepilogo conferimenti nei documenti di vendita
    Set Field = New FormField
    Set Field.Control = Me.chKVisRiepConfVend
    Field.Name = "VisRiepConfInVend"
    Field.Visible = True
    Me.chKVisRiepConfVend.Tag = Field.Name
    m_FormFields.Add Field

    'Prezzi articoli da ordine
    Set Field = New FormField
    Set Field.Control = Me.chkPrezziArtDaOrdine
    Field.Name = "PrezziArticoloDaOrdine"
    Field.Visible = True
    Me.chkPrezziArtDaOrdine.Tag = Field.Name
    m_FormFields.Add Field



    
    'Numeri di zeri davanti al numero documento
    Set Field = New FormField
    Set Field.Control = Me.chkZeriRifDoc
    Field.Name = "NumeroZeriRifDoc"
    Field.Visible = True
    Me.chkZeriRifDoc.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo spaccatura lotto di lavorazione
    Set Field = New FormField
    Set Field.Control = Me.cboTipoSpezzatura
    Field.Name = "IDRV_POTipoSpaccaturaLavorazione"
    Field.Visible = True
    Me.cboTipoSpezzatura.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo funzione DDT acquisto per conferimento
    Set Field = New FormField
    Set Field.Control = Me.cboFunzioneDDTAcq
    Field.Name = "IDFunzioneDDTAcqConf"
    Field.Visible = True
    Me.cboFunzioneDDTAcq.Tag = Field.Name
    m_FormFields.Add Field

    'Tipo funzione FA acquisto per conferimento
    Set Field = New FormField
    Set Field.Control = Me.cboFunzioneFAAcq
    Field.Name = "IDFunzioneFAAcqConf"
    Field.Visible = True
    Me.cboFunzioneFAAcq.Tag = Field.Name
    m_FormFields.Add Field

    'Prezzo medio da conferimento
    Set Field = New FormField
    Set Field.Control = Me.chkPrezzoMedioDaConf
    Field.Name = "PrezzoMedioDaConf"
    Field.Visible = True
    Me.chkPrezzoMedioDaConf.Tag = Field.Name
    m_FormFields.Add Field

    'Aggiorna prezzo medio da conferimento a tutti i collegamenti del conferimento
    Set Field = New FormField
    Set Field.Control = Me.chkAggPrezzoMedioDaConf
    Field.Name = "AggiornaPrezzoMedioDaConf"
    Field.Visible = True
    Me.chkAggPrezzoMedioDaConf.Tag = Field.Name
    m_FormFields.Add Field

    'Aggiorna tipo di lavorazione da conferimento a tutti i collegamenti del conferimento
    Set Field = New FormField
    Set Field.Control = Me.chkAggTipoLavDaConf
    Field.Name = "AggiornaTipoLavDaConf"
    Field.Visible = True
    Me.chkAggTipoLavDaConf.Tag = Field.Name
    m_FormFields.Add Field

    'Articolo di quadratura per una pesatura negativa (Scarto o calo peso)
    Set Field = New FormField
    Set Field.Control = Me.CDArtPesPedNeg
    Field.Name = "IDArtQuadNegDaPesPed"
    Field.Visible = True
    Me.CDArtPesPedNeg.Tag = Field.Name
    m_FormFields.Add Field

    'Articolo di quadratura per una pesatura positiva (Aumento di peso)
    Set Field = New FormField
    Set Field.Control = Me.CDArtPesPedPos
    Field.Name = "IDArtQuadPosDaPesPed"
    Field.Visible = True
    Me.CDArtPesPedPos.Tag = Field.Name
    m_FormFields.Add Field

    'RiportaInOrdineClienteDaPesPed
    Set Field = New FormField
    Set Field.Control = Me.chkRipPedOrdSel
    Field.Name = "RiportaInOrdineClienteDaPesPed"
    Field.Visible = True
    Me.chkRipPedOrdSel.Tag = Field.Name
    m_FormFields.Add Field
    
    'RiportaInOrdineClienteDaPesPed
    Set Field = New FormField
    Set Field.Control = Me.chkVisMasAvvioVeloce
    Field.Name = "ProponiMaschereAvvioVeloce"
    Field.Visible = True
    Me.chkVisMasAvvioVeloce.Tag = Field.Name
    m_FormFields.Add Field

    'PesoArticoloPedanaInFattura
    Set Field = New FormField
    Set Field.Control = Me.chkPesoArtPed
    Field.Name = "PesoArticoloPedanaInFattura"
    Field.Visible = True
    Me.chkPesoArtPed.Tag = Field.Name
    m_FormFields.Add Field
    
    
    'NumerazioneCooperativaFACS
    Set Field = New FormField
    Set Field.Control = Me.chkNumerazioneFACSCoop
    Field.Name = "NumerazioneCooperativaFACS"
    Field.Visible = True
    Me.chkNumerazioneFACSCoop.Tag = Field.Name
    m_FormFields.Add Field
    
    'FocusColliIVGamma
    Set Field = New FormField
    Set Field.Control = Me.chkFocusColliIVGamma
    Field.Name = "FocusColliIVGamma"
    Field.Visible = True
    Me.chkFocusColliIVGamma.Tag = Field.Name
    m_FormFields.Add Field
    
    'MsgConferimentoNegativo
    Set Field = New FormField
    Set Field.Control = Me.MsgConfNeg
    Field.Name = "MsgConferimentoNegativo"
    Field.Visible = True
    Me.MsgConfNeg.Tag = Field.Name
    m_FormFields.Add Field
    
    'ContributiUE
    Set Field = New FormField
    Set Field.Control = Me.chkContributi
    Field.Name = "ContributiUE"
    Field.Visible = True
    Me.chkContributi.Tag = Field.Name
    m_FormFields.Add Field
    
    'DescrizioneContributiUE
    Set Field = New FormField
    Set Field.Control = Me.txtContributi
    Field.Name = "DescrizioneContributiUE"
    Field.Visible = True
    Me.txtContributi.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDRV_POTipoCategoriaCommTrasp
    Set Field = New FormField
    Set Field.Control = Me.cboCatCommTrasporto
    Field.Name = "IDRV_POTipoCategoriaCommTrasp"
    Field.Visible = True
    Me.cboCatCommTrasporto.Tag = Field.Name
    m_FormFields.Add Field
    
    'CalcolaGestioneImballi
    Set Field = New FormField
    Set Field.Control = Me.chkStampaRiepilogoImballiConf
    Field.Name = "CalcolaGestioneImballi"
    Field.Visible = True
    Me.chkStampaRiepilogoImballiConf.Tag = Field.Name
    m_FormFields.Add Field
    
    'AttivaGestioneOrdineVivaio
    Set Field = New FormField
    Set Field.Control = Me.chkGestioneVivaio
    Field.Name = "AttivaGestioneOrdineVivaio"
    Field.Visible = True
    Me.chkGestioneVivaio.Tag = Field.Name
    m_FormFields.Add Field

    'AttivaGestioneOrdineVivaio
    Set Field = New FormField
    Set Field.Control = Me.cboTipoTrattAggVivaio
    Field.Name = "IDRV_POTipoTrattenutaAggiuntivaVivaio"
    Field.Visible = True
    Me.cboTipoTrattAggVivaio.Tag = Field.Name
    m_FormFields.Add Field

    'IDArticoloCommConf
    Set Field = New FormField
    Set Field.Control = Me.CDArticoloCommConf
    Field.Name = "IDArticoloCommConf"
    Field.Visible = True
    Me.CDArticoloCommConf.Tag = Field.Name
    m_FormFields.Add Field
    
    'PrezzoMedioConfInVendita
    Set Field = New FormField
    Set Field.Control = Me.chkRiporaPMConfVend
    Field.Name = "PrezzoMedioConfInVendita"
    Field.Visible = True
    Me.chkRiporaPMConfVend.Tag = Field.Name
    m_FormFields.Add Field
    
    'AbilitaRicercaFornitoriLottoCampagna
    Set Field = New FormField
    Set Field.Control = Me.chkAbilitaForLotto
    Field.Name = "AbilitaRicercaFornitoriLottoCampagna"
    Field.Visible = True
    Me.chkAbilitaForLotto.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDPortoNoCalcoloTrasportoComm
    Set Field = New FormField
    Set Field.Control = Me.cboPorto
    Field.Name = "IDPortoNoCalcoloTrasportoComm"
    Field.Visible = True
    Me.cboPorto.Tag = Field.Name
    m_FormFields.Add Field
    
    'PartenzaDocVendNuovoRecord
    Set Field = New FormField
    Set Field.Control = Me.chkDocVendForzaNuovo
    Field.Name = "PartenzaDocVendNuovoRecord"
    Field.Visible = True
    Me.chkDocVendForzaNuovo.Tag = Field.Name
    m_FormFields.Add Field
    
    'AbilitaCalcoloPesoTrasportoTipoPedana
    Set Field = New FormField
    Set Field.Control = Me.chkCalcoloTraspPerPeso
    Field.Name = "AbilitaCalcoloPesoTrasportoTipoPedana"
    Field.Visible = True
    Me.chkCalcoloTraspPerPeso.Tag = Field.Name
    m_FormFields.Add Field

    'AttivaCalcoloPesoLordoConferimento
    Set Field = New FormField
    Set Field.Control = Me.chkAttivaCalcoloPesoConf
    Field.Name = "AttivaCalcoloPesoLordoConferimento"
    Field.Visible = True
    Me.chkAttivaCalcoloPesoConf.Tag = Field.Name
    m_FormFields.Add Field

    'AbilitaImportoRiepilogoConferimento
    Set Field = New FormField
    Set Field.Control = Me.chkVisImportoF4
    Field.Name = "AbilitaImportoRiepilogoConferimento"
    Field.Visible = True
    Me.chkVisImportoF4.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDSezionaleListaPrelievoOrdine
    Set Field = New FormField
    Set Field.Control = Me.cboSezListaPrelievo
    Field.Name = "IDSezionaleListaPrelievoOrdine"
    Field.Visible = True
    Me.cboSezListaPrelievo.Tag = Field.Name
    m_FormFields.Add Field

    'NumeroListaOrdineCoop
    Set Field = New FormField
    Set Field.Control = Me.txtNListaOrdinePred
    Field.Name = "NumeroListaOrdineCoop"
    Field.Visible = True
    Me.txtNListaOrdinePred.Tag = Field.Name
    m_FormFields.Add Field

    'StampaDocEvasioneOrdAttivo
    Set Field = New FormField
    Set Field.Control = Me.chkStampaEvOrdAtt
    Field.Name = "StampaDocEvasioneOrdAttivo"
    Field.Visible = True
    Me.chkStampaEvOrdAtt.Tag = Field.Name
    m_FormFields.Add Field

    'AbilitaImportiConferimento
    Set Field = New FormField
    Set Field.Control = Me.chkForzaTabulazioneConf
    Field.Name = "AbilitaImportiConferimento"
    Field.Visible = True
    Me.chkForzaTabulazioneConf.Tag = Field.Name
    m_FormFields.Add Field
    
    
    'AttivaUMDocVendita
    Set Field = New FormField
    Set Field.Control = Me.chkAttivaUMDocVend
    Field.Name = "AttivaUMDocVendita"
    Field.Visible = True
    Me.chkAttivaUMDocVend.Tag = Field.Name
    m_FormFields.Add Field

    'NuovoRecordInLav
    Set Field = New FormField
    Set Field.Control = Me.chkNuovoRecLav
    Field.Name = "NuovoRecordInLav"
    Field.Visible = True
    Me.chkNuovoRecLav.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDLuogoPresaMercePredefinito
    Set Field = New FormField
    Set Field.Control = Me.cboLuogoPresaMerce
    Field.Name = "IDLuogoPresaMercePredefinito"
    Field.Visible = True
    Me.cboLuogoPresaMerce.Tag = Field.Name
    m_FormFields.Add Field

    'VisListaMerceOrdinata
    Set Field = New FormField
    Set Field.Control = Me.chkVisListaMerceOrd
    Field.Name = "VisListaMerceOrdinata"
    Field.Visible = True
    Me.chkVisListaMerceOrd.Tag = Field.Name
    m_FormFields.Add Field

    'NonAbilitareImportoImballo
    Set Field = New FormField
    Set Field.Control = Me.chkEscludiImbImpOrd
    Field.Name = "NonAbilitareImportoImballo"
    Field.Visible = True
    Me.chkEscludiImbImpOrd.Tag = Field.Name
    m_FormFields.Add Field

    'RipDiffColliDaLavDaOrd
    Set Field = New FormField
    Set Field.Control = Me.chkRipDiffColliDaLavDaOrd
    Field.Name = "RipDiffColliDaLavDaOrd"
    Field.Visible = True
    Me.chkRipDiffColliDaLavDaOrd.Tag = Field.Name
    m_FormFields.Add Field

    'UsaProtocolloICEPeriodo
    Set Field = New FormField
    Set Field.Control = Me.chkUsaProtICEPeriodo
    Field.Name = "UsaProtocolloICEPeriodo"
    Field.Visible = True
    Me.chkUsaProtICEPeriodo.Tag = Field.Name
    m_FormFields.Add Field

    'StampaDocEvasioneOrdNonAttivo
    Set Field = New FormField
    Set Field.Control = Me.chkStampaDocNonAtt
    Field.Name = "StampaDocEvasioneOrdNonAttivo"
    Field.Visible = True
    Me.chkStampaDocNonAtt.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDProvinciaIntra
    Set Field = New FormField
    Set Field.Control = Me.cboIntraProvincia
    Field.Name = "IDProvinciaIntra"
    Field.Visible = True
    Me.cboIntraProvincia.Tag = Field.Name
    m_FormFields.Add Field

    'VisualizzaListaOrdineNuovaPedana
    Set Field = New FormField
    Set Field.Control = Me.chkVisOrdNewPed
    Field.Name = "VisualizzaListaOrdineNuovaPedana"
    Field.Visible = True
    Me.chkVisOrdNewPed.Tag = Field.Name
    m_FormFields.Add Field

    'StampaEtichettaPDF
    Set Field = New FormField
    Set Field.Control = Me.chkStampaEtiPDF
    Field.Name = "StampaEtichettaPDF"
    Field.Visible = True
    Me.chkStampaEtiPDF.Tag = Field.Name
    m_FormFields.Add Field

    'PercorsoEtichettePDF
    Set Field = New FormField
    Set Field.Control = Me.txtPercorsoEtiPDF
    Field.Name = "PercorsoEtichettePDF"
    Field.Visible = True
    Me.txtPercorsoEtiPDF.Tag = Field.Name
    m_FormFields.Add Field

    'NonRiportareAccontiInFatturaPDC
    Set Field = New FormField
    Set Field.Control = Me.chkNonRipAcconti
    Field.Name = "NonRiportareAccontiInFatturaPDC"
    Field.Visible = True
    Me.chkNonRipAcconti.Tag = Field.Name
    m_FormFields.Add Field
    
    'chkRaggrLiqSocio
    Set Field = New FormField
    Set Field.Control = Me.chkRaggrLiqSocio
    Field.Name = "RaggruppaLiqPerSocio"
    Field.Visible = True
    Me.chkRaggrLiqSocio.Tag = Field.Name
    m_FormFields.Add Field
    
    'chkLottoProdPerSocio
    Set Field = New FormField
    Set Field.Control = Me.chkLottoProdPerSocio
    Field.Name = "CodiceLottoDiProduzionePerSocio"
    Field.Visible = True
    Me.chkLottoProdPerSocio.Tag = Field.Name
    m_FormFields.Add Field

    'Categoria anagrafica per identificare una cooperativa
    Set Field = New FormField
    Set Field.Control = Me.cboCategoriaAnagraficaCoop
    Field.Name = "IDCategoriaAnagraficaCoop"
    Field.Visible = True
    Me.cboCategoriaAnagraficaCoop.Tag = Field.Name
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
    


'**************************SOTTO DOCUMENTO DEI PROCESSI PER DOCUMENTI*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POStoriaRateContratto"
    
    Set m_DocumentsLink1 = m_Document.AddDocumentsLink("RV_POProcessiDocumentoCoop")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink1.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink1.PrimaryKey = "IDRV_POProcessiDocumentoCoop" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sul campo "RV_PODocumentoCoop.IDRV_PODocumentoCoop"
    Set NewLink = m_DocumentsLink1.AddLink("IDDocumentoCoop", "RV_PODocumentiCoop", ltLeft, "IDRV_PODocumentiCoop")
    NewLink.AddLinkColumn "RV_PODocumentiCoop.DocumentoCoop"
    
    'Crea un Link LEFT JOIN sul campo "RV_POIDTipoProcesso"
    Set NewLink = m_DocumentsLink1.AddLink("IDTipoProcessoCoop", "RV_POTipoProcessoCoop", ltLeft, "IDRV_POTipoProcessoCoop")
    NewLink.AddLinkColumn "RV_POTipoProcessoCoop.TipoProcessoCoop"
    
    'Crea un Link LEFT JOIN sul campo "Magazzino"
    'Set NewLink = m_DocumentsLink1.AddLink("IDMagazzino", "Magazzino", ltLeft, "IDMagazzino")
    'NewLink.AddLinkColumn "Magazzino.Magazzino"
    
    'Crea un Link LEFT JOIN sul campo "TipoMagazzino.Magazzino"
    Set NewLink = m_DocumentsLink1.AddLink("IDTipoMagazzino", "RV_POTipoMagazzino", ltLeft, "IDRV_POTipoMagazzino")
    NewLink.AddLinkColumn "RV_POTipoMagazzino.TipoMagazzino"
    
    'Crea un Link LEFT JOIN sul campo "Funzione.Funzione"
    Set NewLink = m_DocumentsLink1.AddLink("IDFunzione", "Funzione", ltLeft, "IDFunzione")
    NewLink.AddLinkColumn "Funzione.Funzione"
    
    
'************************************************************************************
'**************************SOTTO DOCUMENTO SEZIONALE PER DOCUMENTO COOP*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POStoriaRateContratto"
    
    Set m_DocumentsLink2 = m_Document.AddDocumentsLink("RV_POSezionalePerDocumento")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink2.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink2.PrimaryKey = "IDRV_POSezionalePerDocumento" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sul campo "Sezionale.Sezionale"
    Set NewLink = m_DocumentsLink2.AddLink("IDSezionale", "Sezionale", ltLeft, "IDSezionale")
    NewLink.AddLinkColumn "Sezionale.Sezionale"
    
    'Crea un Link LEFT JOIN sul campo "RV_PODocumentiCoop.DocumentoCoop"
    Set NewLink = m_DocumentsLink2.AddLink("IDDocumentoCoop", "RV_PODocumentiCoop", ltLeft, "IDRV_PODocumentiCoop")
    NewLink.AddLinkColumn "RV_PODocumentiCoop.DocumentoCoop"
    
'************************************************************************************
    
'**************************SOTTO DOCUMENTO PROGRESSIVI PROTOCOLLO ICE*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POStoriaRateContratto"
    
    Set m_DocumentsLink3 = m_Document.AddDocumentsLink("RV_POProgProtocolloICE")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink3.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink3.PrimaryKey = "IDRV_POProgProtocolloICE" '<-- Specifica il campo chiave primaria
    
    'Crea un Link LEFT JOIN sul campo "RV_POProtocolloICE.ProtocolloICE"
    Set NewLink = m_DocumentsLink3.AddLink("IDRV_POProtocolloICE", "RV_POProtocolloICE", ltLeft, "IDRV_POProtocolloICE")
    NewLink.AddLinkColumn "RV_POProtocolloICE.ProtocolloICE"
    
    'Crea un Link LEFT JOIN sul campo "Eserczio.Esercizio"
    Set NewLink = m_DocumentsLink3.AddLink("IDEsercizio", "Esercizio", ltLeft, "IDEsercizio")
    NewLink.AddLinkColumn "Esercizio.Esercizio"
    
'************************************************************************************
    'rif6 end
'**************************SOTTO DOCUMENTO DELLA QUADRATURA*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    Set m_DocumentsLink4 = m_Document.AddDocumentsLink("RV_POOperazionePerDoc")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink4.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink4.PrimaryKey = "IDRV_POOperazionePerDoc" '<-- Specifica il campo chiave primaria
   
    'Crea un Link LEFT JOIN sul campo "Funzione.Funzione"
    Set NewLink = m_DocumentsLink4.AddLink("IDRV_PODocumentoCoop", "RV_PODocumentiCoop", ltLeft, "IDRV_PODocumentiCoop")
    NewLink.AddLinkColumn "RV_PODocumentiCoop.DocumentoCoop"
    

    
'************************************************************************************

'**************************SOTTO DOCUMENTO INTERESSI ANTICIPAZIONI***************************************

    'Crea un sottodocumento basato sulla tabella di cross "IDRV_POTassoInteresseAnnuo"
    
    Set m_DocumentsLink6 = m_Document.AddDocumentsLink("RV_POTassoInteresseAnnuo")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink6.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink6.PrimaryKey = "IDRV_POTassoInteresseAnnuo" '<-- Specifica il campo chiave primaria
   

'********************************************************************************************************
'**************************SOTTO DOCUMENTO PARAMETRI UTENTE PER EVASIONE ORDINI*******************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POSchemaCoopQuadratura"
    
    Set m_DocumentsLink7 = m_Document.AddDocumentsLink("RV_POParametriUtenteOrd")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink7.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink7.PrimaryKey = "IDRV_POParametriUtenteOrd" '<-- Specifica il campo chiave primaria
   
    'Crea un Link LEFT JOIN sul campo "Utente"
    Set NewLink = m_DocumentsLink7.AddLink("IDUtente", "Utente", ltLeft, "IDUtente")
    NewLink.AddLinkColumn "Utente.Utente"
    

'************************************************************************************

End Sub




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
    
    

    'Connessione di tipo DMTADODBLib
        ConnessioneDiamanteADO
    
    'Inizializzazioni da fare prima dell'apertura del documento
        OnBeforeOpenDoc
    
        ParametroTipoCaloPeso
        ParametroTipoAumentoPeso
        ParametroTipoScarto
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
        m_DocType.Fields("ID" & m_App.TableName).Value = m_App.CallerFieldValue
        
        'Rimuove il filtro precedente
        m_DocType.RemoveFilter "Temp"
        
        'Crea un nuovo filtro temporaneo a partire dalle condizioni di ricerca
        'e viene reso filtro attivo
        Set m_ActiveFilter = m_DocType.AddFilterWithConditions("Temp")
        
        'Inidica, nel caso di esegui gestione, se riportare il valore corrente al chiamante
        bNotReturnValue = CBool(Val(GetSetting(REGISTRY_KEY, App.EXEName, "NoReturnValue", "0")))
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
    End If
    
        
    'Si comunica al documento quale filtro eseguire all'avvio.
    Set m_Document.ActiveFilter = m_ActiveFilter
    m_Document.Dataset.Recordset.Sort = CAMPO_PER_CAPTION
    Set Me.BrwMain.Recordset = m_Document.Dataset.Recordset
    'Prima di aprire il documento occorre comunicargli qual'è il campo chiave primaria.
    m_Document.PrimaryKey = "ID" & m_Document.TableName
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
    With DmtPrnDlg
        Set .Application = m_App
        Set .DocType = m_DocType
    End With
    
    
    
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
    
    'Vengono creati automaticamente i campi per la ricerca.
    'In una applicazione specifica questo codice andrà sostituito integralmente per definire
    'dei campi di ricerca ad hoc.
    
    'Non viene visualizzata la Check Intervallo perchè attualmente
    'il modello ad oggetti non prevede la gestione di filtri con
    'clausole BETWEEN.
    
    BrwMain.Conditions.Clear
   
    'For Each Field In m_DocType.Fields
        'With Field
            'Vengono esclusi dai filtri i campi ID
        '    If Left(.Name, 2) <> "ID" Then
                Set Cond = BrwMain.Conditions.Add("Filiale", "Filiale", m_DocType.TableName, False, False, False, dgCondTypeText)
                'Non viene visualizzata la Check Intervallo
                Cond.RangeEnabled = False
        '    End If
        'End With
    'Next
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
    m_Report.Copies = 1
    m_Report.Orientation = ocPortrait
    m_Report.PrinterName = ""
    
    If ToolName = "Mnu_Print" Then
        '**+ stampa con dialogo
        Set DmtPrnDlg.Report = m_Report
        DmtPrnDlg.Show
        If Not DmtPrnDlg.Cancel Then
            Screen.MousePointer = vbHourglass
            m_Document.DoPrint m_Report
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
        
            'ATTENZIONE: La stringa sMessage1 deve essere personalizzata a seconda dei casi!!!
            sMessage1 = " il " & m_DocType.Name
            sMessage = sMessage1 & " """ & m_Document.Fields(CAMPO_PER_CAPTION).Value & """"
            
            gResource.CustomStrings.Clear
            gResource.CustomStrings.Add sMessage, 1
                              
            'Viene chiesto se si intende riportare il record corrente al programma chiamante.
            If fnMsgQuestion(gResource.GetCustomizedMessage(MESS_QUERYPASTE), m_App.FunctionName) = vbYes Then
                'Legge l'ID del record corrente affinchè venga riportato all'applicazione chiamante.
                lIDField = m_Document.Fields("ID" & m_App.TableName).Value
            End If
            
        End If
        
        'Scrive sul registry l'ID da passare all'aplicazione chiamante.
        SaveSetting "Diamante", m_App.Caller, "IDField", lIDField
                                
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
        
        BrwMain.Top = PicForm2.ScaleTop
        BrwMain.Left = PicForm2.ScaleLeft
        BrwMain.Width = PicForm2.ScaleWidth
        BrwMain.Height = PicForm2.ScaleHeight
        

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
    Dim i As Integer

    '**+ Salva le impostazioni relative alle toolbar
    With AppOptions
        
        For i = 0 To BarMenu.Bands.Count - 1
            If BarMenu.Bands(i).Type <> ddBTPopup Then
                    .ToolbarDockingArea(i) = BarMenu.Bands(i).DockingArea
                    .ToolbarDockLine(i) = BarMenu.Bands(i).DockLine
                    .ToolbarLeft(i) = BarMenu.Bands(i).Left
                    .ToolbarTop(i) = BarMenu.Bands(i).Top
                    .ToolbarHeight(i) = BarMenu.Bands(i).Height
                    .ToolbarWidth(i) = BarMenu.Bands(i).Width
                    .ToolbarDockingOffset(i) = BarMenu.Bands(i).DockingOffset
            End If
        Next i
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
            BrwMain.SetFocus

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
        BrwMain.SetFocus
        
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
        gResource.CustomStrings.Add Chr(34) & TheApp.FunctionName & Chr(34), 1

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
    
    m_DocType.Fields("IDUtente").Value = 0
    m_DocType.Fields("IDAzienda").Value = m_App.IDFirm
    
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
            Case dgCondTypeText, dgCondTypeNumber, dgCondTypeDate, dgCondTypeTime
                m_DocType.Fields(Cond.FieldName).Value = Cond.FromValue
                
          
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
    Dim cl As DmtGridCtl.dgColumnHeader

    'Inizializzazione della griglia adibita alla visualizzazione tabellare dei sotto-documenti
    '-------------------------------------------------------------------------------
       
       
   
    
    
    If Me.GrigliaProcessi.ColumnsHeader.Count = 0 Then
        With Me.GrigliaProcessi.ColumnsHeader
            .Add "DocumentoCoop", "Documento", dgchar, True, 2000, 0, True, True, False
            .Add "TipoProcessoCoop", "Tipo processo", dgchar, True, 1500, 0, True, True, False
            '.Add "Magazzino", "Magazzino", dgchar, True, 1500, 0, True, True, False
            .Add "TipoMagazzino", "Tipo magazzino", dgchar, True, 1500, 0, True, True, False
            .Add "Funzione", "Causale", dgchar, True, 1500, dgAlignleft, True, True, False
        End With
    End If
    Me.GrigliaProcessi.EnableMove = True

    If Me.GrigliaSezionale.ColumnsHeader.Count = 0 Then
        With Me.GrigliaSezionale.ColumnsHeader
            .Add "Sezionale", "Sezionale", dgchar, True, 2000, 0, True, True, False
            .Add "DocumentoCoop", "Documento", dgchar, True, 1500, 0, True, True, False
            .Add "Predefinito", "Predefinito", dgBoolean, True, 1000, 0, True, True, False
        End With
    End If
    Me.GrigliaSezionale.EnableMove = True
    
    If Me.GrigliaProgressivoICE.ColumnsHeader.Count = 0 Then
        With Me.GrigliaProgressivoICE.ColumnsHeader
            .Add "ProtocolloICE", "Protocollo ICE", dgchar, True, 2000, 0, True, True, False
            .Add "Progressivo", "Numero", dgNumeric, True, 1000, 0, True, True, False
            .Add "Esercizio", "Eserizio", dgchar, True, 1000, 0, True, True, False
            .Add "Predefinito", "Predefinito", dgBoolean, True, 1000, 0, True, True, False
        End With
    End If
    Me.GrigliaProgressivoICE.EnableMove = True
    
    If Me.GrigliaOperazione.ColumnsHeader.Count = 0 Then
        With Me.GrigliaOperazione.ColumnsHeader
            .Add "DocumentoCoop", "Documento", dgchar, True, 2000, 0, True, True, False
            .Add "GestioneArticoli", "Gestione art.", dgBoolean, True, 1500, 0, True, True, False
            .Add "CreazioneAutomaticaLottoVend", "Creazione aut. lotto vend.", dgBoolean, True, 2000, 0, True, True, False
        End With
    End If
    Me.GrigliaOperazione.EnableMove = True
    
    If Me.GrigliaInteressi.ColumnsHeader.Count = 0 Then
        With Me.GrigliaInteressi.ColumnsHeader
            .Add "IDRV_POTassoInteresseAnnuo", "IDRV_POTassoInteresseAnnuo", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POSchemaCoop", "IDRV_POSchemaCoop", dgInteger, False, 500, 0, True, True, False
            .Add "IDAzienda", "IDAzienda", dgInteger, False, 500, 0, True, True, False
            .Add "IDFiliale", "IDFiliale", dgInteger, False, 500, 0, True, True, False
            .Add "DataInizio", "Data inizio", dgDate, True, 1500, 0, True, True, False
            Set cl = .Add("TassoInteresse", "% Annuale", dgDouble, True, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("NumeroGiorniAnno", "N° Giorni anno", dgDouble, True, 2000, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("TassoGiornaliero", "% Giorno", dgDouble, True, 1800, dgAlignRight, True, True, False)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
        End With
    End If
    Me.GrigliaInteressi.EnableMove = True

    If Me.GrigliaUtenteEva.ColumnsHeader.Count = 0 Then
        With Me.GrigliaUtenteEva.ColumnsHeader
            .Add "IDRV_POParametriUtenteOrd", "IDRV_POParametriUtenteOrd", dgInteger, False, 500, 0, True, True, False
            .Add "IDRV_POSchemaCoop", "IDRV_POSchemaCoop", dgInteger, False, 500, 0, True, True, False
            .Add "IDAzienda", "IDAzienda", dgInteger, False, 500, 0, True, True, False
            .Add "IDFiliale", "IDFiliale", dgInteger, False, 500, 0, True, True, False
            .Add "IDUtente", "IDUtente", dgInteger, False, 500, 0, True, True, False
            .Add "Utente", "Utente", dgVarChar, True, 2000, 0, True, True, False
            .Add "IDOggettoSmist", "IDOggettoSmist", dgInteger, False, 500, 0, True, True, False
            .Add "IDClienteSmist", "IDClienteSmist", dgInteger, False, 500, 0, True, True, False
            .Add "NumeroOrdineSmist", "N° ord. smist.", dgVarChar, True, 1500, dgAlignRight, True, True, False
            .Add "DataOrdineSmist", "Data ord. smist.", dgDate, True, 1500, 0, True, True, False
            .Add "IDOggettoPrep", "IDOggettoPrep", dgInteger, False, 500, 0, True, True, False
            .Add "IDClientePrep", "IDClientePrep", dgInteger, False, 500, 0, True, True, False
            .Add "NumeroOrdinePrep", "N° ord. prep.", dgVarChar, True, 1500, dgAlignRight, True, True, False
            .Add "DataOrdinePrep", "Data ord. prep.", dgDate, True, 1500, 0, True, True, False
            .Add "BloccaOrdinePrep", "Blocca ordine", dgBoolean, True, 1500, dgAligncenter, True, True, False
        End With
    End If
    Me.GrigliaUtenteEva.EnableMove = True

'''''''''''''''''''''''''CONTROLLI STANDARD''''''''''''''''''''''''''''''''''''
    
     
    
    'Filiale
    With Me.cboFiliale
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFiliale"
        .DisplayField = "Filiale"
        .SQL = "SELECT Filiale.IDFiliale, Filiale.Filiale "
        .SQL = .SQL & "FROM AttivitaAzienda INNER JOIN "
        .SQL = .SQL & "Azienda ON AttivitaAzienda.IDAzienda = Azienda.IDAzienda INNER JOIN "
        .SQL = .SQL & "Filiale ON AttivitaAzienda.IDAttivitaAzienda = Filiale.IDAttivitaAzienda "
        .SQL = .SQL & "WHERE (Azienda.IDAzienda =" & m_App.IDFirm & ")"
        .Fill
    End With
    
    'Magazzino di carico
    With Me.cboMagazzinoCarico
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .SQL = "SELECT * FROM Magazzino  "
        .SQL = .SQL & "WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With
    
    'Magazzino di vendita
    With Me.cboMagazzinoVendita
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDMagazzino"
        .DisplayField = "Magazzino"
        .SQL = "SELECT * FROM Magazzino  "
        .SQL = .SQL & "WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With
    

    With Me.cboCausaleCarico_Car
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With

    With Me.cboCausaleCarico_Vend
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With
    
    With Me.cboCausaleScarico_Car
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With

    With Me.cboCausaleScarico_Vend
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With
    
    With Me.cboCausaliMagazzino
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With
    
    With Me.cboDocumentoCoop
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PODocumentiCoop"
        .DisplayField = "DocumentoCoop"
        .SQL = "SELECT * FROM RV_PODocumentiCoop ORDER BY DocumentoCoop"
        .Fill
    End With
    
    With Me.cboDocCoopPerOpe
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PODocumentiCoop"
        .DisplayField = "DocumentoCoop"
        .SQL = "SELECT * FROM RV_PODocumentiCoop ORDER BY DocumentoCoop"
        .Fill
    End With

    With Me.cboTipoProcesso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoProcessoCoop"
        .DisplayField = "TipoProcessoCoop"
        .SQL = "SELECT * FROM RV_POTipoProcessoCoop"
        .Fill
    End With

    With Me.cboSezionale
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT * FROM Sezionale WHERE IDFiliale=" & m_App.Branch
        .Fill
    End With

    With Me.cboDocumenti_PerSezionale
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_PODocumentiCoop"
        .DisplayField = "DocumentoCoop"
        .SQL = "SELECT * FROM RV_PODocumentiCoop WHERE IDRV_PODocumentiCoop<10"
        .Fill
    End With
    
    With Me.cboEsercizioICE
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDEsercizio"
        .DisplayField = "Esercizio"
        .SQL = "SELECT * FROM Esercizio WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With
    
    With Me.cboProtocolloICE
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POProtocolloICE"
        .DisplayField = "ProtocolloICE"
        .SQL = "SELECT * FROM RV_POProtocolloICE"
        .Fill
    End With
    
    With Me.cboTipoMagazzino
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoMagazzino"
        .DisplayField = "TipoMagazzino"
        .SQL = "SELECT * FROM RV_POTipoMagazzino"
        .Fill
    End With

    With Me.cboTipoProdottoImballo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDTipoProdotto"
        .DisplayField = "TipoProdotto"
        .SQL = "SELECT * FROM TipoProdotto WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With

    With Me.cboTipoProdottoGrezzo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDTipoProdotto"
        .DisplayField = "TipoProdotto"
        .SQL = "SELECT * FROM TipoProdotto WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With

    With Me.cboTipoProdottoLavorato
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDTipoProdotto"
        .DisplayField = "TipoProdotto"
        .SQL = "SELECT * FROM TipoProdotto WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With

    With Me.cboCategoriaAnagraficaSocio
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDCategoriaAnagrafica"
        .DisplayField = "CategoriaAnagrafica"
        .SQL = "SELECT * FROM CategoriaAnagrafica WHERE IDTipoAnagrafica IS NULL"
        .Fill
    End With

    With Me.cboTipoQuadratura
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoQuadratura"
        .DisplayField = "TipoQuadratura"
        .SQL = "SELECT * FROM RV_POTipoQuadratura"
        .Fill
    End With
    
    With Me.cboTipoCaloPeso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDTipoProdotto"
        .DisplayField = "TipoProdotto"
        .SQL = "SELECT * FROM TipoProdotto WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With
    
    With Me.cboTipoAumentoPeso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDTipoProdotto"
        .DisplayField = "TipoProdotto"
        .SQL = "SELECT * FROM TipoProdotto WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With

    With Me.cboTipoScarto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDTipoProdotto"
        .DisplayField = "TipoProdotto"
        .SQL = "SELECT * FROM TipoProdotto WHERE IDAzienda=" & m_App.IDFirm
        .Fill
    End With

    With Me.cboCausaleTipoScarto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With

    With Me.cboCausaleScaricoTipoScarto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With

    With Me.cboCausaleCaloPeso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With

    With Me.cboCausaleScaricoCaloPeso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With


    With Me.cboCausaleAumentoPeso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With

    With Me.cboCausaleScaricoAumentoPeso
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=9"
        .Fill
    End With


    With Me.cboTipoLavorazioneAut
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .SQL = "SELECT * FROM RV_POTipoLavorazione"
        .Fill
    End With

    With Me.cboIvaPerFatturazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT * FROM Iva"
        .Fill
    End With
    With Me.cboPagamentoPerFatturazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDPagamento"
        .DisplayField = "Pagamento"
        .SQL = "SELECT * FROM Pagamento"
        .Fill
    End With
    With Me.cboValutaPerFatturazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDValuta"
        .DisplayField = "Valuta"
        .SQL = "SELECT * FROM Valuta"
        .Fill
    End With

    With Me.cboCausContPerFatturazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDCausaleContabile"
        .DisplayField = "CausaleContabile"
        .SQL = "SELECT * FROM CausaleContabile WHERE IDRegistroIVA=2 AND IDTipoOggetto=122"
        .Fill
    End With

    'Listino
    With Me.cboListinoImballiDefault
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT * FROM Listino WHERE ("
        .SQL = .SQL & "(IDAzienda=" & m_App.IDFirm & ") AND "
        .SQL = .SQL & "(TipoListino=0))"
        .Fill
    End With

'    With Me.cboTipoGestioneArticoliVend
'        Set .Database = m_App.Database.Connection
'        .AddFieldKey "IDRV_POTipoGestioneArticoliVendita"
'        .DisplayField = "TipoGestioneArticoliVendita"
'        .SQL = "SELECT * FROM RV_POTipoGestioneArticoliVendita"
'        .Fill
'    End With

    With Me.CDTipoPedana
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceTipoPedana"
        .DescriptionField = "TipoPedana"
        .KeyField = "IDRV_POTipoPedana"
        .TableName = "RV_POTipoPedana"
        .Filter = "IDAzienda = " & m_App.IDFirm & " AND IDFiliale = " & m_App.Branch
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        'Caption da associare alla label del campo Descrizione
        .PropDescrizione.Caption = "Descrizione"
        'Caption da associare alla intestazione della colonna della Find per il campo Codice
        .CodeCaption4Find = "Codice "
        'Caption da associare alla intestazione della colonna della Find per il campo Descrizione
        .DescriptionCaption4Find = "Descrizione "
        'Identificativo della Funzione Diamante per l'Esegui Gestione
        .IDExecuteFunction = fncTrovaIDFunzione("RV_POTipoPedana")
        'Indica se il campo Codice è un campo numerico
        .CodeIsNumeric = False
    End With


    With Me.cboTipoArrotondamento
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoArrotondamento"
        .DisplayField = "TipoArrontondamento"
        .SQL = "SELECT * FROM RV_POTipoArrotondamento"
        .Fill
    End With

    With Me.cboTipoArrotondamentoConf
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoArrotondamento"
        .DisplayField = "TipoArrontondamento"
        .SQL = "SELECT * FROM RV_POTipoArrotondamento"
        .Fill
    End With

    With Me.cboTipoCorpoFattSocio
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoCorpoFatturaSocio"
        .DisplayField = "TipoCorpoFatturaSocio"
        .SQL = "SELECT * FROM RV_POTipoCorpoFatturaSocio"
        .Fill
    End With

    With Me.cboLingua
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDLinguaDescrizioneArticolo"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "LinguaDescrizioneArticolo"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM LinguaDescrizioneArticolo"
        .SQL = .SQL & " ORDER BY LinguaDescrizioneArticolo"
    End With


    With Me.cboTipoSceltaArticoloConferito
        'Imposta la connessione corrente al database DMT
        '(passa la connessione di tipo dmtoledblib.AdoConnection)
        Set .Database = TheApp.Database.Connection
        'Indica il campo chiave univoco di accesso al record
        .AddFieldKey "IDRV_POTipoSceltaArticoloLottoCampagna"
        'Indica il campo da visualizzare nella combo
        .DisplayField = "TipoSceltaArticoloLottoCampagna"
        'Indica la query SQL da utilizzare per il reperimento dei dati
        .SQL = "SELECT * FROM RV_POTipoSceltaArticoloLottoCampagna"
        .SQL = .SQL & " ORDER BY TipoSceltaArticoloLottoCampagna"
    End With

    With Me.cboTipoComportamentoLavorazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoComportamentoLavorazione"
        .DisplayField = "TipoComportamentoLavorazione"
        .SQL = "SELECT * FROM RV_POTipoComportamentoLavorazione"
        .Fill
    End With

    With Me.cboTipoPesoArticolo
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoPesoArticolo"
        .DisplayField = "TipoPesoArticolo"
        .SQL = "SELECT * FROM RV_POTipoPesoArticolo"
        .Fill
    End With


    'Cliente per ordine in giacenza
     With Me.CDClienteOrdinePred
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
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

    'Cliente per ordine IV Gamma
     With Me.CDClienteOrdIVGamma
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
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
    
    'Cliente per ordine IV Gamma da lavorazione
     With Me.CdClienteIVGammaLav
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "Anagrafica"
        .KeyField = "IDAnagrafica"
        .TableName = "IERepCliente"
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
    
    With Me.cboUMRigaOrdine
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoUMOrdine"
        .DisplayField = "TipoUMOrdine"
        .SQL = "SELECT * FROM RV_POTipoUMOrdine"
        .SQL = .SQL & " ORDER BY TipoUMOrdine"
    End With

    With Me.cboUtenteOrd
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDUtente"
        .DisplayField = "Utente"
        .SQL = "SELECT * FROM Utente"
        .SQL = .SQL & " ORDER BY Utente"
    End With


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


     With Me.CDClienteSmist
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

    With Me.cboSezionaleCMR
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT * FROM Sezionale WHERE IDFiliale=" & m_App.Branch
        .Fill
    End With

    With Me.CDNaturaTransazione
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "NaturaTransazione"
        .KeyField = "IDNaturaTransazione"
        .TableName = "NaturaTransazione"
        '.Filter = "IDAzienda = " & m_App.IDFirm & " AND IDFiliale = " & m_App.Branch
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Natura di transazione"
        .CodeCaption4Find = "Codice "
        .DescriptionCaption4Find = "Natura di transazione"
        '.IDExecuteFunction = fncTrovaIDFunzione("RV_POTipoPedana")
        .CodeIsNumeric = False
    End With

    With Me.CDModoTrasporto
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "Codice"
        .DescriptionField = "ModoDiTrasporto"
        .KeyField = "IDModoDiTrasporto"
        .TableName = "ModoDiTrasporto"
        '.Filter = "IDAzienda = " & m_App.IDFirm & " AND IDFiliale = " & m_App.Branch
        .MenuFunctions("EseguiGestione").Enabled = False
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Modo di trasporto"
        .CodeCaption4Find = "Codice "
        .DescriptionCaption4Find = "Modo di trasporto "
        '.IDExecuteFunction = fncTrovaIDFunzione("RV_POTipoPedana")
        .CodeIsNumeric = False
    End With
    
    With Me.cboListinoCampionatura
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT * FROM Listino WHERE ("
        .SQL = .SQL & "(IDAzienda=" & TheApp.IDFirm & ") AND "
        .SQL = .SQL & "(TipoListino=1))"
        .Fill
    End With

    With Me.cboDecimaliConf
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoDecimaliPesi"
        .DisplayField = "TipoDecimaliPesi"
        .SQL = "SELECT * FROM RV_POTipoDecimaliPesi ORDER BY IDRV_POTipoDecimaliPesi"
        .Fill
    End With

    With Me.cboDecimaliLav
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoDecimaliPesi"
        .DisplayField = "TipoDecimaliPesi"
        .SQL = "SELECT * FROM RV_POTipoDecimaliPesi ORDER BY IDRV_POTipoDecimaliPesi"
        .Fill
    End With

    With Me.cboDecimaliVend
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoDecimaliPesi"
        .DisplayField = "TipoDecimaliPesi"
        .SQL = "SELECT * FROM RV_POTipoDecimaliPesi ORDER BY IDRV_POTipoDecimaliPesi"
        .Fill
    End With


    With Me.cboTipoSpezzatura
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoSpaccaturaLavorazione"
        .DisplayField = "TipoSpaccaturaLavorazione"
        .SQL = "SELECT * FROM RV_POTipoSpaccaturaLavorazione"
        .Fill
    End With
    
    
    sSQL = ""
    
    With Me.cboFunzioneDDTAcq
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=302"
        .Fill
    End With

    With Me.cboFunzioneFAAcq
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDFunzione"
        .DisplayField = "Funzione"
        .SQL = "SELECT * FROM Funzione WHERE IDTipoOggetto=301"
        .Fill
    End With


     With Me.CDArtPesPedNeg
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND ((IDTipoProdotto = " & Link_TipoScarto & ") OR (IDTipoProdotto = " & Link_TipoCaloPeso & "))"
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzioneArticoli("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With

     With Me.CDArtPesPedPos
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm & " AND IDTipoProdotto = " & Link_TipoAumentoPeso
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = fncTrovaIDFunzioneArticoli("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
    
    With Me.cboCatCommTrasporto
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDRV_POTipoCategoriaComm"
        .DisplayField = "TipoCategoriaComm"
        .SQL = "SELECT * FROM RV_POTipoCategoriaComm"
    End With

    'Articolo per commissione in conferimento (Gestione Vivaio)
    With Me.CDArticoloCommConf
        Set .Application = m_App
        Set .Database = m_App.Database
        .HwndContainer = Me.hwnd
        .CodeField = "CodiceArticolo"
        .DescriptionField = "Articolo"
        .KeyField = "IDArticolo"
        .TableName = "Articolo"
        .Filter = "VirtualDelete = 0 AND IDAzienda = " & m_App.IDFirm '& " AND IDTipoProdotto=" & Link_TipoImballo
        .MenuFunctions("EseguiGestione").Enabled = True
        .PropCodice.Caption = "Codice"
        .PropDescrizione.Caption = "Descrizione"
        .CodeCaption4Find = "Codice Articolo"
        .DescriptionCaption4Find = "Descrizione Articolo"
        .IDExecuteFunction = 6 'fncTrovaIDFunzione("Articoli", "Greentop - Articoli") 'Articoli
        .CodeIsNumeric = False
    End With
    
    With Me.cboTipoTrattAggVivaio
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoTrattenutaAggiuntiva"
        .DisplayField = "TipoTrattenuta"
        .SQL = "SELECT * FROM RV_POTipoTrattenutaAggiuntiva"
        .Fill
    End With
    
    With Me.cboPorto
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDPorto"
        .DisplayField = "Porto"
        .SQL = "SELECT * FROM Porto"
        .Fill
    End With

    With Me.cboSezListaPrelievo
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSezionale"
        .DisplayField = "Sezionale"
        .SQL = "SELECT  Sezionale.IDSezionale, Sezionale.Sezionale, RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "FROM Sezionale INNER JOIN "
        .SQL = .SQL & "RegistroIvaPerTipoOggetto ON Sezionale.IDRegistroIva = RegistroIvaPerTipoOggetto.IDRegistroIva AND "
        .SQL = .SQL & "Sezionale.IDFiliale = RegistroIvaPerTipoOggetto.IDFiliale "
        .SQL = .SQL & "WHERE RegistroIvaPerTipoOggetto.IDTipoOggetto = " & 15
        .SQL = .SQL & " AND RegistroIvaPerTipoOggetto.IDFiliale = " & TheApp.Branch
    End With
    
    With Me.cboIntraProvincia
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDProvincia"
        .DisplayField = "NomeProvincia"
        .SQL = "SELECT * FROM Provincia ORDER BY Provincia"
    End With
    
    With Me.cboCategoriaAnagraficaCoop
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDCategoriaAnagrafica"
        .DisplayField = "CategoriaAnagrafica"
        .SQL = "SELECT * FROM CategoriaAnagrafica"
        .Fill
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
    Dim Field As DmtDocManLib.Field
    Dim DocLink As DmtDocManLib.DocumentsLink
    Dim NuovoContratto As Boolean
    Dim NuovaRateizzazione As Boolean
    
    
    If Me.cboFiliale.CurrentID = 0 Then
        MsgBox "Bisogna inserire la filiale", vbInformation, "Parametri filiale"
        Exit Sub
    End If
    If m_Document("IDRV_POSchemaCoop").Value <= 0 Then
        If EsistenzaFiliale = True Then
            MsgBox "Per la filiale " & Me.cboFiliale.Text & " sono gia stati inseriti i parametri di default." & vbCrLf & "Impossibile salvare", vbInformation, "Parametri filiale"
            Exit Sub
        End If
    End If
    If Not PermissionToSave Then
        Exit Sub
    End If
        
    
    'Passa alla collezione Fields dell'oggetto
    'Document i valori da salvare
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
                        Field.Value = m_FormFields(Field.Name).Control.Value
                
                End Select
                
                'rif4 end
                
            Else
                If Field.Name = "IDAzienda" Then
                    Field.Value = m_App.IDFirm
                End If
                
                If Field.Name = "IDFiliale" Then
                    Field.Value = m_App.Branch
                End If
                If Field.Name = "IDUtente" Then
                    Field.Value = 0
                End If
                If Field.Name = "IDPDCFatturazione" Then
                    Field.Value = Link_ContoPDC
                End If
                If Field.Name = "IDPDCRigheLiqPositive" Then
                    Field.Value = Link_ContoPDC_RigaPos
                End If
                If Field.Name = "IDPDCRigheLiqNegative" Then
                    Field.Value = Link_ContoPDC_RigaNeg
                End If
                If Field.Name = "NonCalcImpDaAssVeloce" Then
                    Field.Value = PAR_NonCalcImpDaAssVeloce
                End If
                If Field.Name = "CalcImpAConfOrd" Then
                    Field.Value = PAR_CalcImpAConfOrd
                End If
                If Field.Name = "NonVisMsgImpZeroConfOrd" Then
                    Field.Value = PAR_NonVisMsgImpZeroConfOrd
                End If
                If Field.Name = "AssNewPedDaAssSingola" Then
                    Field.Value = PAR_AssNewPedDaAssSingola
                End If
                If Field.Name = "NonCalcPrezzoDueRefArtOrd" Then
                    Field.Value = PAR_NonCalcPrezzoDueRefArtOrd
                End If
                If Field.Name = "AvviaNewSelQGammaIn" Then
                    Field.Value = ATT_MULT_SEL_IV_GAMMA
                End If
                If Field.Name = "ChiudiConfQtaSelZeroNewSelQGammaIn" Then
                    Field.Value = CHIUDI_CONF_QTASEL_ZERO
                End If
                If Field.Name = "NumeroGiorniDataScadContratto" Then
                    Field.Value = GG_DATA_SCADENZA_CONTR
                End If
                If Field.Name = "IDRV_POFocusControlIVGammaLav" Then
                    Field.Value = FOCUS_IVGAMMA_LAV
                End If
                If Field.Name = "IDRV_POFocusControlIVGammaConf" Then
                    Field.Value = FOCUS_IVGAMMA_CONF
                End If
                If Field.Name = "IDUMCoopSelAutIVGamma" Then
                    Field.Value = LINK_UM_COOP_SEL_AUT_IVGAMMA
                End If
                If Field.Name = "DataScadenzaContrattoObbligatoria" Then
                    Field.Value = DATA_SCADENZA_OBBL
                End If
                If Field.Name = "IDSocioPerRiscontroPeso" Then
                    Field.Value = PAR_IDSOCIO_RISC_PESO
                End If
                If Field.Name = "VisualizzaNoteLavDaOrdineInElenco" Then
                    Field.Value = PAR_VIS_NOTE_RIGA_ORD_ELENCO
                End If
                If Field.Name = "Gtin" Then
                    Field.Value = GtinMigros
                End If
                If Field.Name = "UrlMigros" Then
                    Field.Value = UrlMigros
                End If
                If Field.Name = "NomeUtenteMigros" Then
                    Field.Value = NomeUtenteMigros
                End If
                If Field.Name = "PasswordMigros" Then
                    Field.Value = PasswordMigros
                End If
                If Field.Name = "UrlServizioFeedentity" Then
                    Field.Value = UrlFeed
                End If
                If Field.Name = "ChiaveClienteFeedentity" Then
                    Field.Value = ChiaveFeed
                End If
                If Field.Name = "AttivaServizioFeedentity" Then
                    Field.Value = AttivaFeed
                End If
                
                If Field.Name = "AttivaGestionePrelieviConferimento" Then
                    Field.Value = ATTIVA_PROD_CONF_LAV
                End If
                If Field.Name = "AttivaGestioneOrdiniInLavorazione" Then
                    Field.Value = ATTIVA_PROD_ORD_LAV
                End If
                If Field.Name = "DisabilitaListaTestaOrdine" Then
                    Field.Value = DISATTIVA_PROD_ORD_TESTA
                End If
                If Field.Name = "AttivaCalcoloPesoPedanaInProduzione" Then
                    Field.Value = ATTIVA_PROD_CALC_COLLI_PED
                End If
                If Field.Name = "SelezionaLottoCampagnaInLavorazione" Then
                    Field.Value = ATTIVA_SEL_SUB_LOTTO_PROD
                End If
                If Field.Name = "AttivaMetodoSubContattiFeedentity" Then
                    Field.Value = ATTIVA_SUB_CONTATTI_FEED
                End If
                If Field.Name = "AttivaMetodoSelMultLottoFeedentity" Then
                    Field.Value = ATTIVA_SEL_MULT_LOTTI_FEED
                End If
                If Field.Name = "AttivaControlloAccessoFifo" Then
                    Field.Value = ATTIVA_PROD_FIFO_ACCESSI_CONF
                End If
                If Field.Name = "CodiceViaggioCompletato" Then
                    Field.Value = CodiceConfermaViaggio
                End If
                If Field.Name = "CodiceRispristinoOperazioni" Then
                    Field.Value = CodiceAnnullaOperazione
                End If
                If Field.Name = "CodiceConfermaOrdine" Then
                    Field.Value = CodiceConfermaOrdine
                End If
                If Field.Name = "CodiceAnnullaOperazione" Then
                    Field.Value = CodiceGestioneErrori
                End If
                If Field.Name = "CreaListaPrelievoAutPedInUscita" Then
                    Field.Value = ATTIVA_GESTIONE_CARICO_MERCE
                End If
                If Field.Name = "NonVisMsgCreaListaPrelievoAutPedUscita" Then
                    Field.Value = NonVisMsgCreaListaPrelAutPedUscita
                End If
                If Field.Name = "VisualizzaAndamentoGrigliaRigaOdine" Then
                    Field.Value = ATTIVA_VisAndGrigliaRigaOdine
                End If
                If Field.Name = "VisualizzaAndamentoGrigliaConferimento" Then
                    Field.Value = ATTIVA_VisAndGrigliaConferimento
                End If
                If Field.Name = "IDRV_POTipoCollegamentoRigaOrdRigaConf" Then
                    Field.Value = IDTipoCollegRigaOrdRigaConf
                End If
                If Field.Name = "ChiusuraPrelievoAggQtaPrelConColliEntrati" Then
                    Field.Value = ATTIVA_RICALCOLO_COLLI_PREL_CHIUSO
                End If
                If Field.Name = "NonVisualizzaRigheOrdiniCompletateInSelLav" Then
                    Field.Value = NonVisRigheOrdiniComplInSelLav
                End If
                If Field.Name = "ModificaLavorazioneSenzaRicalcolo" Then
                    Field.Value = ModificaLavorazioneSenzaRicalcolo
                End If
                If Field.Name = "RiscontroPesoPerOgniRiga" Then
                    Field.Value = RISCONTRO_PESO_PER_CONF
                End If
                If Field.Name = "ConsentiEliminazioneAccessi" Then
                    Field.Value = ATTIVA_ELIMINAZIONE_ACCESSI
                End If
                If Field.Name = "IDAnagraficaSocioMovPreConferimento" Then
                    Field.Value = IDSOCIO_PRE_CONF
                End If
                If Field.Name = "AttivaCommissioniDaOrdine" Then
                    Field.Value = ATTIVA_COMMISSIONI_ORDINI
                End If
                If Field.Name = "RicalcolaCommTipoPedDaOrdInEvasione" Then
                    Field.Value = RIC_COMM_TIPO_PED_EVAS_ORD
                End If
                If Field.Name = "VisElencoRigheOrdineSeNonTroviAssociazione" Then
                    Field.Value = VIS_ELENCO_RIGHE_ORD
                End If
                If Field.Name = "AttivaCalcoloNPedEffInVenditeBI" Then
                    Field.Value = ATTIVA_CALCOLO_N_PED_BI_VENDITE
                End If
                If Field.Name = "RiscontroPesoAValorePerSocio" Then
                    Field.Value = RISCONTRO_PESO_VAL_SOCIO
                End If
                If Field.Name = "RiscontroPesoAValorePerFornitore" Then
                    Field.Value = RISCONTRO_PESO_VAL_FORNITORE
                End If
                If Field.Name = "ObbligatorioNDocInPesConf" Then
                    Field.Value = OBBL_N_DOC_PES_CONF
                End If
                
                If Field.Name = "ObbligatorioRifSfalcioConf" Then
                    Field.Value = ATTIVA_OBBL_SFALCIO_CONF
                End If
                If Field.Name = "ObbligatorioSequenzaSfalcioConf" Then
                    Field.Value = ATTIVA_SEQ_SFALCIO
                End If
                If Field.Name = "RIferimentiDettagliatiInNC" Then
                    Field.Value = RIPORTA_RIF_DETTAGLIO_XML_NC
                End If
                If Field.Name = "RIferimentiDettagliatiInND" Then
                    Field.Value = RIPORTA_RIF_DETTAGLIO_XML_ND
                End If
                If Field.Name = "ConfPresaVisAutOrdDaContr" Then
                    Field.Value = CONF_AUT_CONTR_PRESA_VISIONE
                End If
                If Field.Name = "NonRipAgenteDaEvOrdNoPres" Then
                    Field.Value = NO_RIP_AGENTE_IN_DOC_EVASIONE
                End If
                
                If Field.Name = "RiportaInXMLRifLetteraIntento" Then
                    Field.Value = Rip_InXMLRifLetteraIntento
                End If
                If Field.Name = "RiportaInXMLRifNoteIva" Then
                    Field.Value = Rip_InXMLRifNoteIva
                End If
                If Field.Name = "RiportaInXMLRifNota01Doc" Then
                    Field.Value = Rip_InXMLRifNota01Doc
                End If
                If Field.Name = "RiportaInXMLRifNota02Doc" Then
                    Field.Value = Rip_InXMLRifNota02Doc
                End If
                If Field.Name = "RiportaInXMLRifNota03Doc" Then
                    Field.Value = Rip_InXMLRifNota03Doc
                End If
                If Field.Name = "RiportaInXMLRifNotaDoc" Then
                    Field.Value = Rip_InXMLRifNotaDoc
                End If
                If Field.Name = "RiportaInXMLRifIstrMitt" Then
                    Field.Value = Rip_InXMLRifIstrMitt
                End If
                If Field.Name = "RiportaInXMLRifVettSucc" Then
                    Field.Value = Rip_InXMLRifVettSucc
                End If
                If Field.Name = "RiportaInXMLRifAgenziaTrasp" Then
                    Field.Value = Rip_InXMLRifAgenziaTrasp
                End If
                If Field.Name = "RiportaInXMLRifTargaAutoMezzo" Then
                    Field.Value = Rip_InXMLRifTargaAutoMezzo
                End If
                If Field.Name = "Alyante_CompanyCode" Then
                    Field.Value = COMPANY_CODE
                End If
                If Field.Name = "Alyante_AttivaFido" Then
                    Field.Value = ATTIVA_FIDO_ALY
                End If
                If Field.Name = "Alyante_DisattivaCalcoloDDT" Then
                    Field.Value = DISATTIVA_DDT_FIDO_ALY
                End If
                If Field.Name = "Alyante_DisattivaCalcoloFA" Then
                    Field.Value = DISATTIVA_FA_FIDO_ALY
                End If
                If Field.Name = "Alyante_DisattivaCalcoloFD" Then
                    Field.Value = DISATTIVA_FD_FIDO_ALY
                End If
                If Field.Name = "Alyante_DisattivaCalcoloNC" Then
                    Field.Value = DISATTIVA_NC_FIDO_ALY
                End If
                If Field.Name = "Alyante_DisattivaCalcoloND" Then
                    Field.Value = DISATTIVA_ND_FIDO_ALY
                End If
                If Field.Name = "DisattivaScalataCommTipoPedana" Then
                    Field.Value = DISATTIVA_SCALATA_COMM_TRASP
                End If
                If Field.Name = "AttivaMultilivelloFeedentity" Then
                    Field.Value = ATTIVA_FEED_MULTI_LIVELLO
                End If
                If Field.Name = "RV_POIDFeedentity" Then
                    Field.Value = IDFEED_AZIENDA
                End If
                If Field.Name = "NumeroColliPerAutomezzoCert" Then
                    Field.Value = NUMERO_COLLI_PRED_CERT
                End If
                If Field.Name = "SincronizzaAutFeed" Then
                    Field.Value = SincronizzaAutFeed
                End If
                If Field.Name = "IvaArticoloDaDocumentoCollegato" Then
                    Field.Value = IvaArticoloDaDocColl
                End If
                If Field.Name = "LetteraIntentoDaDocumentoCollegato" Then
                    Field.Value = LetteraIntentoDaDocColl
                End If
                If Field.Name = "RifComuneDaConfigurazioneSocio" Then
                    Field.Value = RifComuneDaConfigSocio
                End If
                If Field.Name = "DocAnaDestUgualeAnaCoop" Then
                    Field.Value = DocAnaDestUgualeAnaCoop
                End If
                If Field.Name = "CodiceCampoIDFeedPerClass01" Then
                    Field.Value = CodiceCampoIDFeedPerClass01
                End If
                If Field.Name = "MsgInDocSeRigaMerceSenzaImballo" Then
                    Field.Value = MsgInDocSeRigaMerceSenzaImballo
                End If
                If Field.Name = "CodiceCampoIDFeedPerClass02" Then
                    Field.Value = CodiceCampoIDFeedPerClass02
                End If
                If Field.Name = "CodiceCampoIDFeedPerAcquisto" Then
                    Field.Value = CodiceCampoIDFeedPerAcquisto
                End If
                If Field.Name = "IDAnagraficaDestinazionePerCertificato" Then
                    Field.Value = IDAnagraficaDestinazionePerCertificato
                End If
                If Field.Name = "IDCategoriaAnagraficaSocioDiretto" Then
                    Field.Value = IDCategoriaAnagraficaSocioDiretto
                End If
                If Field.Name = "IDCategoriaAnagraficaProdAcq" Then
                    Field.Value = IDCategoriaAnagraficaProdAcq
                End If
                If Field.Name = "IDArticoloScartoPerCertificato" Then
                    Field.Value = IDArticoloScartoPerCertificato
                End If
                If Field.Name = "RiportaDestinazioneDaContrattoCertificato" Then
                    Field.Value = RiportaDestinazioneDaContrattoCertificato
                End If
                If Field.Name = "RiportaVettoreDaContrattoCertificato" Then
                    Field.Value = RiportaVettoreDaContrattoCertificato
                End If
                If Field.Name = "ForzaDestinazioneDaContrattoCertificato" Then
                    Field.Value = ForzaDestinazioneDaContrattoCertificato
                End If
                If Field.Name = "ForzaVettoreDaContrattoCertificato" Then
                    Field.Value = ForzaVettoreDaContrattoCertificato
                End If
                If Field.Name = "NonInviareCodiceForInFeedentity" Then
                    Field.Value = NonInviareCodiceForInFeedentity
                End If
                If Field.Name = "PrendiVarietaDaTipologiaFeedentity" Then
                    Field.Value = PrendiVarietaDaTipologiaFeedentity
                End If
                If Field.Name = "AttivaSelezioneSocioCertPerVarieta" Then
                    Field.Value = AttivaSelezioneSocioCertPerVarieta
                End If
                If Field.Name = "AttivaSelezioneAnaVeloceInCert" Then
                    Field.Value = AttivaSelezioneAnaVeloceInCert
                End If
                If Field.Name = "AttivaPaginazioneContattiFeed" Then
                    Field.Value = AttivaPaginazioneContattiFeed
                End If
                If Field.Name = "NumeroElementiContattiPerPagina" Then
                    Field.Value = NumeroElementiContattiPerPagina
                End If
                If Field.Name = "AttivaRicercaFatturaAccontoBIVendite" Then
                    Field.Value = AttivaRicercaFatturaAccontoBIVendite
                End If
                If Field.Name = "AttivaRicercaFatturaAccontoBIFatturato" Then
                    Field.Value = AttivaRicercaFatturaAccontoBIFatturato
                End If
                If Field.Name = "NumeroMesiPerDataRevocaCertificato" Then
                    Field.Value = NumeroMesiPerDataRevocaCertificato
                End If
                If Field.Name = "NonRiportaInXMLRifVsNumOrd" Then
                    Field.Value = NonRiportaInXMLRifVsNumOrd
                End If
                If Field.Name = "DataUltimaSincronizzazioneContattiFeed" Then
                    If (Field.Value <> "") Then
                        Field.Value = DataUltimaSincronizzazioneContattiFeed
                    End If
                End If
                If Field.Name = "DataUltimaSincronizzazioneLottoFeed" Then
                    If (Field.Value <> "") Then
                        Field.Value = DataUltimaSincronizzazioneLottoFeed
                    End If
                End If
                If Field.Name = "AttivaGestioneUltimaSincronizzazioneFeed" Then
                    Field.Value = AttivaGestioneUltimaSincronizzazioneFeed
                End If
                If Field.Name = "NonInviareRifCertificatoIdDDT" Then
                    Field.Value = NonRiportareRifCerticatoInDDT
                End If
                
                If Field.Name = "AttivaControlloEsistenzaLottiInFeedAutInSync" Then
                    Field.Value = AttivaControlloEsistenzaLottiInFeedAutInSync
                End If
                If Field.Name = "NonEliminareLottiDefinitivamenteDaFeed" Then
                    Field.Value = NonEliminareLottiDefinitivamenteDaFeed
                End If
                If Field.Name = "NonEliminareLottiProvvDefinitivamenteDaFeed" Then
                    Field.Value = NonEliminareLottiProvvDefinitivamenteDaFeed
                End If
                If Field.Name = "NonAggiornareRifLottoInFeed" Then
                    Field.Value = NonAggiornareRifLottoInFeed
                End If
                If Field.Name = "DBNameRegFatture" Then
                    Field.Value = DBNameRegFatture
                End If
                If Field.Name = "AttivaMappaturaDaRegFatture" Then
                    Field.Value = AttivaMappaturaDaRegFatture
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
    
    m_Document.SaveDocument
    
    SALVA_PAR_QUAL TheApp.IDFirm
    SET_PAR_CONN_ALY
    
    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
    
    'Refresh delle variabili di stato
    m_Changed = False
    m_Search = False
    m_Saved = True
    
    'Refresh dello stato della ToolBar standard in modalità variazione
    SetStatus4Modality Modify
       
    
  
    'If NuovoContratto = True Then
    '    CreaStoriaContratto m_Document("IDRV_POContratto")
    '    SviluppoRateContratto (m_Document("IDRV_POContratto").Value)
    '    m_DocumentsLink.Refresh
        
    'End If
    
    
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
    Dim sToRemove As String
    Dim DocLink As DmtDocManLib.DocumentsLink
    
    
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

        If Not (m_Document.EOF Or m_Document.BOF) Then
            'Cancella l'eventuale blocco sul record da cancellare.
            m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
        End If
        
        
        
        'rif16
        
        'Cancellazione
        m_Document.DeleteDocument
        
        
        
        
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
        
        'Ripulisce la collezione Fields dell'oggetto DocType.
        For Each Field In m_DocType.Fields
            Field.Value = Empty
        Next
        
        'Viene inserita la condizione di ricerca basata sull'ID del record corrente.
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
    SetStatus4Modality Preview, ClosePrw
        
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

    



Private Sub cboCatCommTrasporto_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCategoriaAnagraficaSocio_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleAumentoPeso_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleAumentoPeso_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleCaloPeso_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleCaloPeso_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleCarico_Car_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleCarico_Vend_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub cboCausaleScarico_Car_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub cboCausaleScarico_Vend_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub cboCausaleScaricoAumentoPeso_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleScaricoCaloPeso_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleScaricoTipoScarto_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausaleTipoScarto_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboCausContPerFatturazione_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboDecimaliConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboDecimaliLav_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboDecimaliVend_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboFiliale_Click()


    With Me.cboLuogoPresaMerce
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDSitoPerAnagrafica"
        .DisplayField = "SitoPerAnagrafica"
        .SQL = "SELECT * FROM SitoPerAnagrafica  "
        .SQL = .SQL & "WHERE IDAnagrafica=" & GET_LINK_ANAGRAFICA_AZIENDA(TheApp.IDFirm)
        .SQL = .SQL & " ORDER BY SitoPerAnagrafica "
        .Fill
    End With


    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboFunzioneDDTAcq_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboFunzioneFAAcq_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub cboIntraProvincia_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboIvaPerFatturazione_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboLingua_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboListinoCampionatura_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboListinoImballiDefault_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboLuogoPresaMerce_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboMagazzinoCarico_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboMagazzinoVendita_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboPagamentoPerFatturazione_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboPorto_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboSezionaleCMR_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboSezListaPrelievo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoArrotondamento_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoArrotondamentoConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoAumentoPeso_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoAumentoPeso_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoCaloPeso_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoCaloPeso_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoComportamentoLavorazione_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoCorpoFattSocio_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoPesoArticolo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoProdottoGrezzo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoProdottoImballo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoProdottoLavorato_Click()
 If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoQuadratura_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoQuadratura_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoScarto_Click()
If Not (BrwMain.Visible) Then Change
End Sub
Private Sub cboTipoSceltaArticoloConferito_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboTipoSpezzatura_Click()
If Not (BrwMain.Visible) Then Change

End Sub

Private Sub cboTipoTrattAggVivaio_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboUMRigaOrdine_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cboValutaPerFatturazione_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDArtPesPedNeg_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDClienteOrdinePred_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub CDModoTrasporto_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDNaturaTransazione_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub CDTipoPedana_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub chkAbilitaForLotto_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAggPrezzoMedioDaConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAggTipoLavDaConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAttivaCalcoloPesoConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAttivaDescDocMultiline_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkAttivaUMDocVend_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkCalcoloTraspPerPeso_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkContributi_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkDocVendForzaNuovo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkEscludiImbImpOrd_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkFocusColliIVGamma_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkForzaTabulazioneConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkGestioneConferimento_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkGestioneVivaio_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkIvaARendere_Click()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkIvaBloccata_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub chkLiquidazioneLordoIVA_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkLottoCampagnaObbligatorio_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub chkLottoProdPerSocio_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkNonRipAcconti_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkNonRipScartiInFatt_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkNumerazioneFACSCoop_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkNuovoCalcolo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkNuovoRecLav_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkOrdineAutomatico_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkPedanaAutomatica_Click()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub chkPrezziArtDaOrdine_Click()
        If Not (BrwMain.Visible) Then Change
End Sub



Private Sub chkPrezzoMedioDaConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkRaggrLiqSocio_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkRipDescrRifLiq_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkRipDiffColliDaLavDaOrd_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkRipPedOrdSel_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkStampaDocNonAtt_Click()
     If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkStampaEtiPDF_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkStampaEvOrdAtt_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkStampaNuovoMetodo_Click()
     If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkStampaPedana_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkStampaRiepilogoImballiConf_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkUsaProtICEPeriodo_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisAndamDaOrd_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisAndamLav_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisEtiLavUtente_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub chkVisEtiPedUtente_Click()
     If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisImportoF4_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisListaMerceOrd_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisMasAvvioVeloce_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisOrdNewPed_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkVisOrdPrep_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub chkVisOrdSmist_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chKVisRiepConfVend_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkZeriRifDoc_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdAggiornaReportEtichette_Click()
On Error GoTo ERR_cmdAggiornaReportEtichette_Click
Dim IDTipoOggettoEtichette As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset


If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

Screen.MousePointer = 11

IDTipoOggettoEtichette = fnGetTipoOggetto("RV_POEtichetteLavorazione")

Set rsNew = New ADODB.Recordset

rsNew.Open "RV_POEtichetteDefault", Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rsNew.EOF
    rsNew!IDReportPerTipoOggetto = GET_LINK_REPORT(IDTipoOggettoEtichette, fnNotNull(rsNew!ReportPerTipoOggetto))
    rsNew.Update
rsNew.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

Screen.MousePointer = 0
Exit Sub
ERR_cmdAggiornaReportEtichette_Click:
    MsgBox Err.Description, vbCritical, "ERR_cmdAggiornaReportEtichette_Click"
    Screen.MousePointer = 0
End Sub
Private Function GET_LINK_REPORT(IDTipoOggetto As Long, NomeReport As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDReportTipoOggetto FROM ReportTipoOggetto "
sSQL = sSQL & "WHERE IDTipoOggetto=" & IDTipoOggetto
sSQL = sSQL & " AND ReportTipoOggetto=" & fnNormString(NomeReport)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_LINK_REPORT = 0
Else
    GET_LINK_REPORT = fnNotNullN(rs!IDReportTipoOggetto)
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
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = fnNotNullN(rs!IDTipoOggetto)
    Else
        fnGetTipoOggetto = 0
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub cmdAlyante_Click()
    frmAlyante.Show vbModal
    If CONFERMA_PARAMETRI_ALY = True Then
        If Not (BrwMain.Visible) Then Change
    End If
End Sub

Private Sub cmdArtDerivatiDaOrd_Click()
On Error Resume Next
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsNew As ADODB.Recordset
Dim NumeroRecord As Long
Dim Unita_Progresso As Double

If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

'''''CALCOLO NUMERO ARTICOLI''''''''''''''''''
sSQL = "SELECT COUNT(IDArticolo) AS NumeroRecord "
sSQL = sSQL & "FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    NumeroRecord = 0
Else
    NumeroRecord = fnNotNullN(rs!NumeroRecord)
End If

rs.CloseResultset
Set rs = Nothing

'''''''''''''''''''''''''''''''''''''''''''''

If NumeroRecord = 0 Then
    MsgBox "Non ci sono articoli da elaborare", vbInformation, "Configurazione gestione articoli da ordine"
    Exit Sub
End If

Me.ProgressBar1.Visible = True
Me.ProgressBar1.Value = 0
Me.ProgressBar1.Max = 1000
Unita_Progresso = FormatNumber((Me.ProgressBar1.Max / NumeroRecord), 4)


sSQL = "SELECT IDArticolo FROM Articolo "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm
sSQL = sSQL & "ORDER BY CodiceArticolo, Articolo"
Set rs = Cn.OpenResultset(sSQL)

sSQL = "SELECT * FROM RV_POArticoloFiglioOrdine "
Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

Me.Enabled = False

While Not rs.EOF
    Screen.MousePointer = 11
    
    If GET_ESISTENZA_ARTICOLO_DERIVATO_DA_ORDINE(fnNotNullN(rs!IDArticolo)) = False Then
        rsNew.AddNew
            rsNew!IDRV_POArticoloFiglioOrdine = fnGetNewKey("RV_POArticoloFiglioOrdine", "IDRV_POArticoloFiglioOrdine")
            rsNew!IDArticolo = fnNotNullN(rs!IDArticolo)
            rsNew!IDArticoloFiglio = fnNotNullN(rs!IDArticolo)
            rsNew!PesoPerOrdinamento = GET_PESO_PER_ORDINAMENTO(fnNotNullN(rs!IDArticolo))
        rsNew.Update
    End If
    
    Screen.MousePointer = 0
    
    If (Me.ProgressBar1.Value + Unita_Progresso) >= Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Max
    Else
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + Unita_Progresso
    End If
    DoEvents
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
Me.Enabled = True
MsgBox "Elaborazione avvenuta con successo", vbInformation, "Configurazione gestione articoli da ordine"
Me.ProgressBar1.Visible = False
End Sub
Private Function GET_PESO_PER_ORDINAMENTO(IDArticolo As Long) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT MAX(PesoPerOrdinamento) AS Numero "
sSQL = sSQL & "FROM RV_POArticoloFiglioOrdine "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_PESO_PER_ORDINAMENTO = 1
Else
    GET_PESO_PER_ORDINAMENTO = fnNotNullN(rs!Numero) + 1
End If
rs.CloseResultset
Set rs = Nothing

End Function
Private Function GET_ESISTENZA_ARTICOLO_DERIVATO_DA_ORDINE(IDArticolo As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POArticoloFiglioOrdine   "
sSQL = sSQL & "WHERE IDArticolo=" & IDArticolo
sSQL = sSQL & " AND IDArticoloFiglio=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_ESISTENZA_ARTICOLO_DERIVATO_DA_ORDINE = False
Else
    GET_ESISTENZA_ARTICOLO_DERIVATO_DA_ORDINE = True
End If
rs.CloseResultset
Set rs = Nothing


End Function
Private Sub cmdElimina_Interessi_Click()
    
    m_DocumentsLink6.DeleteRowFromBuffer
    
    If Not (BrwMain.Visible) Then Change
    
End Sub

Private Sub cmdElimina_OperazionePerDoc_Click()
    
    m_DocumentsLink4.DeleteRowFromBuffer
    
    If Not (BrwMain.Visible) Then Change
    
End Sub

Private Sub cmdElimina_Processo_Click()
    m_DocumentsLink1.DeleteRowFromBuffer
    
    If Not (BrwMain.Visible) Then Change
    
End Sub

Private Sub cmdElimina_Sezionale_Click()
    m_DocumentsLink2.DeleteRowFromBuffer
    
    If Not (BrwMain.Visible) Then Change

End Sub

Private Sub cmdEliminaProtocolloICE_Click()
    m_DocumentsLink3.DeleteRowFromBuffer
    
    If Not (BrwMain.Visible) Then Change

End Sub



Private Sub cmdEliminaRifOrdIVGamma_Click()
Dim Testo As String
If Me.txtIDOrdineIVGamma.Value = 0 Then Exit Sub

If Me.txtIDOrdineIVGamma.Value > 0 Then
    Testo = "Sei sicuro di voler eliminare il riferimento dell'ordine?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento ordine di IV Gamma") = vbNo Then Exit Sub
    Me.txtIDOrdineIVGamma.Value = 0
End If
End Sub

Private Sub cmdEliminaRifOrdIVGammaLav_Click()
Dim Testo As String
If Me.txtIDOrdineIVGammaLav.Value = 0 Then Exit Sub

If Me.txtIDOrdineIVGammaLav.Value > 0 Then
    Testo = "Sei sicuro di voler eliminare il riferimento dell'ordine?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riferimento ordine di IV Gamma") = vbNo Then Exit Sub
    Me.txtIDOrdineIVGammaLav.Value = 0
End If
End Sub

Private Sub cmdEliminaUtenteEva_Click()
    m_DocumentsLink7.Delete
End Sub

Private Sub cmdNuovo_Interessi_Click()
    If m_DocumentsLink6.TableNew Then
        m_DocumentsLink6.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink6.NewRow
    
    Me.txtDataInizioInteressi.Value = Date
    Me.txtDataInizioInteressi.SetFocus
    
End Sub

Private Sub cmdNuovo_OperazionePerDoc_Click()
    If m_DocumentsLink4.TableNew Then
        m_DocumentsLink4.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink4.NewRow
    
    Me.chkGestioneArticoli.Value = vbUnchecked
    Me.chkCreazioneAutomatica.Value = vbUnchecked

    Me.cboDocCoopPerOpe.Enabled = True

End Sub

Private Sub cmdNuovo_Processo_Click()
    If m_DocumentsLink1.TableNew Then
        m_DocumentsLink1.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink1.NewRow
    
    
    Me.cboDocumentoCoop.Enabled = True
End Sub


Private Sub cmdNuovo_Sezionale_Click()
    If m_DocumentsLink2.TableNew Then
        m_DocumentsLink2.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink2.NewRow
End Sub

Private Sub cmdNuovoProtocolloICE_Click()
    If m_DocumentsLink3.TableNew Then
        m_DocumentsLink3.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink3.NewRow

End Sub

Private Sub cmdNuovoUtenteEva_Click()
    If m_DocumentsLink7.TableNew Then
        m_DocumentsLink7.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink7.NewRow
    Me.cboUtenteOrd.SetFocus
End Sub

Private Sub cmdOrdIVGamma_Click()
    
    Set ControlIDOrdine = Me.txtIDOrdineIVGamma
    
    frmSelOrdine.Show vbModal
    
End Sub

Private Sub cmdOrdIVGammaLav_Click()
    Set ControlIDOrdine = Me.txtIDOrdineIVGammaLav
    
    frmSelOrdine.Show vbModal
    
End Sub

Private Sub cmdProtICEPeriodo_Click()
    
    If Me.cboFiliale.CurrentID = 0 Then Exit Sub
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    If Me.chkUsaProtICEPeriodo.Value = vbUnchecked Then Exit Sub
    
    frmProtICEPeriodo.Show vbModal
    
End Sub

Private Sub cmdSalva__OperazionePerDoc_Click()
    m_DocumentsLink4("IDRV_PODocumentoCoop").Value = Me.cboDocCoopPerOpe.CurrentID
    m_DocumentsLink4("GestioneArticoli").Value = Me.chkGestioneArticoli.Value
    m_DocumentsLink4("CreazioneAutomaticaLottoVend").Value = Me.chkCreazioneAutomatica.Value
    m_DocumentsLink4("IDRV_POTipoLavorazione").Value = Me.cboTipoLavorazioneAut.CurrentID
    m_DocumentsLink4.SaveRowToBuffer
    
    m_DocumentsLink4.Move Me.GrigliaOperazione.ListIndex - 1
    
    
    If Not (BrwMain.Visible) Then Change
    
End Sub



Private Sub cmdSalva_Interessi_Click()
    If Me.txtDataInizioInteressi.Value = 0 Then
        MsgBox "Inserire la data di inizio interessi", vbInformation, "Controllo inserimento"
        Me.txtDataInizioInteressi.SetFocus
        Exit Sub
    End If
    If Me.txtTassoInteressi.Value = 0 Then
        MsgBox "Inserire il tasso di interesse", vbInformation, "Controllo inserimento"
        Me.txtTassoInteressi.SetFocus
        Exit Sub
    End If
    
    If Me.txtNumeroGiorniAnno.Value = 0 Then
        txtDataInizioInteressi_LostFocus
    End If
    
    
    txtTassoInteressi_LostFocus
    
    m_DocumentsLink6("IDAzienda").Value = TheApp.IDFirm
    m_DocumentsLink6("IDFiliale").Value = TheApp.Branch
    m_DocumentsLink6("DataInizio").Value = Me.txtDataInizioInteressi.Value
    m_DocumentsLink6("TassoInteresse").Value = Me.txtTassoInteressi.Value
    m_DocumentsLink6("NumeroGiorniAnno").Value = Me.txtNumeroGiorniAnno.Value
    m_DocumentsLink6("TassoGiornaliero").Value = Me.txtPercentualeGiorno.Value
    
    m_DocumentsLink6.SaveRowToBuffer
    
    m_DocumentsLink6.Move Me.GrigliaInteressi.ListIndex - 1
    
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdSalva_Processo_Click()
    m_DocumentsLink1("IDDocumentoCoop").Value = Me.cboDocumentoCoop.CurrentID
    'm_DocumentsLink1("IDMagazzino").Value = Me.cboMagazzinoProcesso.CurrentID
    m_DocumentsLink1("IDTipoProcessoCoop").Value = Me.cboTipoProcesso.CurrentID
    m_DocumentsLink1("IDTipoMagazzino").Value = Me.cboTipoMagazzino.CurrentID
    m_DocumentsLink1("IDFunzione").Value = Me.cboCausaliMagazzino.CurrentID
    
    m_DocumentsLink1.SaveRowToBuffer
    
    m_DocumentsLink1.Move Me.GrigliaProcessi.ListIndex - 1
    
    
    If Not (BrwMain.Visible) Then Change

End Sub



Private Sub cmdSalva_Sezionale_Click()
    
    If Me.chkPredefinito_Sezionale.Value = 1 Then
        If ControlloPredefinitoSezionale = True Then
            m_DocumentsLink2("IDSezionale").Value = Me.cboSezionale.CurrentID
            m_DocumentsLink2("IDDocumentoCoop").Value = Me.cboDocumenti_PerSezionale.CurrentID
            m_DocumentsLink2("Predefinito").Value = Me.chkPredefinito_Sezionale.Value
    
            m_DocumentsLink2.SaveRowToBuffer
    
            m_DocumentsLink2.Move Me.GrigliaSezionale.ListIndex - 1
            
            If Not (BrwMain.Visible) Then Change
        Else
            MsgBox "ATTENZIONE!!" & vbCrLf & "E' già presente un sezionale predefinito per questo documento.", vbInformation, "Impossibile salvare"
            Exit Sub
        End If
    Else
        m_DocumentsLink2("IDSezionale").Value = Me.cboSezionale.CurrentID
        m_DocumentsLink2("IDDocumentoCoop").Value = Me.cboDocumenti_PerSezionale.CurrentID
        m_DocumentsLink2("Predefinito").Value = Me.chkPredefinito_Sezionale.Value
    
        m_DocumentsLink2.SaveRowToBuffer
    
        m_DocumentsLink2.Move Me.GrigliaSezionale.ListIndex - 1
    
    
        If Not (BrwMain.Visible) Then Change
        
        
        
    End If

End Sub

Private Sub cmdSalvaProtocolloICE_Click()
    
    If Me.chkPredefinito_ICE.Value = 1 Then
        If ControlloPredefinitoICE(fnNotNullN(m_DocumentsLink3(m_DocumentsLink3.PrimaryKey).Value)) = True Then
            m_DocumentsLink3("IDRV_POProtocolloICE").Value = Me.cboProtocolloICE.CurrentID
            m_DocumentsLink3("IDEsercizio").Value = Me.cboEsercizioICE.CurrentID
            m_DocumentsLink3("Predefinito").Value = Me.chkPredefinito_ICE.Value
            m_DocumentsLink3("Progressivo").Value = Me.txtNumeroProgressivo.Value
    
        Else
            MsgBox "ATTENZIONE!!" & vbCrLf & "E' già presente un protocollo ICE predefinito", vbInformation, "Impossibile salvare"
            Exit Sub
        End If
    Else
        m_DocumentsLink3("IDRV_POProtocolloICE").Value = Me.cboProtocolloICE.CurrentID
        m_DocumentsLink3("IDEsercizio").Value = Me.cboEsercizioICE.CurrentID
        m_DocumentsLink3("Predefinito").Value = Me.chkPredefinito_ICE.Value
        m_DocumentsLink3("Progressivo").Value = Me.txtNumeroProgressivo.Value
    End If
    
    m_DocumentsLink3.SaveRowToBuffer

    m_DocumentsLink3.Move Me.GrigliaProgressivoICE.ListIndex - 1


    If Not (BrwMain.Visible) Then Change

End Sub



Private Sub cmdSalvaUtenteEva_Click()



If Me.cboUtenteOrd.CurrentID = 0 Then
    MsgBox "Inserire l'utente", vbCritical, "Inserimento parametri evasione ordini"
    Exit Sub
End If
If fnNotNullN(m_DocumentsLink7(m_DocumentsLink7.PrimaryKey).Value) <= 0 Then
    If GET_ESISTENZA_UTENTE_EVA(Me.cboUtenteOrd.CurrentID, TheApp.IDFirm) = True Then
        MsgBox "La configurazione di questo utente è già stata inserita", vbCritical, "Inserimento parametri evasione ordini"
        Exit Sub
    End If
End If
If Me.txtIDOrdineSmist.Value = 0 Then
    MsgBox "Inserire almeno l'ordine da smistare", vbCritical, "Inserimento parametri evasione ordini"
    Exit Sub
End If
If Me.txtIDOrdineSmist.Value = Me.txtIDOrdine.Value Then
    MsgBox "Nell'ordine di smistamento e nell'ordine di preparazione è stato inserito lo stesso ordine", vbCritical, "Inserimento parametri evasione ordini"
    Exit Sub
End If


    m_DocumentsLink7("IDUtente").Value = Me.cboUtenteOrd.CurrentID
    m_DocumentsLink7("IDOggettoOrdinePrep").Value = Me.txtIDOrdine.Value
    m_DocumentsLink7("IDOggettoOrdineSmist").Value = Me.txtIDOrdineSmist.Value
    m_DocumentsLink7("IDAzienda").Value = TheApp.IDFirm
    m_DocumentsLink7("IDFiliale").Value = TheApp.Branch
    
    If Me.txtIDOrdineSmist.Value > 0 Then
        m_DocumentsLink7("DataOrdineSmist").Value = Me.txtDataOrdineSmist.Text
        m_DocumentsLink7("NumeroOrdineSmist").Value = Me.txtNumeroOrdineSmist.Value
        m_DocumentsLink7("IDClienteOrdineSmist").Value = Me.CDClienteSmist.KeyFieldID
    End If
    If Me.txtIDOrdine.Value > 0 Then
        m_DocumentsLink7("DataOrdinePrep").Value = Me.txtDataOrdine.Value
        m_DocumentsLink7("NumeroOrdinePrep").Value = Me.txtNumeroOrdine.Value
        m_DocumentsLink7("IDClienteOrdinePrep").Value = Me.cdCliente.KeyFieldID
        m_DocumentsLink7("BloccaOrdinePrep").Value = Me.chkBloccoOrdPrep.Value
    End If
    m_DocumentsLink7.Save
    
    m_DocumentsLink7.Move Me.GrigliaUtenteEva.ListIndex - 1
    
End Sub

Private Sub cmdTracciabilita_Click()
On Error GoTo ERR_cmdTracciabilita_Click
    If m_Changed = True Then
        MsgBox "Salvare i parametri prima di procedere alla configurazione della tracciabilità on line", vbInformation, "Validazione dati"
        Exit Sub
    End If
    
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    LINK_PARAMETRI_FILIALE = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    frmTracciabilita.Show vbModal
    
    If CONFERMA_TRACC_ONLINE = 1 Then
        NewRecord
        NewSearch
        ExecuteSearch
        brwMain_DblClick
    End If
Exit Sub
ERR_cmdTracciabilita_Click:
    MsgBox Err.Description, vbCritical, "cmdTracciabilita_Click"
End Sub

Private Sub cmdTrovaOrdPrep_Click()
    frmTrovaOrdine.Show vbModal
End Sub

Private Sub cmdTrovaOrdSmist_Click()
    frmTrovaOrdineSmistamento.Show vbModal

End Sub

Private Sub DmtCodDesc1_ChangeElement()
If Not (BrwMain.Visible) Then Change
End Sub



Private Sub Command1_Click()
If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then Exit Sub

frmAltreOperazioni.Show vbModal
End Sub

Private Sub Command2_Click()
If fnNotNullN(m_Document(m_Document.PrimaryKey)) <= 0 Then Exit Sub

    CONFERMA_ALTRI_PARAMETRI = False
    frmAltriParametri.Show vbModal
    
    If CONFERMA_ALTRI_PARAMETRI = True Then
        If Not (BrwMain.Visible) Then Change
    End If
End Sub

Private Sub Command3_Click()
    frmCausaliXML.Show vbModal
    
    If CONFERMA_PARAMETRI_XML = True Then
        If Not (BrwMain.Visible) Then Change
    End If
End Sub

Private Sub Command4_Click()
    frmParametriQual.Show vbModal
    If CONFERMA_PARAMETRI_QUAL = True Then
        If Not (BrwMain.Visible) Then Change
    End If
End Sub

Private Sub Form_Activate()
On Error GoTo ERR_Form_Activate
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL As String
Dim AvviaAutomatismo As Long

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
        AvviaAutomatismo = 0
        
        sSQL = "SELECT IDRV_POSchemaCoop FROM RV_POSchemaCoop "
        sSQL = sSQL & "WHERE IDFiliale=" & TheApp.Branch
        
        Set rs = Cn.OpenResultset(sSQL)
        
        If Not rs.EOF Then
            AvviaAutomatismo = 1
        End If
        
        rs.CloseResultset
        Set rs = Nothing
        
        If AvviaAutomatismo = 1 Then
        
            NewRecord
            NewSearch
            ExecuteSearch
            brwMain_DblClick
        Else
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
    End If
Exit Sub
ERR_Form_Activate:
    MsgBox Err.Description, vbCritical, "Form_Activate"
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
    
    
    Select Case KeyCode
        Case vbKeyPageDown
            DMTSplitBar1.ScrollDown
        Case vbKeyPageUp
            DMTSplitBar1.ScrollUp
    End Select

    
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





Private Sub lblPianodeiDeiConti_Click(Index As Integer)
        VarIDEsercizio = fnGetEsercizio(Date)
        
        Link_PianoDeiConti = GetPianoDeiConti
    
        SetPDCProperties
        
        oPDC.ShowSearchDialog
        
        ShowNodeProperties oPDC.SelectedNode, Index

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
        Set Me.GrigliaProcessi.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink1.TableName).Data
        Set Me.GrigliaSezionale.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink2.TableName).Data
        Set Me.GrigliaProgressivoICE.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink3.TableName).Data
        Set Me.GrigliaOperazione.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink4.TableName).Data
        Set Me.GrigliaInteressi.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink6.TableName).Data
        Set Me.GrigliaUtenteEva.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink7.TableName).Data
    
    'End If
    If m_Document("IDRV_POSchemaCoop").Value > 0 Then
        Me.cboFiliale.Enabled = False
        Link_ContoPDC = fnNotNullN(m_Document("IDPDCFatturazione").Value)
        Link_ContoPDC_RigaNeg = fnNotNullN(m_Document("IDPDCRigheLiqNegative").Value)
        Link_ContoPDC_RigaPos = fnNotNullN(m_Document("PDCCodiceRigheLiqPositive").Value)
        
        PAR_NonCalcImpDaAssVeloce = fnNotNullN(m_Document("NonCalcImpDaAssVeloce").Value)
        PAR_CalcImpAConfOrd = fnNotNullN(m_Document("CalcImpAConfOrd").Value)
        PAR_NonVisMsgImpZeroConfOrd = fnNotNullN(m_Document("NonVisMsgImpZeroConfOrd").Value)
        PAR_AssNewPedDaAssSingola = fnNotNullN(m_Document("AssNewPedDaAssSingola").Value)
        PAR_NonCalcPrezzoDueRefArtOrd = fnNotNullN(m_Document("NonCalcPrezzoDueRefArtOrd").Value)
        ATT_MULT_SEL_IV_GAMMA = fnNotNullN(m_Document("AvviaNewSelQGammaIn").Value)
        CHIUDI_CONF_QTASEL_ZERO = fnNotNullN(m_Document("ChiudiConfQtaSelZeroNewSelQGammaIn").Value)
        GG_DATA_SCADENZA_CONTR = fnNotNullN(m_Document("NumeroGiorniDataScadContratto").Value)
        FOCUS_IVGAMMA_LAV = fnNotNullN(m_Document("IDRV_POFocusControlIVGammaLav").Value)
        FOCUS_IVGAMMA_CONF = fnNotNullN(m_Document("IDRV_POFocusControlIVGammaConf").Value)
        LINK_UM_COOP_SEL_AUT_IVGAMMA = fnNotNullN(m_Document("IDUMCoopSelAutIVGamma").Value)
        PAR_VIS_NOTE_RIGA_ORD_ELENCO = fnNotNullN(m_Document("VisualizzaNoteLavDaOrdineInElenco").Value)
        PAR_VIS_NOTE_RIGA_ORD = fnNotNullN(m_Document("VisualizzaNoteLavDaOrdine").Value)
        DATA_SCADENZA_OBBL = fnNotNullN(m_Document("DataScadenzaContrattoObbligatoria").Value)
        PAR_IDARTICOLO_BOLLO = fnNotNullN(m_Document("IDArticoloBolloNC").Value)
        PAR_IDSOCIO_RISC_PESO = fnNotNullN(m_Document("IDSocioPerRiscontroPeso").Value)
        GtinMigros = fnNotNull(m_Document("Gtin").Value)
        UrlMigros = fnNotNull(m_Document("UrlMigros").Value)
        NomeUtenteMigros = fnNotNull(m_Document("NomeUtenteMigros").Value)
        PasswordMigros = fnNotNull(m_Document("PasswordMigros").Value)
        ChiaveFeed = fnNotNull(m_Document("ChiaveClienteFeedentity").Value)
        UrlFeed = fnNotNull(m_Document("UrlServizioFeedentity").Value)
        AttivaFeed = Abs(fnNotNullN(m_Document("AttivaServizioFeedentity").Value))
        
        ATTIVA_PROD_CONF_LAV = Abs(fnNotNullN(m_Document("AttivaGestionePrelieviConferimento").Value))
        ATTIVA_PROD_ORD_LAV = Abs(fnNotNullN(m_Document("AttivaGestioneOrdiniInLavorazione").Value))
        DISATTIVA_PROD_ORD_TESTA = Abs(fnNotNullN(m_Document("DisabilitaListaTestaOrdine").Value))
        ATTIVA_PROD_CALC_COLLI_PED = Abs(fnNotNullN(m_Document("AttivaCalcoloPesoPedanaInProduzione").Value))
        ATTIVA_SEL_SUB_LOTTO_PROD = Abs(fnNotNullN(m_Document("SelezionaLottoCampagnaInLavorazione").Value))
        
        ATTIVA_SUB_CONTATTI_FEED = Abs(fnNotNullN(m_Document("AttivaMetodoSubContattiFeedentity").Value))
        ATTIVA_SEL_MULT_LOTTI_FEED = Abs(fnNotNullN(m_Document("AttivaMetodoSelMultLottoFeedentity").Value))
        ATTIVA_PROD_FIFO_ACCESSI_CONF = Abs(fnNotNullN(m_Document("AttivaControlloAccessoFifo").Value))
        
        CodiceAnnullaOperazione = fnNotNull(m_Document("CodiceRispristinoOperazioni").Value)
        CodiceConfermaOrdine = fnNotNull(m_Document("CodiceConfermaOrdine").Value)
        CodiceGestioneErrori = fnNotNull(m_Document("CodiceAnnullaOperazione").Value)
        CodiceConfermaViaggio = fnNotNull(m_Document("CodiceViaggioCompletato").Value)
        
        
        ATTIVA_GESTIONE_CARICO_MERCE = Abs(fnNotNullN(m_Document("CreaListaPrelievoAutPedInUscita").Value))
        NonVisMsgCreaListaPrelAutPedUscita = Abs(fnNotNullN(m_Document("NonVisMsgCreaListaPrelievoAutPedUscita").Value))
        ATTIVA_VisAndGrigliaConferimento = Abs(fnNotNullN(m_Document("VisualizzaAndamentoGrigliaConferimento").Value))
        ATTIVA_VisAndGrigliaRigaOdine = Abs(fnNotNullN(m_Document("VisualizzaAndamentoGrigliaRigaOdine").Value))
        ATTIVA_RICALCOLO_COLLI_PREL_CHIUSO = Abs(fnNotNullN(m_Document("ChiusuraPrelievoAggQtaPrelConColliEntrati").Value))
        IDTipoCollegRigaOrdRigaConf = fnNotNullN(m_Document("IDRV_POTipoCollegamentoRigaOrdRigaConf").Value)
        NonVisRigheOrdiniComplInSelLav = fnNotNullN(m_Document("NonVisualizzaRigheOrdiniCompletateInSelLav").Value)
        ModificaLavorazioneSenzaRicalcolo = fnNotNullN(m_Document("ModificaLavorazioneSenzaRicalcolo").Value)
        RISCONTRO_PESO_PER_CONF = Abs(fnNotNullN(m_Document("RiscontroPesoPerOgniRiga").Value))
        ATTIVA_ELIMINAZIONE_ACCESSI = Abs(fnNotNullN(m_Document("ConsentiEliminazioneAccessi").Value))
        IDSOCIO_PRE_CONF = Abs(fnNotNullN(m_Document("IDAnagraficaSocioMovPreConferimento").Value))
        ATTIVA_COMMISSIONI_ORDINI = Abs(fnNotNullN(m_Document("AttivaCommissioniDaOrdine").Value))
        RIC_COMM_TIPO_PED_EVAS_ORD = Abs(fnNotNullN(m_Document("RicalcolaCommTipoPedDaOrdInEvasione").Value))
        VIS_ELENCO_RIGHE_ORD = Abs(fnNotNullN(m_Document("VisElencoRigheOrdineSeNonTroviAssociazione").Value))
        ATTIVA_CALCOLO_N_PED_BI_VENDITE = Abs(fnNotNullN(m_Document("AttivaCalcoloNPedEffInVenditeBI").Value))
        RISCONTRO_PESO_VAL_SOCIO = Abs(fnNotNullN(m_Document("RiscontroPesoAValorePerSocio").Value))
        RISCONTRO_PESO_VAL_FORNITORE = Abs(fnNotNullN(m_Document("RiscontroPesoAValorePerFornitore").Value))
        OBBL_N_DOC_PES_CONF = Abs(fnNotNullN(m_Document("ObbligatorioNDocInPesConf")))
        ATTIVA_OBBL_SFALCIO_CONF = Abs(fnNotNullN(m_Document("ObbligatorioRifSfalcioConf")))
        ATTIVA_SEQ_SFALCIO = Abs(fnNotNullN(m_Document("ObbligatorioSequenzaSfalcioConf")))
        RIPORTA_RIF_DETTAGLIO_XML_NC = Abs(fnNotNullN(m_Document("RIferimentiDettagliatiInNC")))
        RIPORTA_RIF_DETTAGLIO_XML_ND = Abs(fnNotNullN(m_Document("RIferimentiDettagliatiInND")))
        CONF_AUT_CONTR_PRESA_VISIONE = Abs(fnNotNullN(m_Document("ConfPresaVisAutOrdDaContr")))
        NO_RIP_AGENTE_IN_DOC_EVASIONE = Abs(fnNotNullN(m_Document("NonRipAgenteDaEvOrdNoPres")))
        Rip_InXMLRifLetteraIntento = Abs(fnNotNullN(m_Document("RiportaInXMLRifLetteraIntento")))
        Rip_InXMLRifNoteIva = Abs(fnNotNullN(m_Document("RiportaInXMLRifNoteIva")))
        Rip_InXMLRifNota01Doc = Abs(fnNotNullN(m_Document("RiportaInXMLRifNota01Doc")))
        Rip_InXMLRifNota02Doc = Abs(fnNotNullN(m_Document("RiportaInXMLRifNota02Doc")))
        Rip_InXMLRifNota03Doc = Abs(fnNotNullN(m_Document("RiportaInXMLRifNota03Doc")))
        Rip_InXMLRifNotaDoc = Abs(fnNotNullN(m_Document("RiportaInXMLRifNotaDoc")))
        Rip_InXMLRifIstrMitt = Abs(fnNotNullN(m_Document("RiportaInXMLRifIstrMitt")))
        Rip_InXMLRifVettSucc = Abs(fnNotNullN(m_Document("RiportaInXMLRifLetteraIntento")))
        Rip_InXMLRifAgenziaTrasp = Abs(fnNotNullN(m_Document("RiportaInXMLRifAgenziaTrasp")))
        Rip_InXMLRifTargaAutoMezzo = Abs(fnNotNullN(m_Document("RiportaInXMLRifTargaAutoMezzo")))
        
        COMPANY_CODE = fnNotNullN(m_Document("Alyante_CompanyCode"))
        ATTIVA_FIDO_ALY = Abs(fnNotNullN(m_Document("Alyante_AttivaFido")))
        DISATTIVA_DDT_FIDO_ALY = Abs(fnNotNullN(m_Document("Alyante_DisattivaCalcoloDDT")))
        DISATTIVA_FD_FIDO_ALY = Abs(fnNotNullN(m_Document("Alyante_DisattivaCalcoloFD")))
        DISATTIVA_FA_FIDO_ALY = Abs(fnNotNullN(m_Document("Alyante_DisattivaCalcoloFA")))
        DISATTIVA_NC_FIDO_ALY = Abs(fnNotNullN(m_Document("Alyante_DisattivaCalcoloNC")))
        DISATTIVA_ND_FIDO_ALY = Abs(fnNotNullN(m_Document("Alyante_DisattivaCalcoloND")))
        DISATTIVA_SCALATA_COMM_TRASP = Abs(fnNotNullN(m_Document("DisattivaScalataCommTipoPedana")))
        ATTIVA_FEED_MULTI_LIVELLO = Abs(fnNotNullN(m_Document("AttivaMultilivelloFeedentity")))
        IDFEED_AZIENDA = fnNotNull(m_Document("RV_POIDFeedentity").Value)
        NUMERO_COLLI_PRED_CERT = fnNotNullN(m_Document("NumeroColliPerAutomezzoCert").Value)
        
        SincronizzaAutFeed = Abs(fnNotNullN(m_Document("SincronizzaAutFeed").Value))
        IvaArticoloDaDocColl = Abs(fnNotNullN(m_Document("IvaArticoloDaDocumentoCollegato").Value))
        LetteraIntentoDaDocColl = Abs(fnNotNullN(m_Document("LetteraIntentoDaDocumentoCollegato").Value))
        RifComuneDaConfigSocio = Abs(fnNotNullN(m_Document("RifComuneDaConfigurazioneSocio").Value))
        DocAnaDestUgualeAnaCoop = Abs(fnNotNullN(m_Document("DocAnaDestUgualeAnaCoop").Value))
        
        IDClassLottoProdPerFuoriQuota = fnNotNullN(m_Document("IDClassificazioneLottoProdPerFuoriQuota").Value)
        MsgInDocSeRigaMerceSenzaImballo = fnNotNullN(m_Document("MsgInDocSeRigaMerceSenzaImballo").Value)
        
        CodiceCampoIDFeedPerClass01 = fnNotNull(m_Document("CodiceCampoIDFeedPerClass01").Value)
        CodiceCampoIDFeedPerClass02 = fnNotNull(m_Document("CodiceCampoIDFeedPerClass02").Value)
        CodiceCampoIDFeedPerAcquisto = fnNotNull(m_Document("CodiceCampoIDFeedPerAcquisto").Value)
        
        IDCategoriaAnagraficaSocioDiretto = fnNotNullN(m_Document("IDCategoriaAnagraficaSocioDiretto").Value)
        IDCategoriaAnagraficaProdAcq = fnNotNullN(m_Document("IDCategoriaAnagraficaProdAcq").Value)
        IDCategoriaAnagraficaNoProd = fnNotNullN(m_Document("IDCategoriaAnagraficaNoProd").Value)
        IDArticoloScartoPerCertificato = fnNotNullN(m_Document("IDArticoloScartoPerCertificato").Value)
        
        RiportaDestinazioneDaContrattoCertificato = Abs(fnNotNullN(m_Document("RiportaDestinazioneDaContrattoCertificato").Value))
        RiportaVettoreDaContrattoCertificato = Abs(fnNotNullN(m_Document("RiportaVettoreDaContrattoCertificato").Value))
        ForzaDestinazioneDaContrattoCertificato = Abs(fnNotNullN(m_Document("ForzaDestinazioneDaContrattoCertificato").Value))
        ForzaVettoreDaContrattoCertificato = Abs(fnNotNullN(m_Document("ForzaVettoreDaContrattoCertificato").Value))
        
        NonInviareCodiceForInFeedentity = Abs(fnNotNullN(m_Document("NonInviareCodiceForInFeedentity").Value))
        PrendiVarietaDaTipologiaFeedentity = Abs(fnNotNullN(m_Document("PrendiVarietaDaTipologiaFeedentity").Value))
        
        AttivaSelezioneSocioCertPerVarieta = Abs(fnNotNullN(m_Document("AttivaSelezioneSocioCertPerVarieta").Value))
        AttivaSelezioneAnaVeloceInCert = Abs(fnNotNullN(m_Document("AttivaSelezioneAnaVeloceInCert").Value))
        
        
        AttivaPaginazioneContattiFeed = Abs(fnNotNullN(m_Document("AttivaPaginazioneContattiFeed").Value))
        NumeroElementiContattiPerPagina = Abs(fnNotNullN(m_Document("NumeroElementiContattiPerPagina").Value))
        AttivaRicercaFatturaAccontoBIVendite = Abs(fnNotNullN(m_Document("AttivaRicercaFatturaAccontoBIVendite").Value))
        AttivaRicercaFatturaAccontoBIFatturato = Abs(fnNotNullN(m_Document("AttivaRicercaFatturaAccontoBIFatturato").Value))
        NumeroMesiPerDataRevocaCertificato = Abs(fnNotNullN(m_Document("NumeroMesiPerDataRevocaCertificato").Value))
        IDAnagraficaDestinazionePerCertificato = fnNotNullN(m_Document("IDAnagraficaDestinazionePerCertificato").Value)
        NonRiportaInXMLRifVsNumOrd = fnNotNullN(m_Document("NonRiportaInXMLRifVsNumOrd").Value)
        DataUltimaSincronizzazioneContattiFeed = fnNotNull(m_Document("DataUltimaSincronizzazioneContattiFeed").Value)
        DataUltimaSincronizzazioneLottoFeed = fnNotNull(m_Document("DataUltimaSincronizzazioneLottoFeed").Value)
        AttivaGestioneUltimaSincronizzazioneFeed = fnNotNullN(m_Document("AttivaGestioneUltimaSincronizzazioneFeed").Value)
        NonRiportareRifCerticatoInDDT = fnNotNullN(m_Document("NonInviareRifCertificatoIdDDT").Value)
        
        AttivaControlloEsistenzaLottiInFeedAutInSync = fnNotNullN(m_Document("AttivaControlloEsistenzaLottiInFeedAutInSync").Value)
        NonEliminareLottiDefinitivamenteDaFeed = fnNotNullN(m_Document("NonEliminareLottiDefinitivamenteDaFeed").Value)
        NonEliminareLottiProvvDefinitivamenteDaFeed = fnNotNullN(m_Document("NonEliminareLottiProvvDefinitivamenteDaFeed").Value)
        NonAggiornareRifLottoInFeed = fnNotNullN(m_Document("NonAggiornareRifLottoInFeed").Value)
        
        DBNameRegFatture = fnNotNull(m_Document("DBNameRegFatture").Value)
        AttivaMappaturaDaRegFatture = fnNotNullN(m_Document("AttivaMappaturaDaRegFatture").Value)
        
        GET_PAR_CONN_ALY
        
    Else
        Me.cboFiliale.Enabled = True
        Link_ContoPDC = 0
        Link_ContoPDC_RigaNeg = 0
        Link_ContoPDC_RigaPos = 0
        PAR_NonCalcImpDaAssVeloce = 0
        PAR_CalcImpAConfOrd = 0
        PAR_NonVisMsgImpZeroConfOrd = 0
        PAR_AssNewPedDaAssSingola = 0
        PAR_NonCalcPrezzoDueRefArtOrd = 0
        ATT_MULT_SEL_IV_GAMMA = 0
        CHIUDI_CONF_QTASEL_ZERO = 0
        GG_DATA_SCADENZA_CONTR = 0
        FOCUS_IVGAMMA_LAV = 0
        FOCUS_IVGAMMA_CONF = 0
        LINK_UM_COOP_SEL_AUT_IVGAMMA = 0
        PAR_VIS_NOTE_RIGA_ORD_ELENCO = 0
        PAR_VIS_NOTE_RIGA_ORD = 0
        DATA_SCADENZA_OBBL = 0
        PAR_IDARTICOLO_BOLLO = 0
        PAR_IDSOCIO_RISC_PESO = 0
        GtinMigros = ""
        UrlMigros = ""
        NomeUtenteMigros = ""
        PasswordMigros = ""
        ChiaveFeed = ""
        UrlFeed = ""
        AttivaFeed = 0
        ATTIVA_PROD_CONF_LAV = 0
        ATTIVA_PROD_ORD_LAV = 0
        DISATTIVA_PROD_ORD_TESTA = 0
        ATTIVA_PROD_CALC_COLLI_PED = 0
        ATTIVA_SEL_SUB_LOTTO_PROD = 0
        
        ATTIVA_SUB_CONTATTI_FEED = 0
        ATTIVA_SEL_MULT_LOTTI_FEED = 0
        ATTIVA_PROD_FIFO_ACCESSI_CONF = 0
        
        CodiceAnnullaOperazione = ""
        CodiceConfermaOrdine = ""
        CodiceGestioneErrori = ""
        CodiceConfermaViaggio = ""
        
        ATTIVA_GESTIONE_CARICO_MERCE = 0
        NonVisMsgCreaListaPrelAutPedUscita = 0
        ATTIVA_VisAndGrigliaConferimento = 0
        ATTIVA_VisAndGrigliaRigaOdine = 0
        ATTIVA_RICALCOLO_COLLI_PREL_CHIUSO = 0
        IDTipoCollegRigaOrdRigaConf = 0
        NonVisRigheOrdiniComplInSelLav = 0
        ModificaLavorazioneSenzaRicalcolo = 0
        RISCONTRO_PESO_PER_CONF = 0
        ATTIVA_ELIMINAZIONE_ACCESSI = 0
        IDSOCIO_PRE_CONF = 0
        ATTIVA_COMMISSIONI_ORDINI = 0
        RIC_COMM_TIPO_PED_EVAS_ORD = 0
        VIS_ELENCO_RIGHE_ORD = 0
        ATTIVA_CALCOLO_N_PED_BI_VENDITE = 0
        RISCONTRO_PESO_VAL_SOCIO = 0
        RISCONTRO_PESO_VAL_FORNITORE = 0
        OBBL_N_DOC_PES_CONF = 0
        ATTIVA_OBBL_SFALCIO_CONF = 0
        ATTIVA_SEQ_SFALCIO = 0
        RIPORTA_RIF_DETTAGLIO_XML_NC = 0
        RIPORTA_RIF_DETTAGLIO_XML_ND = 0
        CONF_AUT_CONTR_PRESA_VISIONE = 0
        NO_RIP_AGENTE_IN_DOC_EVASIONE = 0
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
        
        COMPANY_CODE = 0
        ATTIVA_FIDO_ALY = 0
        DISATTIVA_DDT_FIDO_ALY = 0
        DISATTIVA_FD_FIDO_ALY = 0
        DISATTIVA_FA_FIDO_ALY = 0
        DISATTIVA_NC_FIDO_ALY = 0
        DISATTIVA_ND_FIDO_ALY = 0
        DISATTIVA_SCALATA_COMM_TRASP = 0
        ATTIVA_FEED_MULTI_LIVELLO = 0
        
        SincronizzaAutFeed = 0
        IvaArticoloDaDocColl = 0
        LetteraIntentoDaDocColl = 0
        RifComuneDaConfigSocio = 0
        DocAnaDestUgualeAnaCoop = 0
        IDClassLottoProdPerFuoriQuota = 0
        MsgInDocSeRigaMerceSenzaImballo = 0
        CarettareClassLottoProd = ""
        CarettareSeparazioneClassLottoProd = ""
        IDAnagraficaDestinazionePerCertificato = 0
        
        IDCategoriaAnagraficaSocioDiretto = 0
        IDCategoriaAnagraficaProdAcq = 0
        IDCategoriaAnagraficaNoProd = 0
        IDArticoloScartoPerCertificato = 0
        
        RiportaDestinazioneDaContrattoCertificato = 0
        RiportaVettoreDaContrattoCertificato = 0
        ForzaDestinazioneDaContrattoCertificato = 0
        ForzaVettoreDaContrattoCertificato = 0
        NonInviareCodiceForInFeedentity = 0
        PrendiVarietaDaTipologiaFeedentity = 0
        AttivaSelezioneSocioCertPerVarieta = 0
        AttivaSelezioneAnaVeloceInCert = 0
        AttivaPaginazioneContattiFeed = 0
        NumeroElementiContattiPerPagina = 0
        AttivaRicercaFatturaAccontoBIVendite = 0
        AttivaRicercaFatturaAccontoBIFatturato = 0
        NumeroMesiPerDataRevocaCertificato = 0
        NonRiportaInXMLRifVsNumOrd = 0
        
        DataUltimaSincronizzazioneContattiFeed = ""
        DataUltimaSincronizzazioneLottoFeed = ""
        AttivaGestioneUltimaSincronizzazioneFeed = 0
        NonRiportareRifCerticatoInDDT = 0
        
        AttivaControlloEsistenzaLottiInFeedAutInSync = 0
        NonEliminareLottiDefinitivamenteDaFeed = 0
        NonEliminareLottiProvvDefinitivamenteDaFeed = 0
        NonAggiornareRifLottoInFeed = 0
        DBNameRegFatture = ""
        AttivaMappaturaDaRegFatture = 0
    End If

 
    RECUPERO_PAR_QUAL_AZ TheApp.IDFirm
    'rif11 end
    
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



End Sub
Public Sub ConnessioneDiamanteADO()
On Error GoTo ERR_ConnessioneDiamanteADO
    '------------------------------
    'APERTURA DELLA CONNESSIONE
    '------------------------------
    
    'Leggiamo il tipo di database utilizzato (Access o SQL Server)
    'Apriamo la connessione in base al tipo di database rilevato
    '(MenuOptions.DBType restituisce il valore del DBType)
    'Select Case MenuOptions.DBType
    '    Case 0 'CONNESSIONE_SQL_SERVER            'Microsoft SQL Server
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=sa;PWD=")
    '    Case 1 'CONNESSIONE_ACCESS               'Microsoft ACCESS
    '        Set Cn = adoEngine.adoEnvironments(0).OpenConnection("", , , "DSN=Diamante;UID=admin;PWD=dmt192981046")
    '    Case -1
            'Se la voce DBType non viene trovata nel file di registro
            'vuol dire che Diamante non è stato installato correttamente
    '        MsgBox "Impossibile avviare il programma. Diamante non è stato installatto correttamente!", vbCritical, "Aggiornamento scadenze"
    '        End
    'End Select
    
    Set Cn = m_App.Database.Connection
    
Exit Sub
ERR_ConnessioneDiamanteADO:
    MsgBox Err.Description, vbCritical, "Connessione Diamante di tipo ADO"
End Sub



Private Sub m_DocumentsLink1_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink1.BOF And m_DocumentsLink1.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.cboDocumentoCoop.WriteOn IIf(IsNull(m_DocumentsLink1("IDDocumentoCoop").Value), 0, m_DocumentsLink1("IDDocumentoCoop").Value)
        Me.cboTipoProcesso.WriteOn IIf(IsNull(m_DocumentsLink1("IDTipoProcessoCoop").Value), 0, m_DocumentsLink1("IDTipoProcessoCoop").Value)
        Me.cboTipoMagazzino.WriteOn IIf(IsNull(m_DocumentsLink1("IDTipoMagazzino").Value), 0, m_DocumentsLink1("IDTipoMagazzino").Value)
        Me.cboCausaliMagazzino.WriteOn fnNotNullN(m_DocumentsLink1("IDFunzione").Value)
        
        

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
        Me.cboDocumentoCoop.WriteOn 0
        Me.cboTipoProcesso.WriteOn 0
        Me.cboTipoMagazzino.WriteOn 0
        Me.cboCausaliMagazzino.WriteOn 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
        Me.cboDocumentoCoop.Enabled = bValue
        Me.cboTipoProcesso.Enabled = bValue
        Me.cboTipoMagazzino.Enabled = bValue
        Me.cboCausaliMagazzino.Enabled = bValue
        
  
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
        If m_DocumentsLink1("IDRV_POProcessoPerDocumento").Value > 0 Then
            Me.cboDocumentoCoop.Enabled = False
        Else
            Me.cboDocumentoCoop.Enabled = True
        End If
        
        
        Me.cmdNuovo_Processo.Enabled = True
        Me.cmdSalva_Processo.Enabled = bValue
        Me.cmdElimina_Processo.Enabled = bValue
End Sub
Private Function EsistenzaFiliale() As Boolean
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    sSQL = "SELECT IDFiliale FROM RV_POSchemaCoop WHERE ("
    sSQL = sSQL & "(IDFiliale=" & Me.cboFiliale.CurrentID & ") AND "
    sSQL = sSQL & "(IDUtente=0))"
    
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        EsistenzaFiliale = True
    Else
        EsistenzaFiliale = False
    End If
    
    rs.CloseResultset
    Set rs = Nothing
End Function

Private Sub m_DocumentsLink2_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink2.BOF And m_DocumentsLink2.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.cboSezionale.WriteOn IIf(IsNull(m_DocumentsLink2("IDSezionale").Value), 0, m_DocumentsLink2("IDSezionale").Value)
        Me.cboDocumenti_PerSezionale.WriteOn IIf(IsNull(m_DocumentsLink2("IDDocumentoCoop").Value), 0, m_DocumentsLink2("IDDocumentoCoop").Value)
        Me.chkPredefinito_Sezionale.Value = fnNormBoolean(IIf(IsNull(m_DocumentsLink2("Predefinito").Value), 0, m_DocumentsLink2("Predefinito").Value))
        
        
        

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
        Me.cboSezionale.WriteOn 0
        Me.cboDocumenti_PerSezionale.WriteOn 0
        Me.chkPredefinito_Sezionale.Value = 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
        
        Me.cboSezionale.Enabled = bValue
        Me.cboDocumenti_PerSezionale.Enabled = bValue
        Me.chkPredefinito_Sezionale.Enabled = bValue
        
        
 
  
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovo_Sezionale.Enabled = True
        Me.cmdSalva_Sezionale.Enabled = bValue
        Me.cmdElimina_Sezionale.Enabled = bValue
End Sub

Private Sub m_DocumentsLink3_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink3.BOF And m_DocumentsLink3.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.cboProtocolloICE.WriteOn IIf(IsNull(m_DocumentsLink3("IDRV_POProtocolloICE").Value), 0, m_DocumentsLink3("IDRV_POProtocolloICE").Value)
        Me.txtNumeroProgressivo.Value = IIf(IsNull(m_DocumentsLink3("Progressivo").Value), 1, m_DocumentsLink3("Progressivo").Value)
        Me.cboEsercizioICE.WriteOn IIf(IsNull(m_DocumentsLink3("IDEsercizio").Value), 0, m_DocumentsLink3("IDEsercizio").Value)
        Me.chkPredefinito_ICE.Value = fnNormBoolean(IIf(IsNull(m_DocumentsLink3("Predefinito").Value), 0, fnNormBoolean(m_DocumentsLink3("Predefinito").Value)))
        
        

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
        Me.cboProtocolloICE.WriteOn 0
        Me.txtNumeroProgressivo.Value = 1
        Me.cboEsercizioICE.WriteOn 0
        Me.chkPredefinito_ICE.Value = 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
      
        Me.cboProtocolloICE.Enabled = bValue
        Me.txtNumeroProgressivo.Enabled = bValue
        Me.cboEsercizioICE.Enabled = bValue
        Me.chkPredefinito_ICE.Enabled = bValue
        
 
        If m_DocumentsLink3("IDRV_POProgProtocolloICE").Value > 0 Then
            Me.cboProtocolloICE.Enabled = False
        Else
            Me.cboProtocolloICE.Enabled = True
        End If
        
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovoProtocolloICE.Enabled = True
        Me.cmdSalvaProtocolloICE.Enabled = bValue
        Me.cmdEliminaProtocolloICE.Enabled = bValue
End Sub
Private Function ControlloPredefinitoICE(ID As Long) As Boolean
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    sSQL = "SELECT Predefinito FROM RV_POProgProtocolloICE "
    sSQL = sSQL & "WHERE Predefinito=" & fnNormBoolean(1)
    sSQL = sSQL & " AND IDRV_POSchemaCoop=" & fnNotNullN(m_Document("IDRV_POSchemaCoop").Value)
    sSQL = sSQL & " AND IDRV_POProgProtocolloICE<>" & ID
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        ControlloPredefinitoICE = True
    Else
        ControlloPredefinitoICE = False
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
    
End Function
Private Function ControlloPredefinitoSezionale() As Boolean
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    
    sSQL = "SELECT Predefinito FROM RV_POSezionalePerDocumento WHERE ("
    sSQL = sSQL & "(Predefinito=" & fnNormBoolean(1) & ") AND "
    sSQL = sSQL & "(IDDocumentoCoop=" & Me.cboDocumenti_PerSezionale.CurrentID & ") AND "
    sSQL = sSQL & "(IDRV_POSchemaCoop=" & m_Document("IDRV_POSchemaCoop").Value & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF Then
        ControlloPredefinitoSezionale = True
    Else
        ControlloPredefinitoSezionale = False
    End If
    
    rs.CloseResultset
    Set rs = Nothing
    
End Function

Private Sub m_DocumentsLink4_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink4.BOF And m_DocumentsLink4.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.cboDocCoopPerOpe.WriteOn fnNotNullN(m_DocumentsLink4("IDRV_PODocumentoCoop"))
        Me.chkGestioneArticoli.Value = fnNormBoolean(IIf(IsNull(m_DocumentsLink4("GestioneArticoli").Value), 0, fnNormBoolean(m_DocumentsLink4("GestioneArticoli").Value)))
        Me.chkCreazioneAutomatica.Value = fnNormBoolean(IIf(IsNull(m_DocumentsLink4("CreazioneAutomaticaLottoVend").Value), 0, fnNormBoolean(m_DocumentsLink4("CreazioneAutomaticaLottoVend").Value)))
        Me.cboTipoLavorazioneAut.WriteOn fnNotNullN(m_DocumentsLink4("IDRV_POTipoLavorazione"))
        

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
        Me.cboDocCoopPerOpe.WriteOn 0
        Me.chkGestioneArticoli.Value = vbUnchecked
        Me.chkCreazioneAutomatica.Value = vbUnchecked
        Me.cboTipoLavorazioneAut.WriteOn 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
      

        Me.cboDocCoopPerOpe.Enabled = bValue
        Me.chkGestioneArticoli.Enabled = bValue
        Me.chkCreazioneAutomatica.Enabled = bValue
        Me.cboTipoLavorazioneAut.Enabled = bValue
        
        If m_DocumentsLink4(m_DocumentsLink4.PrimaryKey).Value > 0 Then
            Me.cboDocCoopPerOpe.Enabled = False
        Else
            Me.cboDocCoopPerOpe.Enabled = True
        End If
        
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovo_OperazionePerDoc.Enabled = True
        Me.cmdSalva__OperazionePerDoc.Enabled = bValue
        Me.cmdElimina_OperazionePerDoc.Enabled = bValue

End Sub


Private Sub m_DocumentsLink6_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink6.BOF And m_DocumentsLink6.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.txtDataInizioInteressi.Value = fnNotNullN(m_DocumentsLink6("DataInizio").Value)
        Me.txtTassoInteressi.Value = fnNotNullN(m_DocumentsLink6("TassoInteresse").Value)
        Me.txtNumeroGiorniAnno.Value = fnNotNullN(m_DocumentsLink6("NumeroGiorniAnno").Value)
        Me.txtPercentualeGiorno.Value = fnNotNullN(m_DocumentsLink6("TassoGiornaliero").Value)
        
        

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
        Me.txtDataInizioInteressi.Value = 0
        Me.txtTassoInteressi.Value = 0
        Me.txtNumeroGiorniAnno.Value = 0
        Me.txtPercentualeGiorno.Value = 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
      
        Me.txtDataInizioInteressi.Enabled = bValue
        Me.txtTassoInteressi.Enabled = bValue
        Me.txtNumeroGiorniAnno.Enabled = bValue
        'Me.txtPercentualeGiorno.Enabled = bValue
        


        
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovo_Interessi.Enabled = True
        Me.cmdSalva_Interessi.Enabled = bValue
        Me.cmdElimina_Interessi.Enabled = bValue
End Sub

Private Sub m_DocumentsLink7_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink7.BOF And m_DocumentsLink7.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.cboUtenteOrd.WriteOn fnNotNullN(m_DocumentsLink7("IDUtente").Value)
        Me.txtIDOrdine.Value = fnNotNullN(m_DocumentsLink7("IDOggettoOrdinePrep").Value)
        Me.txtIDOrdineSmist.Value = fnNotNullN(m_DocumentsLink7("IDOggettoOrdineSmist").Value)
        Me.chkBloccoOrdPrep.Value = Abs(fnNotNullN(m_DocumentsLink7("BloccaOrdinePrep").Value))
        

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
        Me.cboUtenteOrd.WriteOn 0
        Me.txtIDOrdine.Value = 0
        Me.txtIDOrdineSmist.Value = 0
        Me.chkBloccoOrdPrep.Value = 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
      
        Me.cboUtenteOrd.Enabled = bValue
        Me.chkBloccoOrdPrep.Enabled = bValue
        'Me.txtPercentualeGiorno.Enabled = bValue
        


        
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
        Me.cmdNuovoUtenteEva.Enabled = True
        Me.cmdSalvaUtenteEva.Enabled = bValue
        Me.cmdEliminaUtenteEva.Enabled = bValue

End Sub

Private Sub MsgConfNeg_Click()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtCodiceAssociato_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtCodiceConto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtCodiceConto_RigaNeg_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtCodiceConto_RigaPos_Change()
If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtContributi_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDataInizioInteressi_LostFocus()
Dim DataInizio As String
Dim DataFine As String

If Me.txtDataInizioInteressi.Value > 0 Then
    If Me.txtNumeroGiorniAnno.Value = 0 Then
        DataInizio = "01/01/" & Year(Me.txtDataInizioInteressi.Text)
        DataFine = "31/12/" & Year(Me.txtDataInizioInteressi.Text)
        
        Me.txtNumeroGiorniAnno.Value = DateDiff("d", DataInizio, DataFine) + 1
    End If
End If
End Sub

Private Sub txtDataOrdinePred_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtDescrizioneConto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDescrizioneConto_RigaNeg_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtDescrizioneConto_RigaPos_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIDOrdine_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDOggetto, Doc_numero, Doc_data, Link_nom_anagrafica "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & " WHERE IDOggetto=" & Me.txtIDOrdine.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.cdCliente.Load 0
    Me.txtDataOrdine.Value = 0
    Me.txtNumeroOrdine.Value = 0
Else
    If Me.txtIDOrdine.Value > 0 Then
        Me.cdCliente.Load fnNotNullN(rs!Link_nom_anagrafica)
        Me.txtDataOrdine.Value = fnNotNullN(rs!Doc_data)
        Me.txtNumeroOrdine.Value = fnNotNullN(rs!Doc_numero)
    Else
        Me.cdCliente.Load 0
        Me.txtDataOrdine.Value = 0
        Me.txtNumeroOrdine.Value = 0
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub txtIDOrdineIVGamma_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

On Error Resume Next
sSQL = "SELECT * FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & txtIDOrdineIVGamma.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.CDClienteOrdIVGamma.Load 0
    Me.txtDataOrdIVGamma.Value = 0
    Me.txtNOrdIVGamma.Value = 0
    Me.txtNListaOrdIVGamma = 0
    
Else
    Me.CDClienteOrdIVGamma.Load fnNotNullN(rs!Link_nom_anagrafica)
    Me.txtDataOrdIVGamma.Text = fnNotNull(rs!RV_PODataOrdinePadre)
    Me.txtNOrdIVGamma.Value = fnNotNullN(rs!RV_PONumeroOrdinePadre)
    Me.txtNListaOrdIVGamma = fnNotNullN(rs!RV_PONumeroListaPrelievo)
    
End If

rs.CloseResultset
Set rs = Nothing

If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIDOrdineIVGammaLav_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


On Error Resume Next
sSQL = "SELECT * FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & "WHERE IDOggetto=" & Me.txtIDOrdineIVGammaLav.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.CdClienteIVGammaLav.Load 0
    Me.txtDataOrdIVGammaLav.Value = 0
    Me.txtNOrdIVGammaLav.Value = 0
    Me.txtNOrdListaIVGammaLav = 0
Else
    Me.CdClienteIVGammaLav.Load fnNotNullN(rs!Link_nom_anagrafica)
    Me.txtDataOrdIVGammaLav.Text = fnNotNull(rs!RV_PODataOrdinePadre)
    Me.txtNOrdIVGammaLav.Value = fnNotNullN(rs!RV_PONumeroOrdinePadre)
    Me.txtNOrdListaIVGammaLav = fnNotNullN(rs!RV_PONumeroListaPrelievo)
End If

rs.CloseResultset
Set rs = Nothing
    
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtIDOrdineSmist_Change()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT IDOggetto, Doc_numero, Doc_data, Link_nom_anagrafica "
sSQL = sSQL & "FROM ValoriOggettoPerTipo000F "
sSQL = sSQL & " WHERE IDOggetto=" & Me.txtIDOrdineSmist.Value

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.CDClienteSmist.Load 0
    Me.txtDataOrdineSmist.Value = 0
    Me.txtNumeroOrdineSmist.Value = 0
    
Else
    If Me.txtIDOrdineSmist.Value > 0 Then
        Me.CDClienteSmist.Load fnNotNullN(rs!Link_nom_anagrafica)
        Me.txtDataOrdineSmist.Value = fnNotNullN(rs!Doc_data)
        Me.txtNumeroOrdineSmist.Value = fnNotNullN(rs!Doc_numero)
    Else
        Me.CDClienteSmist.Load 0
        Me.txtDataOrdineSmist.Value = 0
        Me.txtNumeroOrdineSmist.Value = 0
    End If
End If

rs.CloseResultset
Set rs = Nothing
End Sub

Private Sub txtIscrizioneAlboCoop_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNListaOrdinePred_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumerazioneLotto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumerazioneLottoConf_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumeroEtichettePedana_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtNumeroGiorniAnno_LostFocus()
If Me.txtNumeroGiorniAnno.Value > 0 Then
    Me.txtPercentualeGiorno.Value = Me.txtTassoInteressi.Value / Me.txtNumeroGiorniAnno.Value
End If
End Sub

Private Sub txtNumeroOrdinePred_Change()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub txtParametroBNDOO_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPercorsoEtiPDF_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtQtaMinimaPerConferimento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtQtaMinimaPerVendita_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Function GetPianoDeiConti() As Long
On Error GoTo ERR_GetPianoDeiConti
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT IDPianoDeiConti FROM PianoDeiConti WHERE ("
    sSQL = sSQL & "(IDAzienda = " & m_App.Branch & ") AND "
    sSQL = sSQL & "(IDEsercizio= " & VarIDEsercizio & "))"
    
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

Private Sub SetPDCProperties()
    Set oPDC = New DmtPDC.PDCServices
    'Imposta le proprietà dell'oggetto PDCServices
    With oPDC
        'Viene fornita al controllo la connessione al database DMT.
        'La connessione è di tipo ADO.Connection quindi viene
        'passata la proprietà InternalConnection dell'oggetto Database
        Set .Connection = m_App.Database.InternalConnection
        'Indica l'identificativo del Piano dei conti da visualizzare
        .IDPDC = fnNotNullN(Link_PianoDeiConti)
        .HideAccounts = False
        .BranchType = btcAllBranchs
        '.BranchType = .BranchType + btcRevenuesBranch
        .AccountType = atcAllAccounts
        
    End With
End Sub
Private Sub ShowNodeProperties(ByVal oNode As DmtPDC.INode, Index As Integer)
    'Rappresenta un conto
    Dim oAccount As DmtPDC.Account
    'Rappresenta un ramo
    Dim oBranch As DmtPDC.Branch
    
    'Vengono visualizzati nei campi appositi tutte
    'le caratteristiche del conto o ramo selezionato
        
    'Controlla se è stato passato un elemento valido
    If Not oNode Is Nothing Then
        'Riporta i dati comuni del conto o del ramo
        If Index = 4 Then
            'Identificativo unico del Conto o del Ramo
            Link_ContoPDC = oNode.ID
            'Codifica completa del Conto o del Ramo
            Me.txtCodiceConto.Text = oNode.CompletedCode
            Me.txtDescrizioneConto.Text = oNode.Description
        End If
        If Index = 1 Then
            'Identificativo unico del Conto o del Ramo
            Link_ContoPDC_RigaNeg = oNode.ID
            'Codifica completa del Conto o del Ramo
            Me.txtCodiceConto_RigaNeg = oNode.CompletedCode
            Me.txtDescrizioneConto_RigaNeg.Text = oNode.Description
        End If
        If Index = 0 Then
            'Identificativo unico del Conto o del Ramo
            Link_ContoPDC_RigaNeg = oNode.ID
            'Codifica completa del Conto o del Ramo
            Me.txtCodiceConto_RigaPos = oNode.CompletedCode
            Me.txtDescrizioneConto_RigaPos.Text = oNode.Description
        End If

    End If
End Sub

Private Sub txtRigaFinalePerFattura_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtRigaInizialePerFattura_Change()
If Not (BrwMain.Visible) Then Change
End Sub
Private Function fncTrovaIDFunzione(Gestore As String) As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

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

Private Sub txtTassoInteressi_LostFocus()
If Me.txtNumeroGiorniAnno.Value > 0 Then
    Me.txtPercentualeGiorno.Value = Me.txtTassoInteressi.Value / Me.txtNumeroGiorniAnno.Value
End If
End Sub
Private Function GET_ESISTENZA_UTENTE_EVA(IDUtente As Long, IDAzienda As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POParametriUtenteOrd "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda
sSQL = sSQL & " AND IDUtente=" & IDUtente

Set rs = Cn.OpenResultset(sSQL)
If rs.EOF Then
    GET_ESISTENZA_UTENTE_EVA = False
Else
    GET_ESISTENZA_UTENTE_EVA = True
End If


rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_CONTROLLO_PROCESSO_FUNZIONE(IDFunzione As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDProcessoPerFunzione "
sSQL = sSQL & "FROM ProcessoPerFunzione "
sSQL = sSQL & "WHERE IDFunzione=" & IDFunzione
sSQL = sSQL & " AND IDProcesso=528"

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CONTROLLO_PROCESSO_FUNZIONE = True
Else
    GET_CONTROLLO_PROCESSO_FUNZIONE = False
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
Private Function fncTrovaIDFunzioneArticoli(Gestore As String, Optional Funzione As String) As Long
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
    fncTrovaIDFunzioneArticoli = fnNotNullN(rs!IDFunzione)
Else
    fncTrovaIDFunzioneArticoli = 0
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

Private Sub SET_PAR_CONN_ALY()
On Error GoTo ERR_SET_PAR_CONN_ALY
Dim rs As ADODB.Recordset
Dim sSQL As String

If CONFERMA_PARAMETRI_ALY = False Then Exit Sub

sSQL = "SELECT * FROM RV_POAlyanteParametriConnessione "
sSQL = sSQL & "WHERE IDAzienda=" & TheApp.IDFirm

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If rs.EOF Then
    rs.AddNew
    rs!IDAzienda = TheApp.IDFirm
End If

rs!NomeServer = NOME_SERVER_ALY
rs!DBName = NOME_DB_ALY
rs!DBUser = USER_PROP_SERVER
rs!Password = PWD_USER_PROP
rs!UserGroup = GRUPPO_ALY
rs!ApplicationUser = USER_ALY
rs!ApplicationPassword = PWD_USER_ALY

rs.Update

rs.Close
Set rs = Nothing

CONFERMA_PARAMETRI_ALY = False

Exit Sub
ERR_SET_PAR_CONN_ALY:
    MsgBox Err.Description, vbCritical, "SET_PAR_CONN_ALY"
End Sub
Private Sub RECUPERO_PAR_QUAL_AZ(IDAzienda As Long)
On Error GoTo ERR_RECUPERO_PAR_QUAL
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

QUAL01 = 0
QUAL02 = 0
QUAL03 = 0
QUAL04 = 0
QUAL05 = 0
QUAL06 = 0
QUAL07 = 0
QUAL08 = 0
QUAL09 = 0
QUAL10 = 0
QUAL11 = 0
QUAL12 = 0
QUAL13 = 0
QUAL14 = 0
QUAL15 = 0
QUAL16 = 0
QUALPRZ16 = 0

sSQL = "SELECT * FROM RV_POParametriQualitaAzienda "
sSQL = sSQL & "WHERE IDAzienda=" & IDAzienda

Set rs = Cn.OpenResultset(sSQL)

If Not rs.EOF Then
    QUAL01 = fnNotNullN(rs!Qualita01)
    QUAL02 = fnNotNullN(rs!Qualita02)
    QUAL03 = fnNotNullN(rs!Qualita03)
    QUAL04 = fnNotNullN(rs!Qualita04)
    QUAL05 = fnNotNullN(rs!Qualita05)
    QUAL06 = fnNotNullN(rs!Qualita06)
    QUAL07 = fnNotNullN(rs!Qualita07)
    QUAL08 = fnNotNullN(rs!Qualita08)
    QUAL09 = fnNotNullN(rs!Qualita09)
    QUAL10 = fnNotNullN(rs!Qualita10)
    QUAL11 = fnNotNullN(rs!Qualita11)
    QUAL12 = fnNotNullN(rs!Qualita12)
    QUAL13 = fnNotNullN(rs!Qualita13)
    QUAL14 = fnNotNullN(rs!Qualita14)
    QUAL15 = fnNotNullN(rs!Qualita15)
    QUAL16 = fnNotNullN(rs!Qualita16)
    QUALPRZ16 = fnNotNullN(rs!QualitaPrezzo16)
End If

rs.CloseResultset
Set rs = Nothing

Exit Sub
ERR_RECUPERO_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "RECUPERO_PAR_QUAL_AZ"
End Sub
Private Sub SALVA_PAR_QUAL(IDAzienda As Long)
On Error GoTo ERR_SALVA_PAR_QUAL
Dim sSQL As String
Dim rs As ADODB.Recordset

If (ELIMINA_PAR_QUAL(IDAzienda) = False) Then Exit Sub

sSQL = "SELECT * FROM RV_POParametriQualitaAzienda "
sSQL = sSQL & " WHERE IDAzienda=" & IDAzienda

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

rs.AddNew
    rs!IDAzienda = IDAzienda
    rs!Qualita01 = QUAL01
    rs!Qualita02 = QUAL02
    rs!Qualita03 = QUAL03
    rs!Qualita04 = QUAL04
    rs!Qualita05 = QUAL05
    rs!Qualita06 = QUAL06
    rs!Qualita07 = QUAL07
    rs!Qualita08 = QUAL08
    rs!Qualita09 = QUAL09
    rs!Qualita10 = QUAL10
    rs!Qualita11 = QUAL11
    rs!Qualita12 = QUAL12
    rs!Qualita13 = QUAL13
    rs!Qualita14 = QUAL14
    rs!Qualita15 = QUAL15
    rs!Qualita16 = QUAL16
    rs!QualitaPrezzo16 = QUALPRZ16
rs.Update

rs.Close
Set rs = Nothing
Exit Sub
ERR_SALVA_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "SALVA_PAR_QUAL"
    
End Sub

Private Function ELIMINA_PAR_QUAL(IDAzienda As Long) As Boolean
On Error GoTo ERR_ELIMINA_PAR_QUAL
Dim sSQL As String
Dim rs As ADODB.Recordset

ELIMINA_PAR_QUAL = False

sSQL = "DELETE FROM RV_POParametriQualitaAzienda "
sSQL = sSQL & " WHERE IDAzienda=" & IDAzienda

Cn.Execute sSQL

ELIMINA_PAR_QUAL = True

Exit Function
ERR_ELIMINA_PAR_QUAL:
    MsgBox Err.Description, vbCritical, "ELIMINA_PAR_QUAL"
End Function
