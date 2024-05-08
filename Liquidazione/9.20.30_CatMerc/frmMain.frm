VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{7A1D73E4-F461-11D0-8F01-004033A00AF2}#1.0#0"; "DmtWheel.ocx"
Object = "{5C67DB53-40E7-11D3-AF44-00105A2FBE61}#11.8#0"; "DmtGridCtl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{910385FB-4687-11D3-935C-00105A2E9BA7}#4.9#0"; "DmtCodDesc.ocx"
Object = "{E0BE4700-0D0C-11D2-B957-002018813989}#10.1#0"; "DMTDataCmb.OCX"
Object = "{2ACC5784-9960-11D1-A947-0040335881DA}#1.0#0"; "DMTDateTime.ocx"
Object = "{E9A7E3D8-0C2C-11D2-B92E-00201880103B}#1.0#0"; "dmteditnum.ocx"
Object = "{9385BB2E-6637-11D1-850D-002018802E11}#3.1#0"; "Dmtsplit.ocx"
Object = "{41B8DADF-1874-4E5A-BB7B-4CE86D43F217}#1.2#0"; "DmtActBox.OCX"
Object = "{FCA49525-5F72-11D2-B9EB-00201880103B}#18.1#0"; "DMTPrinterDialog.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   11580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19065
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
   ScaleHeight     =   11580
   ScaleWidth      =   19065
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin ActiveBar3LibraryCtl.ActiveBar3 BarMenu 
      Height          =   11235
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   19065
      _LayoutVersion  =   2
      _ExtentX        =   33629
      _ExtentY        =   19817
      _DataPath       =   ""
      Bands           =   "frmMain.frx":4781A
      Begin DMTPrinterDialog.DMTDialog DmtPrnDlg 
         Left            =   120
         Top             =   6120
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin DMTSPLIT.DMTSplitBar DMTSplitBar1 
         Height          =   510
         Left            =   2280
         TabIndex        =   55
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
         ScaleWidth      =   19005
         TabIndex        =   57
         Top             =   0
         Width           =   19035
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
            ScaleWidth      =   18705
            TabIndex        =   58
            Top             =   120
            Width           =   18735
            Begin VB.CheckBox chkConguaglio 
               Caption         =   "Conguaglio"
               Enabled         =   0   'False
               Height          =   315
               Left            =   14160
               TabIndex        =   132
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtNomeSocioFatt 
               BackColor       =   &H8000000F&
               Height          =   340
               Left            =   10920
               Locked          =   -1  'True
               TabIndex        =   127
               TabStop         =   0   'False
               Top             =   360
               Width           =   2415
            End
            Begin VB.CommandButton cmdTrattenuteUtilizzate 
               Caption         =   "TRATTENUTE APPLICATE"
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
               Left            =   16080
               TabIndex        =   126
               Top             =   840
               Width           =   2535
            End
            Begin VB.CommandButton cmdRicalcolaTotali 
               Caption         =   "RICALCOLA TOTALI"
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
               Left            =   16080
               TabIndex        =   59
               Top             =   240
               Width           =   2535
            End
            Begin VB.Frame Frame1 
               Height          =   8535
               Left            =   16080
               TabIndex        =   104
               Top             =   2280
               Width           =   2535
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleLordoDocumento 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   105
                  Top             =   1680
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleIvaDocumento 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   106
                  Top             =   1080
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleDocumento 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   107
                  Top             =   480
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenute 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   108
                  Top             =   4080
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtNettoLiquidazione 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   109
                  Top             =   5880
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenutePerLavorazioni 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   110
                  Top             =   2280
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenuteGenerali 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   111
                  Top             =   2880
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenuteAgg 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   112
                  Top             =   4680
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattAggRiep 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   113
                  Top             =   5280
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenuteConferimento 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   123
                  Top             =   3480
                  Width           =   2295
                  _Version        =   65536
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   253
                  Text            =   " 0"
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin VB.Label Label3 
                  Caption         =   "Trattenute conferimento"
                  Height          =   255
                  Index           =   9
                  Left            =   120
                  TabIndex        =   124
                  Top             =   3240
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Trattenute generali"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   122
                  Top             =   2640
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Trattenute per lavorazioni"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   121
                  Top             =   2040
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Totale lordo documento"
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   120
                  Top             =   1440
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Totale I.V.A. documento"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   119
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Importo di liquidazione"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   118
                  Top             =   5640
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Totale trattenute"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   117
                  Top             =   3840
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Totale netto documento"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   116
                  Top             =   240
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Totale tratt. aggiuntive"
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   115
                  Top             =   4440
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Totale tratt. agg. riepilogo"
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   114
                  Top             =   5040
                  Width           =   2295
               End
            End
            Begin VB.TextBox txtSocio 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   0
               Top             =   360
               Width           =   5655
            End
            Begin VB.CheckBox chkPassaggioInFatt 
               Caption         =   "In fattura"
               Height          =   315
               Left            =   14160
               TabIndex        =   5
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txtOggetto 
               Height          =   285
               Left            =   120
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   1560
               Width           =   9375
            End
            Begin VB.CheckBox chkUfficiale 
               Caption         =   "Ufficiale"
               Height          =   315
               Left            =   14160
               TabIndex        =   4
               Top             =   240
               Width           =   1575
            End
            Begin DMTEDITNUMLib.dmtNumber txtNumeroLiquidazione 
               Height          =   315
               Left            =   120
               TabIndex        =   1
               Top             =   960
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Enabled         =   0   'False
               Appearance      =   1
            End
            Begin DMTDATETIMELib.dmtDate txtDataLiquidazione 
               Height          =   315
               Left            =   2400
               TabIndex        =   2
               Top             =   960
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Enabled         =   0   'False
               Appearance      =   1
            End
            Begin DMTDataCmb.DMTCombo cboPeriodoLiquidazione 
               Height          =   315
               Left            =   3840
               TabIndex        =   3
               Top             =   960
               Width           =   5655
               _ExtentX        =   9975
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
            Begin TabDlg.SSTab SSTab1 
               Height          =   8895
               Left            =   120
               TabIndex        =   60
               Top             =   1920
               Width           =   15855
               _ExtentX        =   27966
               _ExtentY        =   15690
               _Version        =   393216
               TabHeight       =   520
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Trattenute elaborate"
               TabPicture(0)   =   "frmMain.frx":479EA
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Label4(16)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Label4(15)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Label4(14)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "Label4(13)"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "Label4(12)"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "Label4(9)"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "Label4(11)"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "Label4(10)"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "Label4(8)"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "Label4(7)"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "Label4(6)"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "Label4(5)"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "Label4(4)"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).Control(13)=   "Label4(3)"
               Tab(0).Control(13).Enabled=   0   'False
               Tab(0).Control(14)=   "Label4(2)"
               Tab(0).Control(14).Enabled=   0   'False
               Tab(0).Control(15)=   "Label4(1)"
               Tab(0).Control(15).Enabled=   0   'False
               Tab(0).Control(16)=   "Label4(0)"
               Tab(0).Control(16).Enabled=   0   'False
               Tab(0).Control(17)=   "Label4(18)"
               Tab(0).Control(17).Enabled=   0   'False
               Tab(0).Control(18)=   "Line4"
               Tab(0).Control(18).Enabled=   0   'False
               Tab(0).Control(19)=   "Label4(19)"
               Tab(0).Control(19).Enabled=   0   'False
               Tab(0).Control(20)=   "Label4(20)"
               Tab(0).Control(20).Enabled=   0   'False
               Tab(0).Control(21)=   "Label4(21)"
               Tab(0).Control(21).Enabled=   0   'False
               Tab(0).Control(22)=   "Label4(22)"
               Tab(0).Control(22).Enabled=   0   'False
               Tab(0).Control(23)=   "Label4(24)"
               Tab(0).Control(23).Enabled=   0   'False
               Tab(0).Control(24)=   "Label4(25)"
               Tab(0).Control(24).Enabled=   0   'False
               Tab(0).Control(25)=   "Label4(26)"
               Tab(0).Control(25).Enabled=   0   'False
               Tab(0).Control(26)=   "Label4(28)"
               Tab(0).Control(26).Enabled=   0   'False
               Tab(0).Control(27)=   "Label4(29)"
               Tab(0).Control(27).Enabled=   0   'False
               Tab(0).Control(28)=   "txtTrattPercLav2"
               Tab(0).Control(28).Enabled=   0   'False
               Tab(0).Control(29)=   "txtTrattPercLav1"
               Tab(0).Control(29).Enabled=   0   'False
               Tab(0).Control(30)=   "txtTrattValLav2"
               Tab(0).Control(30).Enabled=   0   'False
               Tab(0).Control(31)=   "txtTrattValLav1"
               Tab(0).Control(31).Enabled=   0   'False
               Tab(0).Control(32)=   "txtTrattPercGen2"
               Tab(0).Control(32).Enabled=   0   'False
               Tab(0).Control(33)=   "txtTrattPercGen1"
               Tab(0).Control(33).Enabled=   0   'False
               Tab(0).Control(34)=   "txtTrattValGen2"
               Tab(0).Control(34).Enabled=   0   'False
               Tab(0).Control(35)=   "txtTrattValGen1"
               Tab(0).Control(35).Enabled=   0   'False
               Tab(0).Control(36)=   "txtTrattenuteConferimento"
               Tab(0).Control(36).Enabled=   0   'False
               Tab(0).Control(37)=   "GrigliaTrattenuteEla"
               Tab(0).Control(37).Enabled=   0   'False
               Tab(0).Control(38)=   "txtQtaLavorata"
               Tab(0).Control(38).Enabled=   0   'False
               Tab(0).Control(39)=   "txtQtaQuadLav"
               Tab(0).Control(39).Enabled=   0   'False
               Tab(0).Control(40)=   "txtTotaleLavorata"
               Tab(0).Control(40).Enabled=   0   'False
               Tab(0).Control(41)=   "txtQtaVenduta"
               Tab(0).Control(41).Enabled=   0   'False
               Tab(0).Control(42)=   "txtImportoTotaleVenduto"
               Tab(0).Control(42).Enabled=   0   'False
               Tab(0).Control(43)=   "txtTotaleTrattenuta"
               Tab(0).Control(43).Enabled=   0   'False
               Tab(0).Control(44)=   "txtImportoUnitarioVenduto"
               Tab(0).Control(44).Enabled=   0   'False
               Tab(0).Control(45)=   "txtTrattenutaPerLavorazione"
               Tab(0).Control(45).Enabled=   0   'False
               Tab(0).Control(46)=   "txtTrattenuteGenerali"
               Tab(0).Control(46).Enabled=   0   'False
               Tab(0).Control(47)=   "cboTipoLavorazione"
               Tab(0).Control(47).Enabled=   0   'False
               Tab(0).Control(48)=   "txtDifferenza"
               Tab(0).Control(48).Enabled=   0   'False
               Tab(0).Control(49)=   "txtLottoLavorato"
               Tab(0).Control(49).Enabled=   0   'False
               Tab(0).Control(50)=   "txtArticoloLavorato"
               Tab(0).Control(50).Enabled=   0   'False
               Tab(0).Control(51)=   "txtQtaConferita"
               Tab(0).Control(51).Enabled=   0   'False
               Tab(0).Control(52)=   "cmdSalva"
               Tab(0).Control(52).Enabled=   0   'False
               Tab(0).Control(53)=   "cmdElimina"
               Tab(0).Control(53).Enabled=   0   'False
               Tab(0).Control(54)=   "txtLottoConferito"
               Tab(0).Control(54).Enabled=   0   'False
               Tab(0).Control(55)=   "txtArticoloConferito"
               Tab(0).Control(55).Enabled=   0   'False
               Tab(0).Control(56)=   "cboTipoPrezzoMedio"
               Tab(0).Control(56).Enabled=   0   'False
               Tab(0).Control(57)=   "cmdTipoPrezzoMedio"
               Tab(0).Control(57).Enabled=   0   'False
               Tab(0).Control(58)=   "txtOrigineDocumento"
               Tab(0).Control(58).Enabled=   0   'False
               Tab(0).Control(59)=   "Command1"
               Tab(0).Control(59).Enabled=   0   'False
               Tab(0).Control(60)=   "Command2"
               Tab(0).Control(60).Enabled=   0   'False
               Tab(0).ControlCount=   61
               TabCaption(1)   =   "Nuove trattenute"
               TabPicture(1)   =   "frmMain.frx":47A06
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label5(1)"
               Tab(1).Control(1)=   "Label5(0)"
               Tab(1).Control(2)=   "Label4(17)"
               Tab(1).Control(3)=   "Label6(6)"
               Tab(1).Control(4)=   "Label6(7)"
               Tab(1).Control(5)=   "Label4(30)"
               Tab(1).Control(6)=   "cboTipoRicalcoloComm"
               Tab(1).Control(7)=   "cboSegnoTrattenuta"
               Tab(1).Control(8)=   "txtPercentualeTrattenuta"
               Tab(1).Control(9)=   "cboTipoTrattenutaAggiuntiva"
               Tab(1).Control(10)=   "GrigliaNuoveTrattenute"
               Tab(1).Control(11)=   "txtImportoNuovaTrattenuta"
               Tab(1).Control(12)=   "txtNuovaTrattenuta"
               Tab(1).Control(13)=   "cmdEliminaTrattenuta"
               Tab(1).Control(13).Enabled=   0   'False
               Tab(1).Control(14)=   "cmdSalvaTrattenuta"
               Tab(1).Control(15)=   "cmdNuovaTrattenuta"
               Tab(1).ControlCount=   16
               TabCaption(2)   =   "Righe in fattura"
               TabPicture(2)   =   "frmMain.frx":47A22
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Label6(5)"
               Tab(2).Control(1)=   "lblPianodeiDeiConti(4)"
               Tab(2).Control(2)=   "Label6(4)"
               Tab(2).Control(3)=   "Label6(3)"
               Tab(2).Control(4)=   "Label6(2)"
               Tab(2).Control(5)=   "Label6(1)"
               Tab(2).Control(6)=   "Label6(0)"
               Tab(2).Control(7)=   "GridRigheFatt"
               Tab(2).Control(8)=   "cboIva"
               Tab(2).Control(9)=   "txtAliquotaIva"
               Tab(2).Control(10)=   "txtImportoUnitarioFatt"
               Tab(2).Control(11)=   "txtImportoTotaleFatt"
               Tab(2).Control(12)=   "txtQtaRigaFatt"
               Tab(2).Control(13)=   "txtDescrizioneConto"
               Tab(2).Control(13).Enabled=   0   'False
               Tab(2).Control(14)=   "txtCodiceConto"
               Tab(2).Control(14).Enabled=   0   'False
               Tab(2).Control(15)=   "cmdNuovo_RigheFatt"
               Tab(2).Control(16)=   "cmdSalva_RigheFatt"
               Tab(2).Control(17)=   "cmdElimina_RigheFatt"
               Tab(2).Control(17).Enabled=   0   'False
               Tab(2).Control(18)=   "txtDescrizioneRigaFatt"
               Tab(2).ControlCount=   19
               Begin VB.CommandButton Command2 
                  Caption         =   "MODIFICHE"
                  Height          =   435
                  Left            =   13560
                  Style           =   1  'Graphical
                  TabIndex        =   135
                  ToolTipText     =   "Calcolo del prezzo medio"
                  Top             =   1800
                  Width           =   2175
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "TRATTENUTE "
                  Height          =   435
                  Left            =   13560
                  Style           =   1  'Graphical
                  TabIndex        =   134
                  ToolTipText     =   "Calcolo del prezzo medio"
                  Top             =   600
                  Width           =   2175
               End
               Begin VB.TextBox txtOrigineDocumento 
                  Height          =   315
                  Left            =   120
                  Locked          =   -1  'True
                  TabIndex        =   133
                  Top             =   3000
                  Width           =   6375
               End
               Begin VB.CommandButton cmdTipoPrezzoMedio 
                  Caption         =   "PREZZO MEDIO"
                  Height          =   435
                  Left            =   13560
                  Style           =   1  'Graphical
                  TabIndex        =   131
                  ToolTipText     =   "Calcolo del prezzo medio"
                  Top             =   1200
                  Width           =   2175
               End
               Begin DMTDataCmb.DMTCombo cboTipoPrezzoMedio 
                  Height          =   315
                  Left            =   6600
                  TabIndex        =   129
                  Top             =   3000
                  Width           =   6735
                  _ExtentX        =   11880
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
               Begin VB.CommandButton cmdNuovaTrattenuta 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -60600
                  TabIndex        =   40
                  Top             =   2640
                  Width           =   1335
               End
               Begin VB.CommandButton cmdSalvaTrattenuta 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -60600
                  TabIndex        =   39
                  Top             =   3480
                  Width           =   1335
               End
               Begin VB.CommandButton cmdEliminaTrattenuta 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -60600
                  TabIndex        =   41
                  TabStop         =   0   'False
                  Top             =   4440
                  Width           =   1335
               End
               Begin VB.TextBox txtArticoloConferito 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00C00000&
                  Height          =   315
                  Left            =   120
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   4695
               End
               Begin VB.TextBox txtNuovaTrattenuta 
                  Height          =   315
                  Left            =   -74880
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   34
                  Top             =   1140
                  Width           =   6375
               End
               Begin VB.TextBox txtLottoConferito 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   4920
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   4095
               End
               Begin VB.CommandButton cmdElimina 
                  Caption         =   "Elimina"
                  Height          =   495
                  Left            =   13560
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   4560
                  Width           =   2175
               End
               Begin VB.CommandButton cmdSalva 
                  Caption         =   "Salva"
                  Height          =   495
                  Left            =   13560
                  TabIndex        =   32
                  Top             =   3600
                  Width           =   2175
               End
               Begin VB.TextBox txtQtaConferita 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   9120
                  Locked          =   -1  'True
                  TabIndex        =   9
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.TextBox txtArticoloLavorato 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   120
                  Locked          =   -1  'True
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   1240
                  Width           =   4695
               End
               Begin VB.TextBox txtLottoLavorato 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FF0000&
                  Height          =   315
                  Left            =   4920
                  Locked          =   -1  'True
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   1240
                  Width           =   5655
               End
               Begin VB.TextBox txtDescrizioneRigaFatt 
                  Height          =   285
                  Left            =   -74760
                  TabIndex        =   42
                  Top             =   960
                  Width           =   7215
               End
               Begin VB.CommandButton cmdElimina_RigheFatt 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   -60600
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   4440
                  Width           =   1335
               End
               Begin VB.CommandButton cmdSalva_RigheFatt 
                  Caption         =   "Salva"
                  Height          =   375
                  Left            =   -60600
                  TabIndex        =   50
                  Top             =   3480
                  Width           =   1335
               End
               Begin VB.CommandButton cmdNuovo_RigheFatt 
                  Caption         =   "Nuovo"
                  Height          =   375
                  Left            =   -60600
                  TabIndex        =   51
                  Top             =   2640
                  Width           =   1335
               End
               Begin VB.TextBox txtCodiceConto 
                  Height          =   315
                  Left            =   -67440
                  TabIndex        =   48
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1815
               End
               Begin VB.TextBox txtDescrizioneConto 
                  Height          =   315
                  Left            =   -65640
                  TabIndex        =   49
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   3255
               End
               Begin DMTEDITNUMLib.dmtNumber txtDifferenza 
                  Height          =   315
                  Left            =   12240
                  TabIndex        =   11
                  Top             =   720
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  ForeColor       =   12582912
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
                  BorderStyle     =   1
                  Enabled         =   0   'False
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboTipoLavorazione 
                  Height          =   315
                  Left            =   10680
                  TabIndex        =   14
                  Top             =   1245
                  Width           =   2655
                  _ExtentX        =   4683
                  _ExtentY        =   556
                  BackColor       =   16777215
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
               Begin DMTEDITNUMLib.dmtNumber txtTrattenuteGenerali 
                  Height          =   315
                  Left            =   10320
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   1800
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  DecimalPlaces   =   5
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtTrattenutaPerLavorazione 
                  Height          =   315
                  Left            =   8760
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   1800
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  DecimalPlaces   =   5
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtImportoUnitarioVenduto 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   19
                  Top             =   1800
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  DecimalPlaces   =   7
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtQtaRigaFatt 
                  Height          =   315
                  Left            =   -74760
                  TabIndex        =   43
                  Top             =   1560
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
               Begin DMTEDITNUMLib.dmtCurrency txtImportoTotaleFatt 
                  Height          =   315
                  Left            =   -68880
                  TabIndex        =   47
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImportoUnitarioFatt 
                  Height          =   315
                  Left            =   -73800
                  TabIndex        =   44
                  Top             =   1560
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   " 0"
                  BackColor       =   16777215
                  Appearance      =   1
                  CurrencyDecimalPlaces=   5
                  CurrencySymbol  =   ""
                  AllowEmpty      =   0   'False
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtAliquotaIva 
                  Height          =   315
                  Left            =   -69840
                  TabIndex        =   46
                  TabStop         =   0   'False
                  Top             =   1560
                  Width           =   855
                  _Version        =   65536
                  _ExtentX        =   1508
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboIva 
                  Height          =   315
                  Left            =   -72360
                  TabIndex        =   45
                  Top             =   1560
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
               Begin DMTEDITNUMLib.dmtCurrency txtTotaleTrattenuta 
                  Height          =   315
                  Left            =   11880
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   1800
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencyDecimalPlaces=   5
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImportoTotaleVenduto 
                  Height          =   315
                  Left            =   7320
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   1800
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtQtaVenduta 
                  Height          =   315
                  Left            =   4440
                  TabIndex        =   18
                  Top             =   1800
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtTotaleLavorata 
                  Height          =   315
                  Left            =   3000
                  TabIndex        =   17
                  Top             =   1800
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtQtaQuadLav 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   16
                  Top             =   1800
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtNumber txtQtaLavorata 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   15
                  Top             =   1800
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  DecFinalZeros   =   -1  'True
               End
               Begin DMTEDITNUMLib.dmtCurrency txtImportoNuovaTrattenuta 
                  Height          =   315
                  Left            =   -62040
                  TabIndex        =   38
                  Top             =   1140
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _StockProps     =   253
                  BackColor       =   16777215
                  Enabled         =   0   'False
                  Appearance      =   1
                  UseSeparator    =   -1  'True
                  CurrencySymbol  =   ""
                  DecFinalZeros   =   -1  'True
               End
               Begin DmtGridCtl.DmtGrid GrigliaNuoveTrattenute 
                  Height          =   6915
                  Left            =   -74880
                  TabIndex        =   61
                  Top             =   1740
                  Width           =   14175
                  _ExtentX        =   25003
                  _ExtentY        =   12197
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
               Begin DmtGridCtl.DmtGrid GrigliaTrattenuteEla 
                  Height          =   5115
                  Left            =   120
                  TabIndex        =   62
                  Top             =   3360
                  Width           =   13215
                  _ExtentX        =   23310
                  _ExtentY        =   9022
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
               Begin DmtGridCtl.DmtGrid GridRigheFatt 
                  Height          =   6735
                  Left            =   -74880
                  TabIndex        =   63
                  Top             =   2040
                  Width           =   14175
                  _ExtentX        =   25003
                  _ExtentY        =   11880
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
               Begin DMTDataCmb.DMTCombo cboTipoTrattenutaAggiuntiva 
                  Height          =   315
                  Left            =   -68400
                  TabIndex        =   35
                  Top             =   1140
                  Width           =   3015
                  _ExtentX        =   5318
                  _ExtentY        =   556
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
               End
               Begin DMTEDITNUMLib.dmtNumber txtPercentualeTrattenuta 
                  Height          =   315
                  Left            =   -63120
                  TabIndex        =   37
                  Top             =   1140
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
                  DecimalPlaces   =   5
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTDataCmb.DMTCombo cboSegnoTrattenuta 
                  Height          =   315
                  Left            =   -63960
                  TabIndex        =   101
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   735
                  _ExtentX        =   1296
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
               Begin DMTEDITNUMLib.dmtNumber txtTrattenuteConferimento 
                  Height          =   315
                  Left            =   10680
                  TabIndex        =   10
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   1455
                  _Version        =   65536
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _StockProps     =   253
                  Text            =   "0"
                  BackColor       =   16777215
                  BorderStyle     =   1
                  UseSeparator    =   -1  'True
                  DecimalPlaces   =   5
                  DecFinalZeros   =   -1  'True
                  AllowEmpty      =   0   'False
               End
               Begin DMTEDITNUMLib.dmtCurrency txtTrattValGen1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   2355
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
               Begin DMTEDITNUMLib.dmtCurrency txtTrattValGen2 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   2355
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
               Begin DMTEDITNUMLib.dmtCurrency txtTrattPercGen1 
                  Height          =   315
                  Left            =   3240
                  TabIndex        =   26
                  TabStop         =   0   'False
                  Top             =   2400
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
               Begin DMTEDITNUMLib.dmtCurrency txtTrattPercGen2 
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   2400
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
               Begin DMTEDITNUMLib.dmtCurrency txtTrattValLav1 
                  Height          =   315
                  Left            =   6360
                  TabIndex        =   28
                  TabStop         =   0   'False
                  Top             =   2400
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
               Begin DMTEDITNUMLib.dmtCurrency txtTrattValLav2 
                  Height          =   315
                  Left            =   7920
                  TabIndex        =   29
                  TabStop         =   0   'False
                  Top             =   2400
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
               Begin DMTEDITNUMLib.dmtCurrency txtTrattPercLav1 
                  Height          =   315
                  Left            =   9480
                  TabIndex        =   30
                  TabStop         =   0   'False
                  Top             =   2400
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
               Begin DMTEDITNUMLib.dmtCurrency txtTrattPercLav2 
                  Height          =   315
                  Left            =   11040
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   2400
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
               Begin DMTDataCmb.DMTCombo cboTipoRicalcoloComm 
                  Height          =   315
                  Left            =   -65280
                  TabIndex        =   36
                  Top             =   1140
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   556
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
               End
               Begin VB.Label Label4 
                  Caption         =   "Tipo ric. tratt."
                  Height          =   255
                  Index           =   30
                  Left            =   -65280
                  TabIndex        =   148
                  Top             =   900
                  Width           =   2055
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. % lav. 2"
                  Height          =   255
                  Index           =   29
                  Left            =   11040
                  TabIndex        =   147
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. % lav. 1"
                  Height          =   255
                  Index           =   28
                  Left            =   9480
                  TabIndex        =   146
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. Val. lav. 2"
                  Height          =   255
                  Index           =   26
                  Left            =   7920
                  TabIndex        =   143
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. Val. lav. 1"
                  Height          =   255
                  Index           =   25
                  Left            =   6360
                  TabIndex        =   142
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. % gen. 2"
                  Height          =   255
                  Index           =   24
                  Left            =   4800
                  TabIndex        =   141
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. % gen. 1"
                  Height          =   255
                  Index           =   22
                  Left            =   3240
                  TabIndex        =   138
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. Val. gen. 2"
                  Height          =   255
                  Index           =   21
                  Left            =   1680
                  TabIndex        =   137
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. Val. gen. 1"
                  Height          =   255
                  Index           =   20
                  Left            =   120
                  TabIndex        =   136
                  Top             =   2160
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Caption         =   "Tipo prezzo medio"
                  Height          =   255
                  Index           =   19
                  Left            =   6600
                  TabIndex        =   130
                  Top             =   2760
                  Width           =   6375
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H00FF0000&
                  BorderWidth     =   2
                  X1              =   13440
                  X2              =   13440
                  Y1              =   480
                  Y2              =   3240
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. Conf."
                  Height          =   255
                  Index           =   18
                  Left            =   10680
                  TabIndex        =   125
                  Top             =   525
                  Width           =   1455
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "%"
                  Height          =   255
                  Index           =   7
                  Left            =   -63120
                  TabIndex        =   103
                  Top             =   900
                  Width           =   975
               End
               Begin VB.Label Label6 
                  Caption         =   "Segno"
                  Height          =   255
                  Index           =   6
                  Left            =   -63960
                  TabIndex        =   102
                  Top             =   480
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.Label Label4 
                  Caption         =   "Tipo trattenuta"
                  Height          =   255
                  Index           =   17
                  Left            =   -68400
                  TabIndex        =   100
                  Top             =   900
                  Width           =   2175
               End
               Begin VB.Label Label4 
                  Caption         =   "Articolo conferito"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   89
                  Top             =   525
                  Width           =   3975
               End
               Begin VB.Label Label5 
                  Caption         =   "Trattenuta"
                  Height          =   255
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   88
                  Top             =   900
                  Width           =   6375
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Importo"
                  Height          =   255
                  Index           =   1
                  Left            =   -62280
                  TabIndex        =   87
                  Top             =   900
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Caption         =   "Documento di conferimento"
                  Height          =   255
                  Index           =   1
                  Left            =   4920
                  TabIndex        =   86
                  Top             =   525
                  Width           =   3975
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Q.tà conferita"
                  Height          =   255
                  Index           =   2
                  Left            =   9120
                  TabIndex        =   85
                  Top             =   525
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Caption         =   "Articolo lavorato/venduto"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   84
                  Top             =   1040
                  Width           =   3375
               End
               Begin VB.Label Label4 
                  Caption         =   "Lotto lavorato/venduto"
                  Height          =   255
                  Index           =   4
                  Left            =   4920
                  TabIndex        =   83
                  Top             =   1035
                  Width           =   5655
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Q.tà lavorata"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   82
                  Top             =   1605
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Q.tà Q. Lav."
                  Height          =   255
                  Index           =   6
                  Left            =   1560
                  TabIndex        =   81
                  Top             =   1605
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Q.tà Tot. lav. "
                  Height          =   255
                  Index           =   7
                  Left            =   3000
                  TabIndex        =   80
                  Top             =   1600
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Q.tà Venduta"
                  Height          =   255
                  Index           =   8
                  Left            =   4440
                  TabIndex        =   79
                  Top             =   1600
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tot. Imp."
                  Height          =   255
                  Index           =   10
                  Left            =   7320
                  TabIndex        =   78
                  Top             =   1605
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tot. Tratt."
                  Height          =   255
                  Index           =   11
                  Left            =   11880
                  TabIndex        =   77
                  Top             =   1600
                  Width           =   1335
               End
               Begin VB.Label Label6 
                  Caption         =   "Riga di fatturazione"
                  Height          =   255
                  Index           =   0
                  Left            =   -74760
                  TabIndex        =   76
                  Top             =   720
                  Width           =   6135
               End
               Begin VB.Label Label6 
                  Caption         =   "I.V.A."
                  Height          =   255
                  Index           =   1
                  Left            =   -72360
                  TabIndex        =   75
                  Top             =   1320
                  Width           =   2415
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Aliquota"
                  Height          =   255
                  Index           =   2
                  Left            =   -69840
                  TabIndex        =   74
                  Top             =   1320
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Importo unitario"
                  Height          =   255
                  Index           =   3
                  Left            =   -73800
                  TabIndex        =   73
                  Top             =   1320
                  Width           =   1335
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Importo totale"
                  Height          =   255
                  Index           =   4
                  Left            =   -68760
                  TabIndex        =   72
                  Top             =   1320
                  Width           =   1215
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
                  Left            =   -67440
                  MouseIcon       =   "frmMain.frx":47A3E
                  MousePointer    =   99  'Custom
                  TabIndex        =   71
                  Top             =   1320
                  Width           =   5055
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Q.tà"
                  Height          =   255
                  Index           =   5
                  Left            =   -74760
                  TabIndex        =   70
                  Top             =   1320
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Imp. Uni. "
                  Height          =   255
                  Index           =   9
                  Left            =   6000
                  TabIndex        =   69
                  Top             =   1600
                  Width           =   1335
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. per Lav."
                  Height          =   255
                  Index           =   12
                  Left            =   8760
                  TabIndex        =   68
                  Top             =   1605
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tratt. gen."
                  Height          =   255
                  Index           =   13
                  Left            =   10320
                  TabIndex        =   67
                  Top             =   1600
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Caption         =   "Tipo lavorazione"
                  Height          =   255
                  Index           =   14
                  Left            =   10680
                  TabIndex        =   66
                  Top             =   1040
                  Width           =   2535
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Differenza"
                  Height          =   255
                  Index           =   15
                  Left            =   12360
                  TabIndex        =   65
                  Top             =   525
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Caption         =   "Orgine documento:"
                  Height          =   255
                  Index           =   16
                  Left            =   120
                  TabIndex        =   64
                  Top             =   2760
                  Width           =   1695
               End
            End
            Begin DmtGridCtl.DmtGrid BrwMain 
               Height          =   735
               Left            =   6000
               TabIndex        =   90
               Top             =   3720
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
            Begin DMTEDITNUMLib.dmtNumber txtNumeroInterno 
               Height          =   315
               Left            =   1200
               TabIndex        =   98
               Top             =   960
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   556
               _StockProps     =   253
               BackColor       =   16777215
               Enabled         =   0   'False
               Appearance      =   1
            End
            Begin DmtCodDescCtl.DmtCodDesc CDSocioFatt 
               Height          =   615
               Left            =   5880
               TabIndex        =   128
               Top             =   120
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   1085
               PropCodice      =   $"frmMain.frx":47D48
               BeginProperty PropCodiceFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PropDescrizione =   $"frmMain.frx":47D96
               BeginProperty PropDescrizioneFontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MenuFunctions   =   $"frmMain.frx":47DFB
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
            Begin DMTDataCmb.DMTCombo cboListino 
               Height          =   315
               Left            =   9600
               TabIndex        =   149
               Top             =   960
               Width           =   3735
               _ExtentX        =   6588
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
            Begin DMTDataCmb.DMTCombo cboCatMerc 
               Height          =   315
               Left            =   9600
               TabIndex        =   151
               Top             =   1560
               Width           =   3735
               _ExtentX        =   6588
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
            Begin VB.Label Label1 
               Caption         =   "Categoria merceolgica"
               Height          =   255
               Index           =   4
               Left            =   9600
               TabIndex        =   152
               Top             =   1320
               Width           =   3735
            End
            Begin VB.Label Label1 
               Caption         =   "Listino"
               Height          =   255
               Index           =   15
               Left            =   9600
               TabIndex        =   150
               Top             =   720
               Width           =   1335
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   14040
               X2              =   14040
               Y1              =   120
               Y2              =   1680
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   15960
               X2              =   18600
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               X1              =   15840
               X2              =   15840
               Y1              =   120
               Y2              =   1680
            End
            Begin VB.Label Label1 
               Caption         =   "N° Interno"
               Height          =   255
               Index           =   3
               Left            =   1200
               TabIndex        =   99
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblPercentualeIstat 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   96
               Top             =   2640
               Width           =   2055
            End
            Begin VB.Label Label1 
               Caption         =   "Socio"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   95
               Top             =   120
               Width           =   3855
            End
            Begin VB.Label Label2 
               Caption         =   "Periodo di liquidazione"
               Height          =   255
               Index           =   0
               Left            =   3840
               TabIndex        =   94
               Top             =   720
               Width           =   5655
            End
            Begin VB.Label Label1 
               Caption         =   "N° Reg."
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   93
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Data Reg."
               Height          =   255
               Index           =   2
               Left            =   2400
               TabIndex        =   92
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Documento collegato"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   91
               Top             =   1320
               Width           =   9375
            End
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
         Left            =   360
         ScaleHeight     =   4935
         ScaleWidth      =   60
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   60
      End
      Begin DmtActBox.DmtActBoxCtl ActivityBox 
         Height          =   7875
         Left            =   -240
         TabIndex        =   97
         Top             =   120
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
         Left            =   1185
         Top             =   660
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   4695
         Left            =   1080
         MousePointer    =   9  'Size W E
         Top             =   0
         Width           =   60
      End
      Begin VB.Line Line2 
         X1              =   2040
         X2              =   6960
         Y1              =   3360
         Y2              =   3360
      End
   End
   Begin MSComctlLib.StatusBar stbStatusbar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   53
      Top             =   11235
      Width           =   19065
      _ExtentX        =   33629
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin DMTEDITNUMLib.dmtCurrency dmtCurrency3 
      Height          =   315
      Left            =   0
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   195
      Width           =   135
      _Version        =   65536
      _ExtentX        =   238
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
   Begin DMTEDITNUMLib.dmtCurrency dmtCurrency7 
      Height          =   315
      Left            =   0
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   195
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Tratt. Val. gen. 1"
      Height          =   255
      Index           =   27
      Left            =   0
      TabIndex        =   145
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Tratt. Val. gen. 1"
      Height          =   255
      Index           =   23
      Left            =   0
      TabIndex        =   140
      Top             =   0
      Width           =   1695
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
'ELABORAZIONE LIQUIDAZIONE
Private WithEvents m_DocumentsLink As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink.VB_VarHelpID = -1
'TRATTENUTE AGGIUNTIVE
Private WithEvents m_DocumentsLink1 As DmtDocManLib.DocumentsLink
Attribute m_DocumentsLink1.VB_VarHelpID = -1
'RIGHE DI FATTURAZIONE
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


'****************************VARIABILI CONTRATTO**********************************
Public NuovaRata As Integer
Public ALIQUOTA_IVA_PRODOTTO As Double
Public Link_RiferimentoRata As Long
Public Var_Percentuale As Double

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

    PermissionToSave = True
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
    
    
    NewSearch
    
    
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
             OnChangeView sToolName
            'OnNew sToolName
           
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
                    Field.Control.Value = fnNotNullN(m_Document.Fields(Field.Name).Value)
                Case "DmtSearchACS"
                        Field.Control.Description = fnNotNull(m_Document.Fields("Anagrafica").Value)
                        Field.Control.SecondDescription = fnNotNull(m_Document.Fields("Nome").Value)
                        Field.Control.IDAnagrafica = m_Document.Fields(Field.Name).Value
                Case "CheckBox"
                    If fnNotNullN(m_Document.Fields(Field.Name).Value) = 0 Then
                        Field.Control.Value = vbUnchecked
                    Else
                        Field.Control.Value = vbChecked
                    End If
                        
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
    
    
    'Totale trattenute
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleTrattenute
    Field.Name = "TotaleTrattenuta"
    Field.Visible = True
    Me.txtTotaleTrattenute.Tag = Field.Name
    m_FormFields.Add Field
    
    'Netto liquidazione
    Set Field = New FormField
    Set Field.Control = Me.txtNettoLiquidazione
    Field.Name = "NettoLiquidazione"
    Field.Visible = True
    Me.txtNettoLiquidazione.Tag = Field.Name
    m_FormFields.Add Field
    
    'Totale Documento Lordo Iva
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleLordoDocumento
    Field.Name = "TotaleDocumentoLordoIva"
    Field.Visible = True
    Me.txtTotaleLordoDocumento.Tag = Field.Name
    m_FormFields.Add Field
    
    'Totale Iva documento
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleIvaDocumento
    Field.Name = "TotaleIva"
    Field.Visible = True
    Me.txtTotaleIvaDocumento.Tag = Field.Name
    m_FormFields.Add Field
    
    'Totale netto Documento
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleDocumento
    Field.Name = "TotaleDocumento"
    Field.Visible = True
    Me.txtTotaleDocumento.Tag = Field.Name
    m_FormFields.Add Field
    
    'Totale trattenute generali
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleTrattenuteGenerali
    Field.Name = "TotaleTrattenutaGenerale"
    Field.Visible = True
    Me.txtTotaleTrattenuteGenerali.Tag = Field.Name
    m_FormFields.Add Field
    
    'Totale trattenute per lavorazioni
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleTrattenutePerLavorazioni
    Field.Name = "TotaleTrattenutaPerLavorazione"
    Field.Visible = True
    Me.txtTotaleTrattenutePerLavorazioni.Tag = Field.Name
    m_FormFields.Add Field
    
    
    'Flag che indica che la liquidazione è stata fatturata
    Set Field = New FormField
    Set Field.Control = Me.chkPassaggioInFatt
    Field.Name = "PassaggioInFatturazione"
    Field.Visible = True
    Me.chkPassaggioInFatt.Tag = Field.Name
    m_FormFields.Add Field

    'Documento collegato
    Set Field = New FormField
    Set Field.Control = Me.txtOggetto
    Field.Name = "Oggetto"
    Field.Visible = True
    Me.txtOggetto.Tag = Field.Name
    m_FormFields.Add Field

    'Ufficiale
    Set Field = New FormField
    Set Field.Control = Me.chkUfficiale
    Field.Name = "Ufficiale"
    Field.Visible = True
    Me.chkUfficiale.Tag = Field.Name
    m_FormFields.Add Field
    
    'Totale trattenute aggiuntive
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleTrattenuteAgg
    Field.Name = "TotaleTrattenuteAggiuntive"
    Field.Visible = True
    Me.txtTotaleTrattenuteAgg.Tag = Field.Name
    m_FormFields.Add Field
    
    'Totale trattenute aggiuntive nel riepilogo
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleTrattAggRiep
    Field.Name = "TotaleTrattenuteRiepilogo"
    Field.Visible = True
    Me.txtTotaleTrattAggRiep.Tag = Field.Name
    m_FormFields.Add Field

    'Totale trattenute aggiuntive nel riepilogo
    Set Field = New FormField
    Set Field.Control = Me.txtTotaleTrattenuteConferimento
    Field.Name = "TotaleTrattenuteConferimento"
    Field.Visible = True
    Me.txtTotaleTrattenuteConferimento.Tag = Field.Name
    m_FormFields.Add Field
    
    'Conguaglio
    Set Field = New FormField
    Set Field.Control = Me.chkConguaglio
    Field.Name = "Conguaglio"
    Field.Visible = True
    Me.chkConguaglio.Tag = Field.Name
    m_FormFields.Add Field

    'IDListino
    Set Field = New FormField
    Set Field.Control = Me.cboListino
    Field.Name = "IDListino"
    Field.Visible = True
    Me.cboListino.Tag = Field.Name
    m_FormFields.Add Field
    
    'IDAnagraficaFatturazione
    Set Field = New FormField
    Set Field.Control = Me.CDSocioFatt
    Field.Name = "IDAnagraficaFatturazione"
    Field.Visible = True
    Me.CDSocioFatt.Tag = Field.Name
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
    
'**************************SOTTO DOCUMENTI ***********************************************************

    'Crea un sottodocumento basato sulla tabella di cross "RV_POLiquidazioneEla"
    
    Set m_DocumentsLink = m_Document.AddDocumentsLink("RV_POLiquidazioneRigheEla")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink.PrimaryKey = "IDRV_POLiquidazioneRigheEla" '<-- Specifica il campo chiave primaria
            

    m_DocumentsLink.AddOrderedColumn "DataConferimento", ocDescending
    m_DocumentsLink.AddOrderedColumn "NumeroDocumento", ocDescending
    m_DocumentsLink.AddOrderedColumn "TipoRiga", ocAscending
    
    'rif6 end
'*****************************************************************************************************
    
    Set m_DocumentsLink1 = m_Document.AddDocumentsLink("RV_POLiquidazioneRighe")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink1.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink1.PrimaryKey = "IDRV_POLiquidazioneRighe" '<-- Specifica il campo chiave primaria

    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POTipoTrattenutaAggiuntiva", "RV_POTipoTrattenutaAggiuntiva", ltLeft, "IDRV_POTipoTrattenutaAggiuntiva")
    NewLink.AddLinkColumn "RV_POTipoTrattenutaAggiuntiva.TipoTrattenuta"

    Set NewLink = m_DocumentsLink1.AddLink("IDRV_POSegnoTrattenuta", "RV_POSegnoTrattenuta", ltLeft, "IDRV_POSegnoTrattenuta")
    NewLink.AddLinkColumn "RV_POSegnoTrattenuta.SegnoTrattenuta"
    

'*****************************************************************************************************
    Set m_DocumentsLink2 = m_Document.AddDocumentsLink("RV_POLiquidazioneRigheFatt")
    
    'Impostazioni dell'oggetto DocumentsLink
    m_DocumentsLink2.EnableRefreshLinks = True '<-- Abilita il refresh dei campi collegati
    m_DocumentsLink2.PrimaryKey = "IDRV_POLiquidazioneRigheFatt" '<-- Specifica il campo chiave primaria
    
    Set NewLink = m_DocumentsLink2.AddLink("IDIva", "Iva", ltLeft, "IDIva")
    NewLink.AddLinkColumn "Iva.Iva"
'*****************************************************************************************************
    
    
    
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
    m_Document.Dataset.Recordset.Sort = "NumeroLiquidazione DESC, Anagrafica"
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

    BrwMain.Conditions.WidthConditions = 200
    BrwMain.Conditions.WidthFields = 200
    BrwMain.Conditions.WidthIntervals = 100
    
    BrwMain.Title.BackColor = vb3DFace
    BrwMain.Title.ForeColor = vbBlack
    BrwMain.Title.Font.Bold = True


    Set Cond = BrwMain.Conditions.Add("Anagrafica", "Anagrafica", m_DocType.TableName, True, False, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("Codice", "Codice", m_DocType.TableName, True, False, , dgCondTypeText)
    
    Set Cond = BrwMain.Conditions.Add("NumeroLiquidazione", "Numero liquidazione", m_DocType.TableName, True, True, , dgCondTypeNumber)
    Set Cond = BrwMain.Conditions.Add("NumeroProtInt", "Numero interno", m_DocType.TableName, True, True, , dgCondTypeNumber)
    
    Set Cond = BrwMain.Conditions.Add("DataLiquidazione", "Data liquidazione", m_DocType.TableName, False, True, , dgCondTypeDate)
        Cond.RangeChecked = True
        Cond.FromValue = Date - 90
        Cond.ToValue = Date
    Set Cond = BrwMain.Conditions.Add("Ufficiale", "Ufficiale", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
        Cond.FromValue = "SI"
    Set Cond = BrwMain.Conditions.Add("AnagraficaFatturazione", "Anagrafica di fatturazione", m_DocType.TableName, False, False, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("CodiceAnagraficaFatturazione", "Codice di fatturazione", m_DocType.TableName, False, False, , dgCondTypeText)
    Set Cond = BrwMain.Conditions.Add("Conguaglio", "Conguaglio", m_DocType.TableName, False, False, False, dgCondTypeBoolean)
    Set Cond = BrwMain.Conditions.Add("IDCategoriaMerceologica", "Categoria merceologica", m_DocType.TableName, False, False, , dgCondTypeComboDB)
        Cond.Indentation = 20
        Cond.RecordSource = "SELECT * FROM CategoriaMerceologica"
        Cond.DisplayField = "CategoriaMerceologica"
        Cond.KeyField = "IDCategoriaMerceologica"
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
    'm_Report.Copies = 1
    'm_Report.Orientation = ocPortrait
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
        If m_App.Caller = "RV_POMenuGreenTop" Then Exit Function
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
        gResource.CustomStrings.Add Chr(34) & m_App.TableName & Chr(34), 1

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
        sWhere = ""
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
    m_DocType.Fields("IDUtente").Value = m_App.IDUser
    
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
       
   
   
    If Me.GrigliaTrattenuteEla.ColumnsHeader.Count = 0 Then
        With Me.GrigliaTrattenuteEla.ColumnsHeader
            .Add "DataConferimento", "Data conf.", dgDate, True, 1000, 0, True, True, False
            .Add "NumeroDocumento", "N° conf.", dgchar, True, 1000, 0, True, True, False
            .Add "CodiceArticolo", "Cod. Art. Vend.", dgchar, True, 1500, 0, True, True, False
            .Add "Articolo", "Art. vend.", dgchar, True, 2500, 0, True, True, False
            '.Add "CodiceLottoArticolo", "Codice lotto", dgchar, True, 1500, 0, True, True, False
            .Add "CodiceArticolo_Conf", "Codice Art. Conf.", dgchar, True, 1500, 0, True, True, False
            .Add "Articolo_Conf", "Articolo Conf.", dgchar, True, 2500, 0, True, True, False
            '.Add "CodiceLottoArticolo_Conf", "Codice lotto conf", dgchar, False, 1500, 0, True, True, False
            Set cl = .Add("QuantitaLavorata", "Q.tà Lav.", dgDouble, False, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("QuantitaQuadrata", "Q.tà Quad. Lav.", dgDouble, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("QuantitaTotaleLavorata", "Q.tà Tot. lav.", dgDouble, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("QuantitaVenduta", "Q.tà Vend", dgDouble, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("ImponibileDaReg", "Imp. Tot. vend.", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("TrattenutePerLavorazione", "Tratt. per lav.", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("TrattenuteGenerali", "Tot. Tratt.", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("TrattenuteTotali", "Tot. Tratt.", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
                
            Set cl = .Add("ImpUniVendDocNettoVendita", "Imp. uni. vend.", dgCurrency, False, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            
            Set cl = .Add("TrattenutaValorePreLiq1", "Tratt. pre liq. 1", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("TrattenutaValorePreLiq2", "Tratt. pre liq. 2", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
        
        End With
    End If
    Me.GrigliaTrattenuteEla.EnableMove = True
   
   
    If Me.GrigliaNuoveTrattenute.ColumnsHeader.Count = 0 Then
        With Me.GrigliaNuoveTrattenute.ColumnsHeader
            .Add "DescrizioneAggiuntiva", "Trattenuta", dgchar, True, 3000, 0, True, True, False
            .Add "IDRV_POTipoTrattenutaAggiuntiva", "IDRV_POTipoTrattenutaAggiuntiva", dgInteger, False, 500, dgAlignleft
            .Add "TipoTrattenuta", "Tipo trattenuta", dgchar, True, 2000, dgAlignleft, True, True, False
            .Add "IDRV_POSegnoTrattenuta", "IDRV_POTipoTrattenutaAggiuntiva", dgInteger, False, 500, dgAlignleft
            Set cl = .Add("Percentuale", "Percentuale", dgDouble, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 5
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("ImportoTrattenuta", "Importo", dgDouble, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
        End With
    End If
    Me.GrigliaNuoveTrattenute.EnableMove = True
   

    If Me.GridRigheFatt.ColumnsHeader.Count = 0 Then
        With Me.GridRigheFatt.ColumnsHeader
            .Add "DescrizioneRiga", "Descrizione riga", dgchar, True, 3000, 0, True, True, False
            Set cl = .Add("Quantita", "Quantità", dgDouble, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("ImportoUnitario", "Importo Unitario", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            .Add "Iva", "I.V.A.", dgchar, True, 3000, 0, True, True, False
            Set cl = .Add("AliquotaIva", "Aliquota IVA", dgDouble, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."
            Set cl = .Add("ImportoTotale", "Importo totale", dgCurrency, True, 1500, dgAlignRight)
                cl.FormatOptions.FormatNumericRegionalSettings = False
                cl.FormatOptions.UseFormatControlSettings = False
                'cl.FormatOptions.FormatNumericCurSymbol = "€  "
                cl.FormatOptions.FormatNumericDecSep = ","
                cl.FormatOptions.FormatNumericDecimals = 2
                cl.FormatOptions.FormatNumericThousandSep = "."

        End With
    End If
    Me.GridRigheFatt.EnableMove = True

'''''''''''''''''''''''''CONTROLLI STANDARD''''''''''''''''''''''''''''''''''''
    
     
    

    

    'Periodo di liquidazione
    With Me.cboPeriodoLiquidazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POLiquidazionePeriodo"
        .DisplayField = "Periodo"
        .SQL = "SELECT * FROM RV_POLiquidazionePeriodo"
        .Fill
    End With
    
    'IVA
    With Me.cboIva
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDIva"
        .DisplayField = "Iva"
        .SQL = "SELECT * FROM Iva"
        .Fill
    End With
    
    'Tipo lavorazione
    With Me.cboTipoLavorazione
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoLavorazione"
        .DisplayField = "TipoLavorazione"
        .SQL = "SELECT * FROM RV_POTipoLavorazione"
        .Fill
    End With

    'Tipo trattenute aggiuntive
    With Me.cboTipoTrattenutaAggiuntiva
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoTrattenutaAggiuntiva"
        .DisplayField = "TipoTrattenuta"
        .SQL = "SELECT * FROM RV_POTipoTrattenutaAggiuntiva ORDER BY TipoTrattenuta"
        .Fill
    End With

    'Tipo calcolo per commissioni
    With Me.cboTipoRicalcoloComm
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoRicalcoloComm"
        .DisplayField = "TipoRicalcoloComm"
        .SQL = "SELECT * FROM RV_POTipoRicalcoloComm"
        .Fill
    End With
    
    'Segno trattenuta
    With Me.cboSegnoTrattenuta
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POSegnoTrattenuta"
        .DisplayField = "SegnoTrattenuta"
        .SQL = "SELECT * FROM RV_POSegnoTrattenuta"
        .Fill
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
    
    'Tipo di prezzo medio
    With Me.cboTipoPrezzoMedio
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDRV_POTipoPrezzoMedio"
        .DisplayField = "TipoPrezzoMedio"
        .SQL = "SELECT * FROM RV_POTipoPrezzoMedio"
        .Fill
    End With

    With cboListino
        Set .Database = TheApp.Database.Connection
        .AddFieldKey "IDListino"
        .DisplayField = "Listino"
        .SQL = "SELECT IDListino, Listino FROM Listino"
        .SQL = .SQL & " WHERE IDAzienda = " & TheApp.IDFirm
        .SQL = .SQL & " AND TipoListino = 0"
        .SQL = .SQL & " ORDER BY Listino"
    End With
    'Categoria merceologica
    With Me.cboCatMerc
        Set .Database = m_App.Database.Connection
        .AddFieldKey "IDCategoriaMerceologica"
        .DisplayField = "CategoriaMerceologica"
        .SQL = "SELECT * FROM CategoriaMerceologica"
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
                        Field.Value = fnNormBoolean(m_FormFields(Field.Name).Control.Value)
                
                End Select
                
                'rif4 end
                
            Else
                If Field.Name = "IDAzienda" Then
                    Field.Value = m_App.IDFirm
                End If
                
                If Field.Name = "IDFiliale" Then
                    Field.Value = m_App.Branch
                End If
                If Field.Name = "Anno" Then
                    Field.Value = Year(Date)
                End If
    
                If Field.Name = "IDUtenteModifica" Then
                    Field.Value = TheApp.IDUser
                End If
                
                If Field.Name = "DataModifica" Then
                    Field.Value = Date
                End If
                
                If Field.Name = "PCModifica" Then
                    Field.Value = GET_NOMECOMPUTER
                End If
                
                If Field.Name = "NomeUtentePCModifica" Then
                    Field.Value = GET_NOMEUTENTE
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

    m_Semaphore.ClearObjectAction m_DocType.ID, SemAllObjects, SemAllActions
    m_Semaphore.SetObjectAction m_DocType.ID, m_Document.Fields("ID" & m_App.TableName).Value, SemAllActions
    
    'Refresh delle variabili di stato
    m_Changed = False
    m_Search = False
    m_Saved = True
    
    'Refresh dello stato della ToolBar standard in modalità variazione
    SetStatus4Modality Modify
       
    
  
    
    
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
    Dim sSQL As String
    
    
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
    
    If Me.chkPassaggioInFatt.Value = vbChecked Then
        MsgBox "Impossibile eliminare la liquidazione poichè risulta collegato ad un documento per conto dei soci", vbCritical, TheApp.FunctionName
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
        
''''''''''''''''''''''''''''''''''''''''''ELIMINAZIONE DATI'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        sSQL = "DELETE FROM RV_POLiquidazioneRigheEla "
        sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        Cn.Execute sSQL
        
        sSQL = "DELETE FROM RV_POLiquidazioneRighe "
        sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        Cn.Execute sSQL
        
        sSQL = "DELETE FROM RV_POLiquidazioneRigheTratt "
        sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        Cn.Execute sSQL

        sSQL = "DELETE FROM RV_POLiquidazioneRigheFatt "
        sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
        Cn.Execute sSQL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Cancellazione
        m_DocumentsLink.Refresh
        m_DocumentsLink1.Refresh
        m_DocumentsLink2.Refresh
        
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
    
    GET_OGGETTI_PER_STAMPE
    
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
        'For Each Field In m_DocType.Fields
        '    Field.Value = Empty
        'Next
        
        'Viene inserita la condizione di ricerca basata sull'ID del record corrente.
        m_DocType.Fields("IDUtente").Value = TheApp.IDUser
        m_DocType.Fields("ID" & m_App.TableName).Value = m_Document.Fields("ID" & m_App.TableName).Value
        
        'Viene creato un filtro temporaneo per il Crystals Reports.
        m_DocType.RemoveFilter "Form"
        Set m_Report.Filter = m_DocType.AddFilterWithConditions("Form")
        
    Else
        'Modalità vista tabellare
        
        'Viene passato il filtro corrente al Crystals Reports.
        m_DocType.Fields("IDUtente").Value = TheApp.IDUser
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

If Name = "Codice" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Codice"
    
    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Codice", "Codice", 1   'STRINGTYPE
    oSearch.AddDisplayField "Anagrafica", "Anagrafica", 1 'STRINGTYPE
    oSearch.AddDisplayField "Nome", "Nome", 1   'STRINGTYPE

 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT Anagrafica.IDAnagrafica, Fornitore.Codice, Anagrafica.Anagrafica, Anagrafica.Nome "
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

If Name = "Numero liquidazione" Then
    'La Caption della finestra di ricerca
    oSearch.Caption = "Periodo di liquidazione"

    'Vengono assegnati i campi su cui effettuare la ricerca.
    'Questi campi verranno visualizzati nella tabella della finestra di ricerca
    'ed in quella per l'impostazione del filtro.
    'NOTA:
    'Il primo campo inserito (ovvero "Anagrafica" in questo caso) verrà associato alla combo
    'presente nella finestra di ricerca.
    oSearch.AddDisplayField "Numero Liquidazione", "NumeroLiquidazione", 0 'NUMBERTYPE
    oSearch.AddDisplayField "Data inizio periodo", "DataInizio", 2   'STRINGTYPE
    oSearch.AddDisplayField "Data fine periodo", "DataFine", 2   'STRINGTYPE
 
    'Assegna la condizione iniziale
    oSearch.Start = Value

    'Query SQL con cui effettuare le ricerche in base dati.
    sSQL = "SELECT IDRV_POLiquidazionePeriodo, NumeroLiquidazione, DataInizio, DataFine "
    sSQL = sSQL & "FROM RV_POLiquidazionePeriodo "
    sSQL = sSQL & "ORDER BY NumeroLiquidazione DESC"

    'Assegnazione della query di ricerca
    oSearch.SQL = sSQL
    
    'Esegue la ricerca e restituisce l'eventuale riga selezionata dall'utente
    Set oRes = oSearch.Exec
    
    If Not oRes.EOF Then
        'Riporta il valore della riga selezionata nella maschera del filtro
             
        Value = oRes("NumeroLiquidazione")
                
    End If
End If
End Sub




Private Sub cboIva_Click()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT AliquotaIva FROM Iva WHERE IDIva=" & Me.cboIva.CurrentID

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtAliquotaIva.Value = 0
Else
    Me.txtAliquotaIva.Value = fnNotNullN(rs!AliquotaIva)
End If

rs.CloseResultset
Set rs = Nothing

Me.txtImportoTotaleFatt.Value = (Me.txtImportoUnitarioFatt.Value + ((Me.txtImportoUnitarioFatt.Value / 100) * Me.txtAliquotaIva.Value)) * Me.txtQtaRigaFatt.Value

End Sub

Private Sub cboTipoRicalcoloComm_Click()
    If Me.cboTipoRicalcoloComm.CurrentID = 1 Then
        Me.txtPercentualeTrattenuta.Enabled = True
        Me.txtImportoNuovaTrattenuta.Enabled = False
    Else
        Me.txtPercentualeTrattenuta.Enabled = False
        Me.txtImportoNuovaTrattenuta.Enabled = True
    
    End If
End Sub

Private Sub CDSocioFatt_ChangeElement()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkPassaggioInFatt_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub chkUfficiale_Click()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdElimina_Click()
Dim Testo As String

If MsgBox("Eliminare la riga selezionata?", vbQuestion + vbYesNo, "Elimazione riga") = vbYes Then
If Me.txtTrattenuteConferimento.Value > 0 Then
    Testo = "Questa riga contiene un valore sulla trattenuta di conferimento" & vbCrLf
    Testo = Testo & "Vuoi continuare ad eliminare la riga?"
    If MsgBox(Testo, vbQuestion + vbYesNo, "Eliminazione riga") = vbNo Then Exit Sub
End If

    m_DocumentsLink.Delete
    
    If Not (BrwMain.Visible) Then Change
    
    RicalcolaTotaleLiquidazione
End If
End Sub

Private Sub cmdElimina_RigheFatt_Click()
    m_DocumentsLink2.DeleteRowFromBuffer
    
    If Not (BrwMain.Visible) Then Change

End Sub

Private Sub cmdEliminaTrattenuta_Click()
Dim Testo As String

    If m_DocumentsLink1("IDRV_POAnticipazioniSocioRighe").Value > 0 Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "Eliminando questa riga si eliminano anche i collegamenti nella riga di anticipazione." & vbCrLf
        Testo = Testo & "Vuoi continuare?"
        
        If MsgBox(Testo, vbQuestion + vbYesNo, "Controllo eliminazione dati") = vbNo Then Exit Sub
        
        AGGIORNA_ANTICIPAZIONI 0, fnNotNullN(m_DocumentsLink1("IDRV_POAnticipazioniSocioRighe").Value), fnNotNullN(m_DocumentsLink1("IDRV_POAnticipazioniSocio").Value)
        
    End If
    
    m_DocumentsLink1.Delete
    
    Me.txtTotaleTrattenuteAgg.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 1)
    Me.txtTotaleTrattAggRiep.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 2)

    If TIPO_IMPORTO_DOCUMENTO <= 1 Then
        Me.txtNettoLiquidazione.Value = Me.txtTotaleDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    Else
        Me.txtNettoLiquidazione.Value = Me.txtTotaleLordoDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    End If
    
    OnSave

End Sub
Private Function AGGIORNA_ANTICIPAZIONI(IDLiquidazione As Long, IDAnticipazioneRighe As Long, IDAnticipazioneTesta As Long)
Dim sSQL As String
Dim rsAgg As ADODB.Recordset

sSQL = "SELECT * FROM RV_POAnticipazioniSocioRighe "
sSQL = sSQL & "WHERE IDRV_POAnticipazioniSocioRighe=" & IDAnticipazioneRighe

Set rsAgg = New ADODB.Recordset

rsAgg.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

If Not rsAgg.EOF Then
    rsAgg!IDRV_POLiquidazione = IDLiquidazione
    rsAgg!IDRV_POTipoStatoAnticipazioneInteresse = 1
    rsAgg!TipoStatoAnticipazioneInteresse = "Da Elaborare"
    rsAgg.Update
End If


rsAgg.Close
Set rsAgg = Nothing
End Function
Private Sub cmdNuovaTrattenuta_Click()
    If m_DocumentsLink1.TableNew Then
        m_DocumentsLink1.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink1.NewRow
    
    Me.txtNuovaTrattenuta.SetFocus
    
End Sub

Private Sub cmdNuovo_RigheFatt_Click()
    If m_DocumentsLink2.TableNew Then
        m_DocumentsLink2.AbortNewRow
    End If
    'Crea una nuova riga vuota nel buffer
    m_DocumentsLink2.NewRow
    
End Sub

Private Sub cmdRicalcolaTotali_Click()
    RicalcolaTotaleLiquidazione
End Sub

Private Sub cmdSalva_Click()
Dim Testo As String

    

    If fnNotNullN(m_DocumentsLink("IDArticolo").Value) > 0 Then
        If TIPO_IMPORTO_ARTICOLO <= 1 Then
            Valore_TotaleIvaDocumento = ((Me.txtImportoTotaleVenduto.Value / 100) * fnNotNullN(m_DocumentsLink("AliquotaIva_per_Imp_Vend").Value))
            Valore_TotaleLordoDocumento = Me.txtImportoTotaleVenduto.Value + Valore_TotaleIvaDocumento
        Else
            Valore_TotaleIvaDocumento = ((Me.txtImportoTotaleVenduto.Value / 100) * fnNotNullN(m_DocumentsLink("AliquotaIva_per_Imp_Medio").Value))
            Valore_TotaleLordoDocumento = Me.txtImportoTotaleVenduto.Value + Valore_TotaleIvaDocumento
        End If
    End If
    
    If GET_CONTROLLO_RIGHE_TRATTENUTE(fnNotNullN(m_Document(m_Document.PrimaryKey).Value)) = False Then
        Testo = "ATTENZIONE!!!!" & vbCrLf
        Testo = Testo & "Una o più righe di Nuove Trattenute non ha impostato il tipo di calcolo per la trattenuta "
        Testo = Testo & "e pertanto non verranno ricalcolate in base alle modifiche effettuate." & vbCrLf
        Testo = Testo & "Vuoi continuare?"
        If MsgBox(Testo, vbQuestion + vbYesNo, "Salvataggio riga di liquidazione") = vbNo Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    'Me.txtTotaleDocumento.Value = Me.txtTotaleDocumento.Value - fnNotNullN(m_DocumentsLink("ImponibileDaReg").Value) + Me.txtImportoTotaleVenduto.Value
    'Me.txtTotaleIvaDocumento.Value = Me.txtTotaleIvaDocumento.Value - fnNotNullN(m_DocumentsLink("ImpostaDaReg").Value) + Valore_TotaleIvaDocumento
    'Me.txtTotaleLordoDocumento.Value = Me.txtTotaleLordoDocumento.Value - fnNotNullN(m_DocumentsLink("ImportoLordoDaReg").Value) + Valore_TotaleLordoDocumento
    'Me.txtTotaleTrattenuteGenerali.Value = Me.txtTotaleTrattenuteGenerali - fnNotNullN(m_DocumentsLink("TrattenuteGenerali").Value) + Me.txtTrattenuteGenerali.Value
    'Me.txtTotaleTrattenutePerLavorazioni.Value = Me.txtTotaleTrattenutePerLavorazioni - fnNotNullN(m_DocumentsLink("TrattenutePerLavorazione").Value) + Me.txtTrattenutaPerLavorazione.Value
    
    'Me.txtTotaleTrattenute.Value = Me.txtTotaleTrattenute.Value - m_DocumentsLink("TrattenuteTotali").Value + Me.txtTotaleTrattenuta.Value
    'Me.txtTotaleTrattenute.Value = Me.txtTotaleTrattenute.Value - (Me.txtTotaleTrattenuteAgg.Value + Me.txtTotaleTrattAggRiep.Value)
    
    'If TIPO_IMPORTO_DOCUMENTO <= 1 Then
    '    Me.txtNettoLiquidazione.Value = Me.txtTotaleDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    'Else
    '    Me.txtNettoLiquidazione.Value = Me.txtTotaleLordoDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    'End If
    

    
    m_DocumentsLink("QuantitaLavorata").Value = Me.txtQtaLavorata.Value
    m_DocumentsLink("QuantitaQuadrata").Value = Me.txtQtaQuadLav.Value
    m_DocumentsLink("QuantitaTotaleLavorata").Value = Me.txtTotaleLavorata.Value
    m_DocumentsLink("QuantitaVenduta").Value = Me.txtQtaVenduta.Value
    If Me.txtImportoUnitarioVenduto.Value > 0 Then
        m_DocumentsLink("ImportoUnitarioDaReg").Value = Me.txtImportoUnitarioVenduto.Value
        m_DocumentsLink("ImponibileDaReg").Value = Me.txtImportoTotaleVenduto.Value
        m_DocumentsLink("ImpostaDaReg").Value = Valore_TotaleIvaDocumento
        m_DocumentsLink("ImportolordoDaReg").Value = Valore_TotaleLordoDocumento
        m_DocumentsLink("Quantita_Per_Totali").Value = m_DocumentsLink("QuantitaVenduta").Value
    Else
        m_DocumentsLink("ImportoUnitarioDaReg").Value = Null
        m_DocumentsLink("ImponibileDaReg").Value = Null
        m_DocumentsLink("ImpostaDaReg").Value = Null
        m_DocumentsLink("ImportolordoDaReg").Value = Null
        m_DocumentsLink("Quantita_Per_Totali").Value = Null
    End If
    m_DocumentsLink("TotaleTrattenutaConferimento").Value = Me.txtTrattenuteConferimento.Value
    m_DocumentsLink("TrattenuteperLavorazione").Value = Me.txtTrattenutaPerLavorazione.Value
    m_DocumentsLink("TrattenuteGenerali").Value = Me.txtTrattenuteGenerali.Value
    m_DocumentsLink("TrattenuteTotali").Value = Me.txtTotaleTrattenuta.Value
    
    'UTENTE DI ULTIMA MODIFICA
    m_DocumentsLink("IDUtenteModifica").Value = TheApp.IDUser
    m_DocumentsLink("DataModifica").Value = Date
    m_DocumentsLink("PCModifica").Value = GET_NOMECOMPUTER
    m_DocumentsLink("NomeUtentePCModifica").Value = GET_NOMEUTENTE

    m_DocumentsLink("TrattenutaValoreGen1").Value = Me.txtTrattValGen1.Value
    m_DocumentsLink("TrattenutaValoreGen2").Value = Me.txtTrattValGen2.Value
    m_DocumentsLink("TrattenutaPercGen1").Value = Me.txtTrattPercGen1.Value
    m_DocumentsLink("TrattenutaPercGen2").Value = Me.txtTrattPercGen2.Value
    
    m_DocumentsLink("TrattenutaValoreLav1").Value = Me.txtTrattValLav1.Value
    m_DocumentsLink("TrattenutaValoreLav2").Value = Me.txtTrattValLav2.Value
    m_DocumentsLink("TrattenutaPercLav1").Value = Me.txtTrattPercLav1.Value
    m_DocumentsLink("TrattenutaPercLav2").Value = Me.txtTrattPercLav2.Value

    
    m_DocumentsLink.Save
    
    Screen.MousePointer = 0
    
    m_DocumentsLink.Move Me.GrigliaTrattenuteEla.ListIndex - 1

    'OnSave
    
    RicalcolaTotaleLiquidazione
    OnSave
    'If Not (BrwMain.Visible) Then Change
    
    
End Sub

Private Sub cmdSalva_RigheFatt_Click()
    
    m_DocumentsLink2("DescrizioneRiga").Value = Me.txtDescrizioneRigaFatt.Text
    m_DocumentsLink2("IDIva").Value = Me.cboIva.CurrentID
    m_DocumentsLink2("AliquotaIva").Value = Me.txtAliquotaIva.Value
    m_DocumentsLink2("ImportoUnitario").Value = Me.txtImportoUnitarioFatt.Value
    m_DocumentsLink2("ImportoTotale").Value = Me.txtImportoTotaleFatt.Value
    m_DocumentsLink2("IDContoPDC").Value = Link_ContoPDC
    m_DocumentsLink2("CodicePDC").Value = Me.txtCodiceConto.Text
    m_DocumentsLink2("DescrizionePDC").Value = Me.txtDescrizioneConto.Text
    m_DocumentsLink2("Quantita").Value = Me.txtQtaRigaFatt.Value
    
    m_DocumentsLink2.SaveRowToBuffer
    
    m_DocumentsLink2.Move Me.GridRigheFatt.ListIndex - 1
    
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub cmdSalvaTrattenuta_Click()

  
    If Me.cboTipoTrattenutaAggiuntiva.CurrentID = 0 Then
        MsgBox "Inserire il tipo di trattenuta aggiuntiva", vbInformation, "Controllo inserimento dati"
        Me.cboTipoTrattenutaAggiuntiva.SetFocus
        Exit Sub
    End If
    

    
    If Me.cboTipoTrattenutaAggiuntiva.CurrentID = 3 Then
        MsgBox "Tipo di trattenuta incompatibile", vbInformation, "Controllo inserimento dati"
        Me.cboTipoTrattenutaAggiuntiva.SetFocus
        Exit Sub
    End If
    
    If Me.cboTipoRicalcoloComm.CurrentID = 0 Then
        MsgBox "Inserire il tipo di calcolo della trattenuta", vbInformation, "Controllo inserimento dati"
        Me.cboTipoRicalcoloComm.SetFocus
        Exit Sub
    End If
    
    If Len(Trim(Me.txtNuovaTrattenuta.Text)) = 0 Then
        MsgBox "Inserire una descrizione di trattenuta aggiuntiva", vbInformation, "Controllo inserimento dati"
        Me.txtNuovaTrattenuta.SetFocus
        Exit Sub
    End If
    If Me.txtPercentualeTrattenuta.Value = 0 Then
        MsgBox "Inserire una percentuale di trattenuta aggiuntiva", vbInformation, "Controllo inserimento dati"
        Me.txtPercentualeTrattenuta.SetFocus
        Exit Sub
    End If
    
    m_DocumentsLink1("DescrizioneAggiuntiva").Value = Me.txtNuovaTrattenuta.Text
    m_DocumentsLink1("IDRV_POTipoTrattenutaAggiuntiva").Value = Me.cboTipoTrattenutaAggiuntiva.CurrentID
    m_DocumentsLink1("IDRV_POSegnoTrattenuta") = Me.cboSegnoTrattenuta.CurrentID
    m_DocumentsLink1("Percentuale") = Me.txtPercentualeTrattenuta.Value
    m_DocumentsLink1("ImportoTrattenuta").Value = Me.txtImportoNuovaTrattenuta.Value
    m_DocumentsLink1("IDRV_POTipoRicalcoloComm").Value = Me.cboTipoRicalcoloComm.CurrentID
    
    m_DocumentsLink1.Save
    
    m_DocumentsLink1.Move Me.GrigliaNuoveTrattenute.ListIndex - 1
    
    Me.txtTotaleTrattenuteAgg.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 1)
    Me.txtTotaleTrattAggRiep.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 2)
    

    If TIPO_IMPORTO_DOCUMENTO <= 1 Then
        Me.txtNettoLiquidazione.Value = Me.txtTotaleDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    Else
        Me.txtNettoLiquidazione.Value = Me.txtTotaleLordoDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    End If

    OnSave
    
End Sub


Private Sub cmdTipoPrezzoMedio_Click()
If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub

LINK_RIGA_PREZZO_MEDIO = m_DocumentsLink("IDRV_POLiquidazioneRigheElaPM").Value
LINK_PERIODO_LIQUIDAZIONE = Me.cboPeriodoLiquidazione.CurrentID

If LINK_RIGA_PREZZO_MEDIO = 0 Then Exit Sub

frmPrezzoMedio.Show vbModal

End Sub

Private Sub cmdTrattenuteUtilizzate_Click()


LINK_LIQUIDAZIONE = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

frmTipoTrattenuteRaggr.Show vbModal
End Sub

Private Sub Command1_Click()
    If fnNotNullN(m_Document(m_Document.PrimaryKey).Value) <= 0 Then Exit Sub
    
    LINK_LIQUIDAZIONE = fnNotNullN(m_Document(m_Document.PrimaryKey).Value)
    
    frmTipoTrattenute.Show vbModal
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
        Set Me.GrigliaTrattenuteEla.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink.TableName).Data
        Set Me.GrigliaNuoveTrattenute.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink1.TableName).Data
        Set Me.GridRigheFatt.Recordset = m_Document.DocumentsLinks.Item(m_DocumentsLink2.TableName).Data

    'End If
    
    Me.txtSocio.Text = fnNotNull(m_Document("Anagrafica").Value) & " " & fnNotNull(m_Document("Nome").Value)
    Me.CDSocioFatt.Load fnNotNullN(m_Document("IDAnagraficaFatturazione").Value)
    Me.txtNumeroLiquidazione.Value = fnNotNullN(m_Document("NumeroLiquidazione").Value)
    Me.txtDataLiquidazione.Text = fnNotNull(m_Document("DataLiquidazione").Value)
    Me.cboPeriodoLiquidazione.WriteOn fnNotNullN(m_Document("IDRV_POLiquidazionePeriodo").Value)
    Me.txtOggetto.Text = fnNotNull(m_Document("Oggetto").Value)
    Me.chkPassaggioInFatt.Value = fnNormBoolean(m_Document("PassaggioInFatturazione").Value)
    Me.txtNumeroInterno.Value = fnNotNullN(m_Document("NumeroProtInt").Value)
    Link_Oggetto = fnNotNullN(m_Document("IDOggetto").Value)
    Me.cboCatMerc.WriteOn fnNotNullN(m_Document("IDCategoriaMerceologica").Value)
    'rif11 end
    
    GET_PARAMETRI_LIQUIDAZIONE fnNotNullN(m_Document("IDRV_POLiquidazionePeriodo").Value)
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
    Dim QtaVenduta As Double
    Dim QtaScarto As Double
    
    On Error Resume Next
    

    If Not (m_DocumentsLink.BOF And m_DocumentsLink.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.txtArticoloConferito.Text = fnNotNull(m_DocumentsLink("Articolo_Conf").Value)
        Me.txtLottoConferito.Text = fnNotNull(m_DocumentsLink("CodiceLottoArticolo_Conf"))
        Me.txtQtaConferita.Text = fnNotNullN(m_DocumentsLink("QuantitaConferita").Value)
        Me.txtArticoloLavorato.Text = fnNotNull(m_DocumentsLink("Articolo").Value)
        Me.txtLottoLavorato.Text = fnNotNull(m_DocumentsLink("CodiceLottoArticolo").Value)
        Me.txtQtaLavorata.Value = m_DocumentsLink("QuantitaLavorata").Value
        Me.txtQtaQuadLav.Value = m_DocumentsLink("QuantitaQuadrata").Value
        Me.txtTotaleLavorata.Value = fnNotNullN(m_DocumentsLink("QuantitaTotaleLavorata").Value)
        Me.txtQtaVenduta.Value = fnNotNullN(m_DocumentsLink("QuantitaVenduta").Value)
        Me.txtTrattenuteConferimento.Value = fnNotNullN(m_DocumentsLink("TotaleTrattenutaConferimento").Value)
        Me.txtTrattenutaPerLavorazione.Value = fnNotNullN(m_DocumentsLink("TrattenutePerLavorazione").Value)
        Me.txtTrattenuteGenerali.Value = fnNotNullN(m_DocumentsLink("TrattenuteGenerali").Value)
        Me.txtTotaleTrattenuta.Value = m_DocumentsLink("TrattenuteTotali").Value
        Me.txtImportoTotaleVenduto.Value = m_DocumentsLink("ImponibileDaReg").Value
        Me.txtImportoUnitarioVenduto.Value = m_DocumentsLink("ImportoUnitarioDaReg").Value
        Me.txtOrigineDocumento.Text = fnNotNull(m_DocumentsLink("Oggetto").Value)
        Me.txtLottoConferito.Text = "Conferimento n° " & fnNotNullN(m_DocumentsLink("NumeroDocumento").Value) & " del " & fnNotNull(m_DocumentsLink("DataConferimento").Value)
        
        Valore_TotaleLordoDocumento = fnNotNullN(m_DocumentsLink("ImportoLordoDaReg").Value)
        Valore_TotaleIvaDocumento = fnNotNullN(m_DocumentsLink("ImpostaDaReg").Value)
        
        Me.cboTipoLavorazione.WriteOn fnNotNullN(m_DocumentsLink("IDRV_POLavorazione").Value)
        Link_CategoriaMerceologica = GET_CATEGORIA_MERCEOLOGICA
        Me.cboTipoPrezzoMedio.WriteOn fnNotNullN(m_DocumentsLink("IDRV_POTipoPrezzoMedio").Value)
        
        Me.txtTrattValGen1.Value = fnNotNullN(m_DocumentsLink("TrattenutaValoreGen1").Value)
        Me.txtTrattValGen2.Value = fnNotNullN(m_DocumentsLink("TrattenutaValoreGen2").Value)
        Me.txtTrattPercGen1.Value = fnNotNullN(m_DocumentsLink("TrattenutaPercGen1").Value)
        Me.txtTrattPercGen2.Value = fnNotNullN(m_DocumentsLink("TrattenutaPercGen2").Value)
        
        Me.txtTrattValLav1.Value = fnNotNullN(m_DocumentsLink("TrattenutaValoreLav1").Value)
        Me.txtTrattValLav2.Value = fnNotNullN(m_DocumentsLink("TrattenutaValoreLav2").Value)
        Me.txtTrattPercLav1.Value = fnNotNullN(m_DocumentsLink("TrattenutaPercLav1").Value)
        Me.txtTrattPercLav2.Value = fnNotNullN(m_DocumentsLink("TrattenutaPercLav2").Value)
        
        bValue = True
        
        If fnNotNullN(m_DocumentsLink("IDTipoOggetto")) = 11 Then
            Me.txtTrattenutaPerLavorazione.Enabled = False
            Me.txtTrattenuteGenerali.Enabled = False
            Me.txtTotaleTrattenuta.Enabled = False
            Me.txtTrattValGen1.Enabled = False
            Me.txtTrattValGen2.Enabled = False
            Me.txtTrattPercGen1.Enabled = False
            Me.txtTrattPercGen2.Enabled = False
            
            Me.txtTrattValLav1.Enabled = False
            Me.txtTrattValLav2.Enabled = False
            Me.txtTrattPercLav1.Enabled = False
            Me.txtTrattPercLav2.Enabled = False
        Else
            Me.txtTrattenutaPerLavorazione.Enabled = True
            Me.txtTrattenuteGenerali.Enabled = True
            Me.txtTotaleTrattenuta.Enabled = True
            Me.txtTrattValGen1.Enabled = True
            Me.txtTrattValGen2.Enabled = True
            Me.txtTrattPercGen1.Enabled = True
            Me.txtTrattPercGen2.Enabled = True
            
            Me.txtTrattValLav1.Enabled = True
            Me.txtTrattValLav2.Enabled = True
            Me.txtTrattPercLav1.Enabled = True
            Me.txtTrattPercLav2.Enabled = True
        End If
        
        '----------------------------------------------------------------------------
        'Popola i controlli associati al sottodocumento con i valori presenti
        'nell'oggetto DocumentsLink
        '----------------------------------------------------------------------------
       
    Else
        'Il DocumentsLink è vuoto - non contiene righe.
        '---------------------------------------------
        'Ripulisce i controlli associati al sottodocumento
        '---------------------------------------------
        
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    

        
 
  
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
QtaScarto = GET_QTA_SCARTO(fnNotNullN(m_DocumentsLink("IDRV_POCaricoMerceRighe").Value), fnNotNullN(m_DocumentsLink("IDRV_POLiquidazione").Value))
QtaVenduta = GET_QTA_VENDUTA(fnNotNullN(m_DocumentsLink("IDRV_POCaricoMerceRighe").Value), fnNotNullN(m_DocumentsLink("IDRV_POLiquidazione").Value))
Me.txtDifferenza.Value = fnNotNullN(Me.txtQtaConferita.Text) - QtaVenduta - QtaScarto

End Sub
Private Function GET_QTA_VENDUTA(IDConferimentoRiga As Long, IDLiquidazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT QuantitaVenduta FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND TipoRiga=1"

GET_QTA_VENDUTA = 0


Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
GET_QTA_VENDUTA = GET_QTA_VENDUTA + fnNotNullN(rs!QuantitaVenduta)

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GET_QTA_SCARTO(IDConferimentoRiga As Long, IDLiquidazione As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset


sSQL = "SELECT QuantitaVenduta FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
sSQL = sSQL & " AND IDRV_POCaricoMerceRighe=" & IDConferimentoRiga
sSQL = sSQL & " AND TipoRiga=2"

GET_QTA_SCARTO = 0


Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
GET_QTA_SCARTO = GET_QTA_SCARTO + fnNotNullN(rs!QuantitaVenduta)

rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
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
        
        Me.txtNuovaTrattenuta.Text = fnNotNull(m_DocumentsLink1("DescrizioneAggiuntiva").Value)
        Me.txtImportoNuovaTrattenuta.Value = fnNotNullN(m_DocumentsLink1("ImportoTrattenuta").Value)
        Me.cboTipoTrattenutaAggiuntiva.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POTipoTrattenutaAggiuntiva").Value)
        Me.txtPercentualeTrattenuta.Value = fnNotNullN(m_DocumentsLink1("Percentuale").Value)
        Me.cboTipoRicalcoloComm.WriteOn fnNotNullN(m_DocumentsLink1("IDRV_POTipoRicalcoloComm").Value)

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
        Me.txtNuovaTrattenuta.Text = ""
        Me.txtImportoNuovaTrattenuta.Value = 0
        Me.cboTipoTrattenutaAggiuntiva.WriteOn 0
        Me.txtPercentualeTrattenuta.Value = 0
        Me.cboTipoRicalcoloComm.WriteOn 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
        'Me.cboCausaleQuadratura.Enabled = bValue
        
    Me.txtNuovaTrattenuta.Enabled = bValue
    'Me.txtImportoNuovaTrattenuta.Enabled = bValue
    Me.cboTipoTrattenutaAggiuntiva.Enabled = bValue
    'Me.txtPercentualeTrattenuta.Enabled = bValue
    Me.cboTipoRicalcoloComm.Enabled = bValue
    
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
    Me.cmdNuovaTrattenuta.Enabled = True
    Me.cmdSalvaTrattenuta.Enabled = bValue
    Me.cmdEliminaTrattenuta.Enabled = bValue
    
End Sub


Private Sub m_DocumentsLink2_OnReposition()
    Dim bValue As Boolean
    Dim iIndex As Integer
    
    On Error Resume Next
    

    If Not (m_DocumentsLink2.BOF And m_DocumentsLink2.EOF) Then
        'Il DocumentsLink non è vuoto - contiene dei dati.
        
        Me.txtDescrizioneRigaFatt.Text = fnNotNull(m_DocumentsLink2("DescrizioneRiga").Value)
        Me.cboIva.WriteOn fnNotNullN(m_DocumentsLink2("IDIva").Value)
        Me.txtAliquotaIva.Value = fnNotNullN(m_DocumentsLink2("AliquotaIva").Value)
        Me.txtImportoUnitarioFatt.Value = fnNotNullN(m_DocumentsLink2("ImportoUnitario").Value)
        Me.txtImportoTotaleFatt.Value = fnNotNullN(m_DocumentsLink2("ImportoTotale").Value)
        Link_ContoPDC = fnNotNullN(m_DocumentsLink2("IDContoPDC").Value)
        Me.txtCodiceConto.Text = fnNotNull(m_DocumentsLink2("CodicePDC").Value)
        Me.txtDescrizioneConto.Text = fnNotNull(m_DocumentsLink2("DescrizionePDC").Value)
        Me.txtQtaRigaFatt.Value = fnNotNullN(m_DocumentsLink2("Quantita").Value)
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
        Me.txtDescrizioneRigaFatt.Text = ""
        Me.cboIva.WriteOn 0
        Me.txtAliquotaIva.Value = 0
        Me.txtImportoUnitarioFatt.Value = 0
        Me.txtImportoTotaleFatt.Value = 0
        Link_ContoPDC = 0
        Me.txtCodiceConto.Text = ""
        Me.txtDescrizioneConto.Text = ""
        Me.txtQtaRigaFatt.Value = 0
        bValue = False
    End If
    
    
    
    
    'Abilita/disabilita i controlli a seconda che ci sia o meno almeno un sottodocumento
    
        'Me.cboCausaleQuadratura.Enabled = bValue
        
     
        Me.txtDescrizioneRigaFatt.Enabled = bValue
        Me.cboIva.Enabled = bValue
        Me.txtAliquotaIva.Enabled = bValue
        Me.txtImportoUnitarioFatt.Enabled = bValue
        Me.txtImportoTotaleFatt.Enabled = bValue
        Me.lblPianodeiDeiConti(4).Enabled = bValue
        Me.txtCodiceConto.Enabled = bValue
        Me.txtDescrizioneConto.Enabled = bValue
        Me.txtQtaRigaFatt.Enabled = bValue
    'Pulsanti Nuovo, Salva, Elimina del sottodocumento.
    
    Me.cmdNuovo_RigheFatt.Enabled = True
    Me.cmdSalva_RigheFatt.Enabled = bValue
    Me.cmdElimina_RigheFatt.Enabled = bValue
End Sub

Private Sub txtImportoNuovaTrattenuta_Change()
'    If Me.txtTotaleDocumento.Value > 0 Then
'        Me.txtPercentualeTrattenuta.Value = (Me.txtImportoNuovaTrattenuta.Value / Me.txtTotaleDocumento.Value) * 100
'    End If
End Sub

Private Sub txtImportoNuovaTrattenuta_LostFocus()
    If Me.txtTotaleDocumento.Value > 0 Then
        Me.txtPercentualeTrattenuta.Value = (Me.txtImportoNuovaTrattenuta.Value / Me.txtTotaleDocumento.Value) * 100
    End If

End Sub

Private Sub txtImportoUnitarioFatt_Change()
    Me.txtImportoTotaleFatt.Value = (Me.txtImportoUnitarioFatt.Value + ((Me.txtImportoUnitarioFatt.Value / 100) * Me.txtAliquotaIva.Value)) * Me.txtQtaRigaFatt.Value
End Sub
Private Sub txtImportoUnitarioVenduto_LostFocus()
    If fnNotNullN(m_DocumentsLink("ImportoUnitarioDaReg").Value) <> Me.txtImportoUnitarioVenduto.Value Then
        Me.txtImportoTotaleVenduto.Value = Me.txtImportoUnitarioVenduto.Value * Me.txtQtaVenduta.Value
        
        If fnNotNullN(m_DocumentsLink("IDTipoOggetto").Value) <> 11 Then
            fncCalcoloTrattenute
        End If
    End If
End Sub

Private Sub txtNettoLiquidazione_Change()
    If Not (BrwMain.Visible) Then Change
End Sub



Private Sub txtOggetto_Change()
    If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtPercentualeTrattenuta_Change()
    If Me.txtTotaleDocumento.Value > 0 Then
        Me.txtImportoNuovaTrattenuta.Value = (Me.txtTotaleDocumento.Value / 100) * Me.txtPercentualeTrattenuta.Value
    End If
End Sub

Private Sub txtQtaLavorata_LostFocus()
    If fnNotNullN(m_DocumentsLink("QuantitaLavorata").Value) <> Me.txtQtaLavorata.Value Then
        Me.txtTotaleLavorata.Value = Me.txtQtaLavorata.Value + Me.txtQtaQuadLav.Value
        If fnNotNullN(m_DocumentsLink("IDTipoOggetto")) <> 11 Then
            fncCalcoloTrattenute
        End If
    End If
End Sub
Private Sub txtQtaQuadLav_LostFocus()
    If fnNotNullN(m_DocumentsLink("QuantitaQuadrata").Value) <> Me.txtQtaQuadLav.Value Then
        Me.txtTotaleLavorata.Value = Me.txtQtaLavorata.Value + Me.txtQtaQuadLav.Value
        
        If fnNotNullN(m_DocumentsLink("IDTipoOggetto")) <> 11 Then
            fncCalcoloTrattenute
        End If

    End If
End Sub

Private Sub txtQtaVenduta_LostFocus()
    If fnNotNullN(m_DocumentsLink("QuantitaVenduta").Value) <> Me.txtQtaVenduta.Value Then
        If fnNotNullN(m_DocumentsLink("IDTipoOggetto")) = 11 Then
            If Me.txtQtaVenduta.Value > 0 Then
                Me.txtQtaVenduta.Value = fnNotNullN(m_DocumentsLink("QuantitaVenduta").Value)
                Exit Sub
            Else
                Me.txtImportoTotaleVenduto.Value = Me.txtImportoUnitarioVenduto.Value * Me.txtQtaVenduta.Value
            End If
        Else
            Me.txtImportoTotaleVenduto.Value = Me.txtImportoUnitarioVenduto.Value * Me.txtQtaVenduta.Value
        End If
        If fnNotNullN(m_DocumentsLink("IDTipoOggetto")) <> 11 Then
            fncCalcoloTrattenute
        End If
    End If
End Sub
Private Sub txtTotaleDocumento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtTotaleIvaDocumento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtTotaleLordoDocumento_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Sub txtTotaleTrattenute_Change()
    If Not (BrwMain.Visible) Then Change
End Sub
Private Function Get_ArticoloConferito(IDArticolo As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Articolo FROM Articolo WHERE IDArticolo=" & IDArticolo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Get_ArticoloConferito = fnNotNull(rs!Articolo)
Else
    Get_ArticoloConferito = ""
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function Get_LottoArticoloConferito(IDLotto As Long) As String
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM LottoArticolo WHERE IDLottoArticolo=" & IDLotto

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF = False Then
    Get_LottoArticoloConferito = fnNotNull(rs!Codice)
Else
    Get_LottoArticoloConferito = ""
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Function GetPianoDeiConti() As Long
On Error GoTo ERR_GetPianoDeiConti
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    sSQL = "SELECT IDPianoDeiConti FROM PianoDeiConti WHERE ("
    sSQL = sSQL & "(IDAzienda = " & m_App.Branch & ") AND "
    sSQL = sSQL & "(IDEsercizio= " & VarIDEsercizio & "))"
    
    Set rs = Cn.OpenResultset(sSQL)
    
    If rs.EOF = False Then
        GetPianoDeiConti = rs!IDPianoDeiConti
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
        .IDPDC = Link_PianoDeiConti
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
            'Identificativo unico del Conto o del Ramo
        Link_ContoPDC = oNode.ID
        'Codifica completa del Conto o del Ramo
        Me.txtCodiceConto.Text = oNode.CompletedCode
        Me.txtDescrizioneConto.Text = oNode.Description


    End If
End Sub
Private Sub GET_PARAMETRI_LIQUIDAZIONE(Link_Periodo As Long)
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT * FROM RV_POLiquidazionePeriodo "
sSQL = sSQL & "WHERE IDRV_POLiquidazionePeriodo=" & Link_Periodo

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    TIPO_IMPORTO_ARTICOLO = 0
    TIPO_IMPORTO_DOCUMENTO = 0
    TIPO_QUANTITA = 0
Else
    TIPO_IMPORTO_ARTICOLO = fnNotNullN(rs!IDTipoImportoArticolo)
    TIPO_IMPORTO_DOCUMENTO = fnNotNullN(rs!IDTipoImportoDocumento)
    TIPO_QUANTITA = fnNotNullN(rs!IDTipoQuantita)
End If



rs.CloseResultset
Set rs = Nothing
End Sub
Private Sub fncCalcoloTrattenute()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim rsTratt As DmtOleDbLib.adoResultset
Dim Quantita_da_prendere_in_considerazione As Double

Select Case TIPO_QUANTITA
    
    Case 1
        Quantita_da_prendere_in_considerazione = Me.txtQtaLavorata.Value
    Case 2
        Quantita_da_prendere_in_considerazione = Me.txtQtaLavorata.Value + Me.txtQtaQuadLav.Value
    Case 3
        Quantita_da_prendere_in_considerazione = Me.txtQtaConferita.Text
    Case 4
        Quantita_da_prendere_in_considerazione = Me.txtQtaVenduta.Value
End Select





sSQL = "SELECT RV_POCalcoloLiqTesta.IDRV_POCalcoloLiqTesta, RV_POCalcoloLiqTesta.IDFiliale, RV_POCalcoloLiqTesta.IDSocio, "
sSQL = sSQL & "RV_POCalcoloLiqTesta.AddebitoImballo, RV_POCalcoloLiqTesta.IDListinoImballo, RV_POCalcoloLiqTesta.IDTipoImportoArticolo,"
sSQL = sSQL & "RV_POCalcoloLiqTesta.IDTipoImportoDocumento, RV_POCalcoloLiqTesta.IDTipoQuantita, RV_POCalcoloLiqTesta.ArticoliDiQuadratura,"
sSQL = sSQL & "RV_POCalcoloLiqTesta.IDTipoLiquidazione, RV_POCalcoloLiqRighe.IDRV_POTipoTrattenuta, RV_POTipoTrattenuta.Tipo1, RV_POTipoTrattenuta.Tipo2,"
sSQL = sSQL & "RV_POTipoTrattenuta.Tipo3 , RV_POTipoTrattenuta.Tipo4 "
sSQL = sSQL & "FROM RV_POTipoTrattenuta RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POCalcoloLiqRighe ON RV_POTipoTrattenuta.IDRV_POTipoTrattenuta = RV_POCalcoloLiqRighe.IDRV_POTipoTrattenuta RIGHT OUTER JOIN "
sSQL = sSQL & "RV_POCalcoloLiqTesta ON RV_POCalcoloLiqRighe.IDRV_POCalcoloLiqTesta = RV_POCalcoloLiqTesta.IDRV_POCalcoloLiqTesta "
sSQL = sSQL & "WHERE RV_POCalcoloLiqTesta.IDFiliale=" & m_App.Branch


Set rs = Cn.OpenResultset(sSQL)
Valore_TrattenutaPerLavorazione = 0
Valore_TrattenutaGenerale = 0
While Not rs.EOF
    sSQL = GetSQL(fnNotNullN(rs!Tipo1), fnNotNullN(rs!Tipo2), fnNotNullN(rs!Tipo3), fnNotNullN(rs!Tipo4), fnNotNullN(m_Document("IDAnagrafica").Value), Link_CategoriaMerceologica, fnNotNullN(m_DocumentsLink("IDArticolo").Value), fnNotNullN(m_DocumentsLink("IDTipoLavorazione").Value))
    If Len(sSQL) > 0 Then
        Set rsTratt = Cn.OpenResultset(sSQL)
            
            If rsTratt.EOF = False Then
                'Trattenuta a valore
                    If fnNotNullN(rs!Tipo4) = 1 Then
                        'X Lavorazione
                        Valore_TrattenutaPerLavorazione = Valore_TrattenutaPerLavorazione + (Quantita_da_prendere_in_considerazione * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        Valore_TrattenutaPerLavorazione = Valore_TrattenutaPerLavorazione + (Quantita_da_prendere_in_considerazione * fnNotNullN(rsTratt!ValoreTrattenuta2))
                    Else
                        'Generale
                        Valore_TrattenutaGenerale = Valore_TrattenutaGenerale + (Quantita_da_prendere_in_considerazione * fnNotNullN(rsTratt!ValoreTrattenuta1))
                        Valore_TrattenutaGenerale = Valore_TrattenutaGenerale + (Quantita_da_prendere_in_considerazione * fnNotNullN(rsTratt!ValoreTrattenuta2))
                    End If
                'Trattenuta a percentuale
                    If fnNotNullN(rs!Tipo4) = 1 Then
                        'X lavorazione
                        Valore_TrattenutaPerLavorazione = Valore_TrattenutaPerLavorazione + ((Me.txtImportoTotaleVenduto / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                        Valore_TrattenutaPerLavorazione = Valore_TrattenutaPerLavorazione + ((Me.txtImportoTotaleVenduto / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                    Else
                        'Generale
                        Valore_TrattenutaGenerale = Valore_TrattenutaGenerale + ((Me.txtImportoTotaleVenduto.Value / 100) * fnNotNullN(rsTratt!PercTrattenuta1))
                        Valore_TrattenutaGenerale = Valore_TrattenutaGenerale + ((Me.txtImportoTotaleVenduto.Value / 100) * fnNotNullN(rsTratt!PercTrattenuta2))
                    End If
            End If
            
        rsTratt.CloseResultset
        Set rsTratt = Nothing
    End If
        
rs.MoveNext
Wend
    'Totali trattenute
    Me.txtTrattenutaPerLavorazione.Value = Valore_TrattenutaPerLavorazione
    Me.txtTrattenuteGenerali.Value = Valore_TrattenutaGenerale
    Me.txtTotaleTrattenuta.Value = Me.txtTrattenutaPerLavorazione.Value + Me.txtTrattenuteGenerali.Value
    
    
rs.CloseResultset
Set rs = Nothing
End Sub

Private Function GetSQL(Tipo1 As Integer, Tipo2 As Integer, Tipo3 As Integer, Tipo4 As Integer, ValueTipo1 As Long, ValueTipo2 As Long, ValueTipo3 As Long, ValueTipo4 As Long) As String
Dim sSQL As String
GetSQL = "SELECT * FROM RV_POTrattenutaPerLiquidazione WHERE "
GetSQL = GetSQL & "IDAzienda=" & m_App.IDFirm & " AND "
GetSQL = GetSQL & "IDFiliale=" & m_App.Branch

sSQL = ""
If Tipo1 = 1 Then
    If ValueTipo1 > 0 Then
        sSQL = sSQL & " AND IDSocio=" & ValueTipo1
    Else
        sSQL = sSQL
    End If
End If
If Tipo2 = 1 Then
    If ValueTipo2 > 0 Then
        sSQL = sSQL & " AND IDCategoriaMerceologica=" & ValueTipo2
    Else
        sSQL = sSQL
    End If
End If
If Tipo3 = 1 Then
    If ValueTipo3 > 0 Then
        sSQL = sSQL & " AND IDArticolo=" & ValueTipo3
    Else
        sSQL = sSQL
    End If
End If
If Tipo4 = 1 Then
    If ValueTipo4 > 0 Then
        sSQL = sSQL & " AND IDTipoLavorazione=" & ValueTipo4
    Else
        sSQL = sSQL
    End If
End If

If sSQL <> "" Then
    GetSQL = GetSQL & sSQL
Else
    GetSQL = ""
End If
End Function
Private Function GET_TIPOLAVORAZIONE() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDTipoLavorazione FROM RV_POLavorazione "
sSQL = sSQL & "WHERE IDRV_POLavorazione=" & fnNotNullN(m_DocumentsLink("IDRV_POLavorazione").Value)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TIPOLAVORAZIONE = 0
Else
    GET_TIPOLAVORAZIONE = fnNotNullN(rs!IDTipoLavorazione)
End If

rs.CloseResultset
Set rs = Nothing
End Function

Private Function GET_CATEGORIA_MERCEOLOGICA() As Long
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT IDCategoriaMerceologica FROM Articolo "
sSQL = sSQL & "WHERE IDArticolo=" & fnNotNullN(m_DocumentsLink("IDArticolo").Value)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_CATEGORIA_MERCEOLOGICA = 0
Else
    GET_CATEGORIA_MERCEOLOGICA = fnNotNullN(rs!IDCategoriaMerceologica)
End If

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub RicalcolaTotaleLiquidazione()
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset




sSQL = "SELECT SUM(ImponibileDaReg) as TotaleImponibile, SUM(ImpostaDaReg) as TotaleImposta, SUM(ImportoLordoDaReg) as TotaleDocumento, "
sSQL = sSQL & "SUM(TrattenuteGenerali) as TotaleTrattenuteGenerali, SUM(TrattenutePerLavorazione) as TotaleTrattenutePerLavorazioni,"
sSQL = sSQL & "SUM(TrattenuteTotali) as TrattenuteTotali,  SUM(TotaleTrattenutaConferimento) as TotaleTrattenutaConferimento "
sSQL = sSQL & "FROM RV_POLiquidazioneRigheEla "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & fnNotNullN(m_Document(m_Document.PrimaryKey).Value)

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    Me.txtTotaleDocumento.Value = 0
    Me.txtTotaleIvaDocumento.Value = 0
    Me.txtTotaleLordoDocumento.Value = 0
    Me.txtTotaleTrattenuteConferimento.Value = 0
    Me.txtTotaleTrattenuteGenerali.Value = 0
    Me.txtTotaleTrattenutePerLavorazioni.Value = 0
    Me.txtTotaleTrattenute.Value = 0
    Me.txtTotaleTrattenuteAgg.Value = 0
    
    
    RICALCOLO_NUOVE_TRATTENUTA fnNotNullN(m_Document(m_Document.PrimaryKey).Value), Me.txtTotaleDocumento.Value
    
    
    Me.txtTotaleTrattenuteAgg.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 1)
    Me.txtTotaleTrattAggRiep.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 2)
    
    If TIPO_IMPORTO_DOCUMENTO <= 1 Then
        Me.txtNettoLiquidazione.Value = Me.txtTotaleDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    Else
        Me.txtNettoLiquidazione.Value = Me.txtTotaleLordoDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    End If


Else
    Me.txtTotaleDocumento.Value = fnNotNullN(rs!TotaleImponibile)
    Me.txtTotaleIvaDocumento.Value = fnNotNullN(rs!TotaleImposta)
    Me.txtTotaleLordoDocumento.Value = fnNotNullN(rs!TotaleDocumento)
    Me.txtTotaleTrattenuteConferimento.Value = fnNotNullN(rs!TotaleTrattenutaConferimento)
    Me.txtTotaleTrattenuteGenerali.Value = fnNotNullN(rs!TotaleTrattenuteGenerali)
    Me.txtTotaleTrattenutePerLavorazioni.Value = fnNotNullN(rs!TotaleTrattenutePerLavorazioni)
    'Me.txtTotaleTrattenute.Value = fnNotNullN(rs!TrattenuteTotali)
    Me.txtTotaleTrattenute.Value = Me.txtTotaleTrattenuteConferimento.Value + Me.txtTotaleTrattenuteGenerali.Value + Me.txtTotaleTrattenutePerLavorazioni.Value
    
    RICALCOLO_NUOVE_TRATTENUTA fnNotNullN(m_Document(m_Document.PrimaryKey).Value), Me.txtTotaleDocumento.Value

    
    Me.txtTotaleTrattenuteAgg.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 1)
    Me.txtTotaleTrattAggRiep.Value = GET_TOTALE_TRATTENUTA_AGGIUNTIVA(fnNotNullN(m_Document(m_Document.PrimaryKey)), 2)
    
    If TIPO_IMPORTO_DOCUMENTO <= 1 Then
        Me.txtNettoLiquidazione.Value = Me.txtTotaleDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    Else
        Me.txtNettoLiquidazione.Value = Me.txtTotaleLordoDocumento.Value - Me.txtTotaleTrattenute.Value - Me.txtTotaleTrattenuteAgg.Value - Me.txtTotaleTrattAggRiep.Value
    End If

End If

rs.CloseResultset
Set rs = Nothing
'onsave
DoEvents
End Sub

Private Sub txtTotaleTrattenuteGenerali_Change()
If Not (BrwMain.Visible) Then Change
End Sub

Private Sub txtTotaleTrattenutePerLavorazioni_Change()
If Not (BrwMain.Visible) Then Change
End Sub
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
Private Function fnGetTipoOggetto() As Long
    Dim sSQL As String
    Dim rs As DmtOleDbLib.adoResultset
    
    sSQL = "SELECT TipoOggetto.IDTipoOggetto "
    sSQL = sSQL & "FROM TipoOggetto INNER JOIN "
    sSQL = sSQL & "Gestore ON TipoOggetto.IDGestore = Gestore.IDGestore "
    sSQL = sSQL & "WHERE Gestore.Gestore=" & fnNormString(App.EXEName)
    
    Set rs = Cn.OpenResultset(sSQL)
    If rs.EOF = False Then
        fnGetTipoOggetto = rs!IDTipoOggetto
    Else
        fnGetTipoOggetto = 0
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
Private Function GET_TOTALE_TRATTENUTA_AGGIUNTIVA(IDLiquidazione As Long, IDTipoTrattenutaAggiutiva As Long) As Double
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

sSQL = "SELECT Sum(ImportoTrattenuta) As TotaleTrattenute "
sSQL = sSQL & "FROM RV_POLiquidazioneRighe "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione
sSQL = sSQL & " AND IDRV_POTipoTrattenutaAggiuntiva=" & IDTipoTrattenutaAggiutiva

Set rs = Cn.OpenResultset(sSQL)

If rs.EOF Then
    GET_TOTALE_TRATTENUTA_AGGIUNTIVA = 0
Else
    GET_TOTALE_TRATTENUTA_AGGIUNTIVA = FormatNumber(fnNotNullN(rs!TotaleTrattenute), 2)
End If



rs.CloseResultset
Set rs = Nothing
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
Private Function GET_CONTROLLO_RIGHE_TRATTENUTE(IDLiquidazione As Long) As Boolean
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset

GET_CONTROLLO_RIGHE_TRATTENUTE = True

sSQL = "SELECT IDRV_POLiquidazioneRighe, IDRV_POLiquidazione, IDRV_POTipoRicalcoloComm "
sSQL = sSQL & "FROM RV_POLiquidazioneRighe "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = Cn.OpenResultset(sSQL)

While Not rs.EOF
    If GET_CONTROLLO_RIGHE_TRATTENUTE = True Then
        If fnNotNullN(rs!IDRV_POTipoRicalcoloComm) = 0 Then
            GET_CONTROLLO_RIGHE_TRATTENUTE = False
        End If
    End If
rs.MoveNext
Wend

rs.CloseResultset
Set rs = Nothing
End Function
Private Sub RICALCOLO_NUOVE_TRATTENUTA(IDLiquidazione As Long, TotaleDocumento As Double)
Dim sSQL As String
Dim rs As ADODB.Recordset



sSQL = "SELECT * FROM RV_POLiquidazioneRighe "
sSQL = sSQL & "WHERE IDRV_POLiquidazione=" & IDLiquidazione

Set rs = New ADODB.Recordset

rs.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic


While Not rs.EOF
    If TotaleDocumento <> 0 Then
        If fnNotNullN(rs!IDRV_POTipoRicalcoloComm) = 1 Then
            rs!ImportoTrattenuta = (TotaleDocumento / 100) * fnNotNullN(rs!Percentuale)
        End If
        
        If fnNotNullN(rs!IDRV_POTipoRicalcoloComm) = 2 Then
            rs!Percentuale = (fnNotNullN(rs!ImportoTrattenuta) / TotaleDocumento) * 100
        End If
    
        rs.Update
    End If

rs.MoveNext
Wend

rs.Close
Set rs = Nothing

m_DocumentsLink1.Refresh

End Sub
Private Sub GET_OGGETTI_PER_STAMPE()
On Error GoTo ERR_GET_OGGETTI_PER_STAMPE
Dim sSQL As String
Dim rs As DmtOleDbLib.adoResultset
Dim sSQL_WHERE As String
Dim Field As DmtDocManLib.Field
Dim Cond As dmtgridctl.dgCondition
Dim OLD_CURSOR As Long
Dim rsNew As ADODB.Recordset

sSQL = "DELETE FROM RV_POTMPLiquidazionePerStampa "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser
Cn.Execute sSQL



sSQL_WHERE = "WHERE IDFiliale=" & TheApp.Branch
sSQL_WHERE = sSQL_WHERE & " AND IDUtente=" & TheApp.IDUser

For Each Cond In BrwMain.Conditions

    Select Case Cond.ConditionType
        
        'Condizione boolean
        Case dgCondTypeBoolean
            Select Case Cond.FromValue
                Case "SI"
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormBoolean(1)
                Case "NO"
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormBoolean(0)
            End Select
            'm_DocType.Fields(Cond.FieldName).Value = IIf(IsEmpty(Cond.FromValue), Empty, Abs(CDbl(Cond.FromValue = "SI")))
            
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
                sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & ">=" & fnNormNumber(Cond.FromValue)
                sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "<=" & fnNormNumber(Cond.ToValue)
            Else
                If fnNotNullN(Cond.FromValue) > 0 Then
                    sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "=" & fnNormNumber(Cond.FromValue)
                End If
            End If
        
        Case dgCondTypeDate
            If Cond.RangeChecked = True Then
                sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & ">=" & fnNormDate(Cond.FromValue)
                sSQL_WHERE = sSQL_WHERE & " AND " & Cond.FieldName & "<=" & fnNormDate(Cond.ToValue)
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

sSQL = "SELECT IDRV_POLiquidazione "
sSQL = sSQL & "FROM RV_PORepLiquidazione "
sSQL = sSQL & sSQL_WHERE

Set rs = Cn.OpenResultset(sSQL)


sSQL = "SELECT * FROM RV_POTMPLiquidazionePerStampa "
sSQL = sSQL & "WHERE IDUtente=" & TheApp.IDUser

Set rsNew = New ADODB.Recordset
rsNew.Open sSQL, Cn.InternalConnection, adOpenKeyset, adLockPessimistic

While Not rs.EOF
    rsNew.AddNew
        rsNew!IDRV_POLiquidazione = fnNotNullN(rs!IDRV_POLiquidazione)
        rsNew!IDUtente = TheApp.IDUser
    rsNew.Update
rs.MoveNext
Wend

rsNew.Close
Set rsNew = Nothing

rs.CloseResultset
Set rs = Nothing
Exit Sub
ERR_GET_OGGETTI_PER_STAMPE:
    MsgBox Err.Description, vbCritical, "GET_OGGETTI_PER_STAMPE"
End Sub

